# -*- coding: utf-8 -*-
import sys, os, json, threading, subprocess, shutil, tempfile
import urllib.request
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame,
    QLabel, QPushButton, QHBoxLayout, QVBoxLayout,
    QGraphicsDropShadowEffect, QMessageBox, QDialog,
    QTextEdit, QProgressBar
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt6.QtGui import QColor, QCursor

# ============================================================
# 配置
# ============================================================
GITHUB_USER  = "ncasd"
GITHUB_REPO  = "amazon-tools"
GITHUB_TOKEN = "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"
LOCAL_VERSION_FILE = "version.json"
SCRIPTS = {
    "pricing":  "定价计算器.py",
    "splitter": "定制尺寸加价.py",
}
# 当前本地版本
LOCAL_VERSION = "1.0.0"
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def github_raw_url(filename):
    return (f"https://raw.githubusercontent.com/"
            f"{GITHUB_USER}/{GITHUB_REPO}/main/{filename}")

def github_api_url(path):
    return f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/{path}"

def gh_request(url, timeout=10):
    req = urllib.request.Request(url, headers={
        "Authorization": f"token {GITHUB_TOKEN}",
        "User-Agent": "AmazonTools/1.0",
        "Accept": "application/vnd.github.v3.raw",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8")

def fetch_remote_version():
    raw = gh_request(github_raw_url("version.json"))
    return json.loads(raw)

def download_file(filename, dest_path):
    raw = gh_request(github_raw_url(filename), timeout=30)
    with open(dest_path, "w", encoding="utf-8") as f:
        f.write(raw)

# ============================================================
# 更新线程
# ============================================================

class UpdateChecker(QThread):
    result = pyqtSignal(dict)   # {"has_update": bool, "remote": dict, "error": str}

    def run(self):
        try:
            remote = fetch_remote_version()
            local  = LOCAL_VERSION
            has_update = remote.get("version", "0") != local
            self.result.emit({"has_update": has_update, "remote": remote, "error": ""})
        except Exception as e:
            self.result.emit({"has_update": False, "remote": {}, "error": str(e)})


class Downloader(QThread):
    progress = pyqtSignal(int, str)   # (percent, message)
    done     = pyqtSignal(bool, str)  # (success, message)

    def __init__(self, files):
        super().__init__()
        self.files = files  # list of filename

    def run(self):
        total = len(self.files)
        try:
            for i, filename in enumerate(self.files):
                self.progress.emit(int((i / total) * 90), f"下载 {filename}…")
                dest = os.path.join(BASE_DIR, filename)
                # 先下载到临时文件
                tmp = dest + ".tmp"
                download_file(filename, tmp)
                shutil.move(tmp, dest)
            self.progress.emit(100, "完成")
            self.done.emit(True, "更新成功，重启后生效")
        except Exception as e:
            self.done.emit(False, str(e))

# ============================================================
# 更新对话框
# ============================================================

class UpdateDialog(QDialog):
    def __init__(self, remote_info, parent=None):
        super().__init__(parent)
        self.setWindowTitle("发现新版本")
        self.setFixedSize(400, 280)
        self.setStyleSheet("""
            QDialog { background: #F5F5F5; }
            QLabel { font-family:'Segoe UI'; color:#111111; background:transparent; }
            QPushButton {
                font-family:'Segoe UI'; font-size:12px; font-weight:600;
                border-radius:7px; padding:8px 20px;
            }
        """)
        self.remote_info = remote_info
        self.downloader  = None
        self._build(remote_info)

    def _build(self, info):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # 标题
        title = QLabel(f"🆕  发现新版本  v{info.get('version','?')}")
        title.setStyleSheet("font-size:15px; font-weight:700;")
        layout.addWidget(title)

        # 更新日志
        log_lbl = QLabel("更新内容：")
        log_lbl.setStyleSheet("font-size:11px; color:#555555;")
        layout.addWidget(log_lbl)

        log_box = QTextEdit()
        log_box.setReadOnly(True)
        log_box.setFixedHeight(80)
        log_box.setText(info.get("changelog", "无说明"))
        log_box.setStyleSheet("""
            QTextEdit {
                font-family:'Segoe UI'; font-size:11px;
                background:#EFEFEF; border:1px solid #CCCCCC;
                border-radius:6px; padding:6px; color:#333333;
            }
        """)
        layout.addWidget(log_box)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar { background:#DDDDDD; border:none; border-radius:3px; }
            QProgressBar::chunk {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #0067C0, stop:1 #40A0E0);
                border-radius:3px;
            }
        """)
        layout.addWidget(self.progress_bar)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("font-size:11px; color:#0067C0;")
        layout.addWidget(self.status_lbl)

        layout.addStretch()

        # 按钮行
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        self.update_btn = QPushButton("立即更新")
        self.update_btn.setStyleSheet("""
            QPushButton { color:white; background:#0067C0; border:none; }
            QPushButton:hover { background:#1478CC; }
            QPushButton:disabled { background:#A0C4E8; }
        """)
        self.update_btn.clicked.connect(self._start_update)
        skip_btn = QPushButton("跳过")
        skip_btn.setStyleSheet("""
            QPushButton { color:#333333; background:#E0E0E0; border:1px solid #CCCCCC; }
            QPushButton:hover { background:#D0D0D0; }
        """)
        skip_btn.clicked.connect(self.reject)
        btn_row.addStretch()
        btn_row.addWidget(skip_btn)
        btn_row.addWidget(self.update_btn)
        layout.addLayout(btn_row)

    def _start_update(self):
        self.update_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        files = list(self.remote_info.get("files", {}).keys())
        self.downloader = Downloader(files)
        self.downloader.progress.connect(self._on_progress)
        self.downloader.done.connect(self._on_done)
        self.downloader.start()

    def _on_progress(self, pct, msg):
        self.progress_bar.setValue(pct)
        self.status_lbl.setText(msg)

    def _on_done(self, success, msg):
        if success:
            self.status_lbl.setText("✅ " + msg)
            QTimer.singleShot(1200, lambda: self._restart())
        else:
            self.status_lbl.setText("❌ " + msg)
            self.update_btn.setEnabled(True)

    def _restart(self):
        self.accept()
        subprocess.Popen([sys.executable, os.path.abspath(__file__)])
        QApplication.quit()

# ============================================================
# 样式
# ============================================================

STYLESHEET = """
QMainWindow, QWidget#central {
    background: #DCDCDE;
}
QFrame#titlebar {
    background: #FAFAFA;
    border-bottom: 1px solid #C8C8C8;
}
QLabel#title {
    font-family: 'Segoe UI'; font-size:17px; font-weight:700; color:#111111;
}
QLabel#subtitle {
    font-family: 'Segoe UI'; font-size:11px; color:#777777;
}
QLabel#version_lbl {
    font-family: 'Segoe UI'; font-size:10px; color:#AAAAAA;
}
"""

def make_shadow(blur=24, dy=3, alpha=22):
    s = QGraphicsDropShadowEffect()
    s.setBlurRadius(blur); s.setOffset(0, dy)
    s.setColor(QColor(0, 0, 0, alpha))
    return s

# ============================================================
# 应用卡片
# ============================================================

class AppCard(QFrame):
    def __init__(self, icon, name, desc, script_key):
        super().__init__()
        self.script_key = script_key
        self.setFixedSize(260, 210)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setGraphicsEffect(make_shadow())
        self._ns = "QFrame{background:#FAFAFA;border-radius:14px;border:1px solid #D0D0D0;}"
        self._hs = "QFrame{background:#FFFFFF;border-radius:14px;border:1.5px solid #0067C0;}"
        self._ps = "QFrame{background:#F0F0F0;border-radius:14px;border:1.5px solid #005294;}"
        self.setStyleSheet(self._ns)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 18)
        layout.setSpacing(8)

        icon_lbl = QLabel(icon)
        icon_lbl.setStyleSheet("font-size:34px; background:transparent; border:none;")

        name_lbl = QLabel(name)
        name_lbl.setStyleSheet(
            "font-family:'Segoe UI'; font-size:14px; font-weight:700; "
            "color:#111111; background:transparent; border:none;")

        desc_lbl = QLabel(desc)
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet(
            "font-family:'Segoe UI'; font-size:11px; color:#666666; "
            "background:transparent; border:none;")

        layout.addWidget(icon_lbl)
        layout.addWidget(name_lbl)
        layout.addWidget(desc_lbl)
        layout.addStretch()

        self.launch_btn = QPushButton("启  动")
        self.launch_btn.setFixedHeight(34)
        self.launch_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.launch_btn.setStyleSheet("""
            QPushButton {
                font-family:'Segoe UI'; font-size:12px; font-weight:700;
                color:white; background:#0067C0; border:none; border-radius:7px;
            }
            QPushButton:hover   { background:#1478CC; }
            QPushButton:pressed { background:#005294; }
        """)
        self.launch_btn.clicked.connect(self._launch)
        layout.addWidget(self.launch_btn)

        self.status_lbl = QLabel("")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet(
            "font-family:'Segoe UI'; font-size:10px; color:#0067C0; "
            "background:transparent; border:none;")
        layout.addWidget(self.status_lbl)

    def _launch(self):
        script = os.path.join(BASE_DIR, SCRIPTS[self.script_key])
        if not os.path.exists(script):
            self.status_lbl.setText(f"❌ 找不到文件")
            return
        try:
            subprocess.Popen([sys.executable, script],
                             creationflags=subprocess.CREATE_NO_WINDOW
                             if sys.platform=="win32" else 0)
            self.status_lbl.setText("✅ 已启动")
            QTimer.singleShot(3000, lambda: self.status_lbl.setText(""))
        except Exception as e:
            self.status_lbl.setText(f"❌ {e}")

    def enterEvent(self, e):
        self.setStyleSheet(self._hs)
        self.setGraphicsEffect(make_shadow(32, 6, 30))

    def leaveEvent(self, e):
        self.setStyleSheet(self._ns)
        self.setGraphicsEffect(make_shadow())

    def mousePressEvent(self, e):   self.setStyleSheet(self._ps)
    def mouseReleaseEvent(self, e): self.setStyleSheet(self._hs)

# ============================================================
# 主窗口
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Amazon 工具箱")
        self.setFixedSize(660, 380)
        self._build()
        # 启动后2秒静默检查更新
        QTimer.singleShot(2000, self._check_update)

    def _build(self):
        central = QWidget(); central.setObjectName("central")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # 标题栏
        tb = QFrame(); tb.setObjectName("titlebar"); tb.setFixedHeight(56)
        tbl = QHBoxLayout(tb); tbl.setContentsMargins(22,0,22,0)
        icon = QLabel("🛍️"); icon.setStyleSheet("font-size:20px;")
        title = QLabel("Amazon 卖家工具箱"); title.setObjectName("title")
        tbl.addWidget(icon); tbl.addSpacing(8); tbl.addWidget(title)
        tbl.addStretch()

        # 更新状态
        self.update_lbl = QLabel("检查更新中…")
        self.update_lbl.setObjectName("version_lbl")
        tbl.addWidget(self.update_lbl)
        tbl.addSpacing(12)

        # 手动检查更新按钮
        upd_btn = QPushButton("🔄 检查更新")
        upd_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        upd_btn.setStyleSheet("""
            QPushButton {
                font-family:'Segoe UI'; font-size:11px; font-weight:600;
                color:#0067C0; background:#DCF0FF;
                border:none; border-radius:6px; padding:5px 12px;
            }
            QPushButton:hover { background:#CCE8FF; }
        """)
        upd_btn.clicked.connect(self._check_update_manual)
        tbl.addWidget(upd_btn)

        # 版本号
        ver_lbl = QLabel(f"v{LOCAL_VERSION}")
        ver_lbl.setObjectName("version_lbl")
        tbl.addSpacing(10); tbl.addWidget(ver_lbl)

        root.addWidget(tb)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#C8C8C8; background:#C8C8C8; max-height:1px;")
        root.addWidget(sep)

        # 卡片区
        body = QWidget(); body.setStyleSheet("background:#DCDCDE;")
        bl = QHBoxLayout(body)
        bl.setContentsMargins(50, 35, 50, 35); bl.setSpacing(30)
        bl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        for icon, name, desc, key in [
            ("🛒", "定价计算器",
             "输入产品成本、尺寸、重量，\n自动计算运费与前台售价。",
             "pricing"),
            ("📦", "定制尺寸加价拆分",
             "批量读取加价表，拆分各维度\n增量，输出 Amazon 定制选项价格。",
             "splitter"),
        ]:
            bl.addWidget(AppCard(icon, name, desc, key))

        root.addWidget(body)

    def _check_update(self):
        self.checker = UpdateChecker()
        self.checker.result.connect(self._on_check_result)
        self.checker.start()

    def _check_update_manual(self):
        self.update_lbl.setText("检查中…")
        self._check_update()

    def _on_check_result(self, res):
        if res["error"]:
            self.update_lbl.setText("⚠ 网络不可用")
            return
        if res["has_update"]:
            v = res["remote"].get("version","?")
            self.update_lbl.setText(f"🆕 有新版本 v{v}")
            dlg = UpdateDialog(res["remote"], self)
            dlg.exec()
        else:
            self.update_lbl.setText("✅ 已是最新版")
            QTimer.singleShot(4000, lambda: self.update_lbl.setText(f"v{LOCAL_VERSION}"))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    app.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())