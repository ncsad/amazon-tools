# -*- coding: utf-8 -*-
import sys, os, subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame,
    QLabel, QPushButton, QHBoxLayout, QVBoxLayout,
    QGraphicsDropShadowEffect
)
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve, QRect, QTimer
from PyQt6.QtGui import QColor, QCursor

# ── 同目录下的脚本名 ──────────────────────────────────────────
SCRIPTS = {
    "pricing":  "定价计算器.py",
    "splitter": "定制尺寸加价.py",
}

STYLESHEET = """
QMainWindow, QWidget#central {
    background: #DCDCDE;
}
QFrame#titlebar {
    background: #FAFAFA;
    border-bottom: 1px solid #C8C8C8;
}
QLabel#title {
    font-family: 'Segoe UI';
    font-size: 17px;
    font-weight: 700;
    color: #111111;
}
QLabel#subtitle {
    font-family: 'Segoe UI';
    font-size: 11px;
    color: #777777;
}
"""

def make_shadow(blur=24, dy=3, alpha=22):
    s = QGraphicsDropShadowEffect()
    s.setBlurRadius(blur)
    s.setOffset(0, dy)
    s.setColor(QColor(0, 0, 0, alpha))
    return s


class AppCard(QFrame):
    """单个应用卡片"""
    def __init__(self, icon, name, desc, script_key, parent=None):
        super().__init__(parent)
        self.script_key = script_key
        self.setFixedSize(260, 200)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setGraphicsEffect(make_shadow())
        self._normal_style = """
            QFrame {
                background: #FAFAFA;
                border-radius: 14px;
                border: 1px solid #D0D0D0;
            }
        """
        self._hover_style = """
            QFrame {
                background: #FFFFFF;
                border-radius: 14px;
                border: 1.5px solid #0067C0;
            }
        """
        self._press_style = """
            QFrame {
                background: #F0F0F0;
                border-radius: 14px;
                border: 1.5px solid #005294;
            }
        """
        self.setStyleSheet(self._normal_style)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 20)
        layout.setSpacing(10)

        # 图标
        icon_lbl = QLabel(icon)
        icon_lbl.setStyleSheet(
            "font-size: 36px; background: transparent; border: none;")
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft)

        # 名称
        name_lbl = QLabel(name)
        name_lbl.setStyleSheet(
            "font-family:'Segoe UI'; font-size:15px; font-weight:700; "
            "color:#111111; background:transparent; border:none;")

        # 描述
        desc_lbl = QLabel(desc)
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet(
            "font-family:'Segoe UI'; font-size:11px; color:#666666; "
            "background:transparent; border:none;")

        layout.addWidget(icon_lbl)
        layout.addWidget(name_lbl)
        layout.addWidget(desc_lbl)
        layout.addStretch()

        # 启动按钮
        self.launch_btn = QPushButton("启  动")
        self.launch_btn.setFixedHeight(34)
        self.launch_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.launch_btn.setStyleSheet("""
            QPushButton {
                font-family:'Segoe UI'; font-size:12px; font-weight:700;
                color: white; background: #0067C0;
                border: none; border-radius: 7px;
            }
            QPushButton:hover   { background: #1478CC; }
            QPushButton:pressed { background: #005294; }
        """)
        self.launch_btn.clicked.connect(self._launch)
        layout.addWidget(self.launch_btn)

        # 状态标签
        self.status_lbl = QLabel("")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet(
            "font-family:'Segoe UI'; font-size:10px; color:#0067C0; "
            "background:transparent; border:none;")
        layout.addWidget(self.status_lbl)

    def _launch(self):
        base  = os.path.dirname(os.path.abspath(__file__))
        script = os.path.join(base, SCRIPTS[self.script_key])
        if not os.path.exists(script):
            self.status_lbl.setText(f"❌ 找不到 {SCRIPTS[self.script_key]}")
            return
        try:
            subprocess.Popen([sys.executable, script],
                             creationflags=subprocess.CREATE_NO_WINDOW
                             if sys.platform == "win32" else 0)
            self.status_lbl.setText("✅ 已启动")
            QTimer.singleShot(3000, lambda: self.status_lbl.setText(""))
        except Exception as e:
            self.status_lbl.setText(f"❌ {e}")

    def enterEvent(self, e):
        self.setStyleSheet(self._hover_style)
        shadow = make_shadow(blur=32, dy=6, alpha=30)
        self.setGraphicsEffect(shadow)

    def leaveEvent(self, e):
        self.setStyleSheet(self._normal_style)
        self.setGraphicsEffect(make_shadow())

    def mousePressEvent(self, e):
        self.setStyleSheet(self._press_style)

    def mouseReleaseEvent(self, e):
        self.setStyleSheet(self._hover_style)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Amazon 工具箱")
        self.setFixedSize(640, 340)
        self._build()

    def _build(self):
        central = QWidget(); central.setObjectName("central")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # 标题栏
        tb = QFrame(); tb.setObjectName("titlebar"); tb.setFixedHeight(56)
        tbl = QHBoxLayout(tb); tbl.setContentsMargins(22, 0, 22, 0)
        icon = QLabel("🛍️"); icon.setStyleSheet("font-size:20px;")
        title = QLabel("Amazon 卖家工具箱"); title.setObjectName("title")
        sub = QLabel("选择工具启动"); sub.setObjectName("subtitle")
        tbl.addWidget(icon); tbl.addSpacing(8)
        tbl.addWidget(title); tbl.addSpacing(10)
        tbl.addWidget(sub); tbl.addStretch()
        root.addWidget(tb)

        # 分割线
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#C8C8C8; background:#C8C8C8; max-height:1px;")
        root.addWidget(sep)

        # 卡片区
        body = QWidget(); body.setStyleSheet("background:#DCDCDE;")
        bl = QHBoxLayout(body)
        bl.setContentsMargins(40, 30, 40, 30)
        bl.setSpacing(28)
        bl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        cards = [
            AppCard("🛒", "定价计算器",
                    "输入产品成本、尺寸、重量，\n自动计算运费与前台售价。",
                    "pricing"),
            AppCard("📦", "定制尺寸加价拆分",
                    "批量读取加价表，拆分各维度\n增量，输出 Amazon 定制选项价格。",
                    "splitter"),
        ]
        for card in cards:
            bl.addWidget(card)

        root.addWidget(body)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    app.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())