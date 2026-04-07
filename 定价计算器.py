# -*- coding: utf-8 -*-
import math, json, threading
import urllib.request
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QLineEdit, QHBoxLayout, QVBoxLayout, QGridLayout, QSizePolicy,
    QGraphicsDropShadowEffect, QMessageBox, QMenu
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QColor, QFont

# ============================================================
# 计算逻辑
# ============================================================

def get_all_rates():
    try:
        url = "https://api.exchangerate-api.com/v4/latest/USD"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as r:
            rates = json.loads(r.read()).get("rates", {})
            usd_cny = rates.get("CNY", 7.2)
            gbp_usd = rates.get("GBP", 0.78)
            return {"USD": usd_cny, "GBP": usd_cny / gbp_usd}
    except:
        return {"USD": 7.2, "GBP": 9.0}

def calc_vol_weight(l, w, h): return l * w * h / 6000
def calc_billing(actual, l, w, h): return max(actual, calc_vol_weight(l, w, h))

def check_oversize_us(l, w, h):
    d = sorted([l,w,h], reverse=True)
    girth = (d[1]+d[2])*2 + d[0]
    combined = d[0]+2*d[1]+2*d[2]
    if d[0] > 273: return 3
    if girth >= 330 or d[0] >= 243: return 2
    if d[0] >= 121 or d[1] >= 76 or combined >= 266: return 1
    return 0

def shipping_us(billing_w, l, w, h, actual_kg):
    d = sorted([l,w,h], reverse=True)
    girth = (d[1]+d[2])*2 + d[0]
    combined = d[0]+2*d[1]+2*d[2]
    level = check_oversize_us(l, w, h)
    c_parcel = 27 + billing_w * 100
    c_small = (80 + math.ceil(billing_w)*30) if billing_w <= 6 else None
    c_large = 290 if billing_w <= 10 else 290 + math.ceil(billing_w-10)*27
    sur_large = 0; warn = []
    if actual_kg > 68: warn.append("⚠ 实重>68KG 大包拒收")
    if d[0] > 273:   sur_large += 5000; warn.append("⚠ 大包超规+5000元")
    elif girth>=330 or d[0]>=243: sur_large += 600; warn.append("⚠ 大包超规+600元")
    elif d[0]>=121 or d[1]>=76 or combined>=266: sur_large += 150; warn.append("⚠ 大包超规+150元")
    if actual_kg > 30:   sur_large += 350; warn.append("⚠ 超重+350元")
    elif actual_kg > 22: sur_large += 160; warn.append("⚠ 超重+160元")
    candidates = []
    if level == 0 and billing_w <= 0.83:
        candidates.append(("小包裹", c_parcel, 0, []))
    if c_small and level < 2:
        sur_s = 120 if level==1 else 0
        w_s = ["⚠ 中小件超规+120元"] if level==1 else []
        candidates.append(("经济中小件", c_small+sur_s, sur_s, w_s))
    candidates.append(("大包经济线", c_large+sur_large, sur_large, warn))
    best = min(candidates, key=lambda x: x[1])
    return best[0], best[1], best[1]-best[2], best[2], best[3]

def shipping_uk(billing_w, l, w, h):
    base = 25 + billing_w*50 if billing_w <= 3 else 57*billing_w + 30
    d = sorted([l,w,h], reverse=True)
    girth = (d[1]+d[2])*2 + d[0]
    length_sur = 0
    for threshold, fee in [(180,200),(170,180),(160,150),(150,100),(140,70),(120,32)]:
        if d[0] >= threshold: length_sur = fee; break
    girth_sur = 150 if girth >= 266 else 0
    sur = max(length_sur, girth_sur)
    warn = [f"⚠ 英国超规+{sur}元"] if sur > 0 else []
    return "英国经济线", base+sur, base, sur, warn

def calc_price(cost, shipping, profit_rate, fx_rate, coeff, zi=10, zs=10):
    E2 = cost + shipping
    D2 = E2 / (1 - profit_rate)
    C2 = D2 / fx_rate / 0.7
    markup = C2 * fx_rate / coeff - zi - zs
    return {"E2": round(E2,2), "sell": round(C2,2),
            "markup": round(markup,2), "profit": round(profit_rate*100,2)}

# ============================================================
# 样式表
# ============================================================

STYLESHEET = """
QMainWindow, QWidget#central {
    background: #DCDCDE;
}
QFrame#card {
    background: #FAFAFA;
    border-radius: 12px;
    border: 1px solid #C4C4C4;
}
QFrame#titlebar {
    background: #FAFAFA;
    border-bottom: 1px solid #C8C8C8;
}
QLabel#title {
    font-family: 'Segoe UI';
    font-size: 16px;
    font-weight: 700;
    color: #111111;
}
QLabel#section {
    font-family: 'Segoe UI';
    font-size: 10px;
    color: #666666;
    font-weight: 700;
    letter-spacing: 1px;
}
QLabel#field_label {
    font-family: 'Segoe UI';
    font-size: 12px;
    font-weight: 500;
    color: #222222;
}
QLabel#result_label {
    font-family: 'Segoe UI';
    font-size: 12px;
    font-weight: 500;
    color: #333333;
}
QLabel#rate_label {
    font-family: 'Segoe UI';
    font-size: 11px;
    color: #666666;
}
QLabel#warn_label {
    font-family: 'Segoe UI';
    font-size: 11px;
    font-weight: 500;
    color: #B02020;
}
QLabel#channel_label {
    font-family: 'Segoe UI';
    font-size: 12px;
    font-weight: 500;
    color: #333333;
}
QLineEdit {
    font-family: 'Segoe UI';
    font-size: 13px;
    font-weight: 500;
    color: #111111;
    background: #F2F2F2;
    border: 1.5px solid #BBBBBB;
    border-radius: 6px;
    padding: 5px 10px;
}
QLineEdit:focus {
    border: 1.5px solid #0067C0;
    background: white;
}

QPushButton#mode_btn {
    font-family: 'Segoe UI';
    font-size: 12px;
    font-weight: 500;
    color: #444444;
    background: transparent;
    border: none;
    border-radius: 6px;
    padding: 5px 14px;
    min-height: 30px;
}
QPushButton#mode_btn:hover { background: #D8D8D8; }
QPushButton#mode_btn_active {
    font-family: 'Segoe UI';
    font-size: 12px;
    font-weight: 700;
    color: #0055A8;
    background: #CCE8FF;
    border: none;
    border-radius: 6px;
    padding: 5px 14px;
    min-height: 30px;
}
"""

# ============================================================
# 工具组件
# ============================================================

def make_card():
    f = QFrame()
    f.setObjectName("card")
    shadow = QGraphicsDropShadowEffect()
    shadow.setBlurRadius(20)
    shadow.setOffset(0, 2)
    shadow.setColor(QColor(0, 0, 0, 18))
    f.setGraphicsEffect(shadow)
    return f

def make_separator():
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setStyleSheet("color: #E8E8E8; background: #E8E8E8; max-height: 1px;")
    return line

class FieldRow(QWidget):
    """输入行：标签 + 输入框 + 单位"""
    def __init__(self, label, default="", unit="", parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.setSpacing(8)

        lbl = QLabel(label)
        lbl.setObjectName("field_label")
        lbl.setFixedWidth(58)
        lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.entry = QLineEdit(default)
        self.entry.setFixedWidth(100)

        layout.addWidget(lbl)
        layout.addWidget(self.entry)

        if unit:
            u = QLabel(unit)
            u.setStyleSheet("font-family:'Segoe UI'; font-size:11px; color:#AAAAAA; background:transparent;")
            layout.addWidget(u)
        layout.addStretch()

    def text(self): return self.entry.text()
    def clear(self, default=""): self.entry.setText(default)


class CopyableEdit(QLineEdit):
    """只读输入框，右键菜单只有复制"""
    def __init__(self, color="#1A1A1A", size="13px", weight="600"):
        super().__init__("—")
        self.setReadOnly(True)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_menu)
        self.setStyleSheet(f"""
            QLineEdit {{
                font-family: 'Segoe UI';
                font-size: {size};
                font-weight: {weight};
                color: {color};
                background: transparent;
                border: none;
                padding: 0px;
            }}
        """)

    def _show_menu(self, pos):
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background: #2A2A2A; color: white;
                border: none; border-radius: 6px; padding: 4px;
                font-family: 'Segoe UI'; font-size: 12px;
            }
            QMenu::item { padding: 6px 20px; border-radius: 4px; }
            QMenu::item:selected { background: #444444; }
        """)
        copy_act = menu.addAction("复制")
        copy_act.setShortcut("Ctrl+C")
        action = menu.exec(self.mapToGlobal(pos))
        if action == copy_act:
            QApplication.clipboard().setText(self.text())


class ResultRow(QWidget):
    """结果行：标签 + 可复制值"""
    def __init__(self, label, highlight=False, copyable=False):
        super().__init__()
        self.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 3, 0, 3)
        layout.setSpacing(8)

        lbl = QLabel(label)
        lbl.setObjectName("result_label")
        lbl.setFixedWidth(72)
        lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        color  = "#0A7A3E" if highlight else "#1A1A1A"
        weight = "700"     if highlight else "600"
        size   = "15px"    if highlight else "13px"
        self.val = CopyableEdit(color=color, size=size, weight=weight)
        self.val.setFixedWidth(150)

        layout.addWidget(lbl)
        layout.addWidget(self.val)
        layout.addStretch()

    def set(self, text): self.val.setText(text)
    def reset(self): self.val.setText("—")

# ============================================================
# 主窗口
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("亚马逊定价计算器")
        self.setFixedSize(680, 540)
        self.rates = {"US": None, "UK": None}
        self.mode  = "US"
        self._build()
        self._fetch_rates()

    def _build(self):
        central = QWidget(); central.setObjectName("central")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0,0,0,0)
        root.setSpacing(0)

        # ── 标题栏 ──
        tb = QFrame(); tb.setObjectName("titlebar"); tb.setFixedHeight(56)
        tbl = QHBoxLayout(tb); tbl.setContentsMargins(20,0,20,0)

        icon = QLabel("🛒"); icon.setStyleSheet("font-size:20px;")
        title = QLabel("亚马逊定价计算器"); title.setObjectName("title")
        tbl.addWidget(icon); tbl.addSpacing(8); tbl.addWidget(title)
        tbl.addStretch()

        # 置顶按钮
        self.pin_btn = QPushButton("📌 置顶")
        self.pin_btn.setObjectName("mode_btn")
        self.pin_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.pin_btn.clicked.connect(self._toggle_pin)
        self._pinned = False
        tbl.addWidget(self.pin_btn)
        tbl.addSpacing(8)

        # 模式切换
        self.mode_btns = {}
        for label, val in [("🇺🇸  美国","US"),("🇬🇧  英国","UK")]:
            btn = QPushButton(label)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _, v=val: self._set_mode(v))
            self.mode_btns[val] = btn
            tbl.addWidget(btn)

        tbl.addSpacing(12)
        self.rate_label = QLabel("正在获取汇率…")
        self.rate_label.setObjectName("rate_label")
        tbl.addWidget(self.rate_label)

        root.addWidget(tb)
        root.addWidget(make_separator())

        # ── 主体 ──
        body = QWidget(); body.setStyleSheet("background:#DCDCDE;")
        bl = QHBoxLayout(body); bl.setContentsMargins(16,16,16,16); bl.setSpacing(14)

        # ── 左列：输入 ──
        left_card = make_card()
        ll = QVBoxLayout(left_card); ll.setContentsMargins(20,16,20,16); ll.setSpacing(6)

        sec1 = QLabel("产品信息"); sec1.setObjectName("section")
        ll.addWidget(sec1); ll.addSpacing(4)

        self.fields = {
            "cost":   FieldRow("产品成本", "", "元"),
            "length": FieldRow("长",       "", "cm"),
            "width":  FieldRow("宽",       "", "cm"),
            "height": FieldRow("高",       "", "cm"),
            "weight": FieldRow("实重",     "", "kg"),
            "profit": FieldRow("利润率",   "50", "%"),
        }
        for f in self.fields.values():
            ll.addWidget(f)
            f.entry.returnPressed.connect(self._calculate)

        ll.addSpacing(8)

        # 按钮行
        btn_row = QHBoxLayout(); btn_row.setSpacing(10)
        self.calc_btn = QPushButton("计  算")
        self.calc_btn.setMinimumHeight(42)
        self.calc_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.calc_btn.clicked.connect(self._calculate)
        self.calc_btn.setStyleSheet("""
            QPushButton {
                font-family:'Segoe UI'; font-size:14px; font-weight:700;
                color:white; background:#0067C0;
                border:none; border-radius:8px; padding:10px 24px;
            }
            QPushButton:hover   { background:#1478CC; }
            QPushButton:pressed { background:#005294; }
        """)
        clear_btn = QPushButton("清  空")
        clear_btn.setMinimumHeight(42)
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.clicked.connect(self._clear)
        clear_btn.setStyleSheet("""
            QPushButton {
                font-family:'Segoe UI'; font-size:13px; font-weight:500;
                color:#2A2A2A; background:#E8E8E8;
                border:1.5px solid #C8C8C8; border-radius:8px; padding:10px 18px;
            }
            QPushButton:hover  { background:#DCDCDC; }
            QPushButton:pressed { background:#CCCCCC; }
        """)
        btn_row.addWidget(self.calc_btn, 3)
        btn_row.addWidget(clear_btn, 1)
        ll.addLayout(btn_row)
        ll.addStretch()

        # ── 右列：结果 ──
        right_card = make_card()
        rl = QVBoxLayout(right_card); rl.setContentsMargins(20,16,20,16); rl.setSpacing(4)

        sec2 = QLabel("计算结果"); sec2.setObjectName("section")
        rl.addWidget(sec2); rl.addSpacing(4)

        # 渠道单独一行（较小字体）
        ch_row = QWidget(); ch_row.setStyleSheet("background:transparent;")
        chl = QHBoxLayout(ch_row); chl.setContentsMargins(0,2,0,6); chl.setSpacing(8)
        ch_lbl = QLabel("渠道"); ch_lbl.setObjectName("result_label"); ch_lbl.setFixedWidth(72)
        ch_lbl.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        self.channel_val = CopyableEdit(color="#333333", size="12px", weight="500")
        chl.addWidget(ch_lbl); chl.addWidget(self.channel_val); chl.addStretch()
        rl.addWidget(ch_row)
        rl.addWidget(make_separator())
        rl.addSpacing(4)

        self.results = {
            "cost_o":   ResultRow("产品成本"),
            "ship_o":   ResultRow("头程运费"),
            "e2_o":     ResultRow("总成本"),
            "sell_o":   ResultRow("售  价",    highlight=True),
            "markup_o": ResultRow("加价(RMB)", highlight=True),
            "profit_o": ResultRow("利润率"),
        }
        for rw in self.results.values():
            rl.addWidget(rw)

        rl.addSpacing(8)
        self.warn_label = QLabel("")
        self.warn_label.setObjectName("warn_label")
        self.warn_label.setWordWrap(True)
        rl.addWidget(self.warn_label)
        rl.addStretch()

        bl.addWidget(left_card, 5)
        bl.addWidget(right_card, 5)
        root.addWidget(body)

        self._set_mode("US")

    def _toggle_pin(self):
        self._pinned = not self._pinned
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, self._pinned)
        self.show()
        self.pin_btn.setText("📌 置顶 ✓" if self._pinned else "📌 置顶")
        self.pin_btn.setObjectName("mode_btn_active" if self._pinned else "mode_btn")
        self.pin_btn.setStyle(self.pin_btn.style())

    def _set_mode(self, val):
        self.mode = val
        for v, btn in self.mode_btns.items():
            btn.setObjectName("mode_btn_active" if v==val else "mode_btn")
            btn.setStyle(btn.style())
        rate = self.rates.get(val)
        sym  = "USD" if val=="US" else "GBP"
        self.rate_label.setText(
            f"1 {sym} = ¥{rate:.4f}" if rate else f"获取 {sym} 汇率中…")
        for r in self.results.values(): r.reset()
        self.channel_val.setText("—")
        self.warn_label.setText("")

    def _fetch_rates(self):
        def _f():
            r = get_all_rates()
            self.rates["US"] = r["USD"]
            self.rates["UK"] = r["GBP"]
            QTimer.singleShot(0, lambda: self._set_mode(self.mode))
        threading.Thread(target=_f, daemon=True).start()

    def _clear(self):
        defaults = {"cost":"","length":"","width":"","height":"","weight":"","profit":"50"}
        for k, f in self.fields.items(): f.clear(defaults.get(k,""))
        for r in self.results.values(): r.reset()
        self.channel_val.setText("—")
        self.warn_label.setText("")

    def _calculate(self):
        try:
            cost   = float(self.fields["cost"].text())
            length = float(self.fields["length"].text())
            width  = float(self.fields["width"].text())
            height = float(self.fields["height"].text())
            weight = float(self.fields["weight"].text())
            profit = float(self.fields["profit"].text())
        except ValueError:
            QMessageBox.warning(self, "输入错误", "请确认所有字段均已填写数字")
            return
        if not (0 < profit < 100):
            QMessageBox.warning(self, "输入错误", "利润率需在 0~100 之间（如 50）")
            return
        profit /= 100

        fx    = self.rates.get(self.mode) or (7.2 if self.mode=="US" else 9.0)
        sym   = "$" if self.mode=="US" else "£"
        coeff = 1.38*2 if self.mode=="US" else 1.34*2

        vol_w     = calc_vol_weight(length, width, height)
        billing_w = calc_billing(weight, length, width, height)
        wt_type   = "体积重" if vol_w > weight else "实重"

        if self.mode == "US":
            channel, total, base, sur, warns = shipping_us(billing_w, length, width, height, weight)
        else:
            channel, total, base, sur, warns = shipping_uk(billing_w, length, width, height)

        result = calc_price(cost, total, profit, fx, coeff)
        warn_extra = ""
        if result["markup"] < 0:
            warn_extra = "⚠ 加价为负，已切换智赢参数5/5重新计算\n"
            result = calc_price(cost, total, profit, fx, coeff, 5, 5)

        self.channel_val.setText(f"{channel}  ({billing_w:.2f}kg · {wt_type})")
        self.results["cost_o"].set(f"¥{cost:.2f}")
        self.results["ship_o"].set(f"¥{total:.2f}")
        self.results["e2_o"].set(f"¥{result['E2']:.2f}")
        self.results["sell_o"].set(f"{sym}{result['sell']:.2f}")
        self.results["markup_o"].set(f"¥{result['markup']:.2f}")
        self.results["profit_o"].set(f"{result['profit']:.1f}%")
        self.warn_label.setText(warn_extra + "\n".join(warns))


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    app.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())    