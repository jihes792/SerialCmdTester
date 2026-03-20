import sys
import csv
import re
import time
from dataclasses import dataclass
from typing import List, Optional, Dict, Any, Tuple

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel, QLineEdit,
    QCheckBox, QSpinBox, QPlainTextEdit, QTableWidget, QTableWidgetItem,
    QHeaderView
)

import serial
from serial.tools import list_ports

from openpyxl import load_workbook


# -------------------- 数据结构 --------------------
@dataclass
class CmdRow:
    command: str
    expected: str = ""
    timeout_ms: int = 1200
    match: str = "contains"   # contains / exact / regex


def normalize_header(s: str) -> str:
    return (s or "").strip().lower().replace(" ", "").replace("_", "")


def pick_column(headers: List[str], candidates: List[str]) -> Optional[int]:
    norm = [normalize_header(h) for h in headers]
    cand_norm = [normalize_header(c) for c in candidates]
    for ci in cand_norm:
        if ci in norm:
            return norm.index(ci)
    return None


def load_csv(path: str) -> List[CmdRow]:
    rows: List[CmdRow] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        all_rows = list(reader)
    if not all_rows:
        return rows

    headers = all_rows[0]
    i_cmd = pick_column(headers, ["command", "cmd", "命令", "指令"])
    i_exp = pick_column(headers, ["expected", "expect", "response", "期望", "反馈", "回包"])
    i_to  = pick_column(headers, ["timeoutms", "timeout", "超时", "超时ms"])
    i_ma  = pick_column(headers, ["match", "匹配", "matchtype"])

    if i_cmd is None:
        raise ValueError("未找到指令列（Command/Cmd/命令/指令）")

    for r in all_rows[1:]:
        if not r or (i_cmd < len(r) and not (r[i_cmd] or "").strip()):
            continue
        cmd = (r[i_cmd] or "").strip()
        exp = (r[i_exp] or "").strip() if i_exp is not None and i_exp < len(r) else ""
        to  = (r[i_to] or "").strip() if i_to is not None and i_to < len(r) else ""
        ma  = (r[i_ma] or "").strip().lower() if i_ma is not None and i_ma < len(r) else ""

        timeout_ms = 1200
        if to:
            try:
                timeout_ms = int(float(to))
            except:
                timeout_ms = 1200

        match = ma if ma in ("contains", "exact", "regex") else "contains"
        rows.append(CmdRow(command=cmd, expected=exp, timeout_ms=timeout_ms, match=match))
    return rows


def load_xlsx(path: str) -> List[CmdRow]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    values = list(ws.values)
    if not values:
        return []

    headers = [str(x) if x is not None else "" for x in values[0]]
    i_cmd = pick_column(headers, ["command", "cmd", "命令", "指令"])
    i_exp = pick_column(headers, ["expected", "expect", "response", "期望", "反馈", "回包"])
    i_to  = pick_column(headers, ["timeoutms", "timeout", "超时", "超时ms"])
    i_ma  = pick_column(headers, ["match", "匹配", "matchtype"])

    if i_cmd is None:
        raise ValueError("未找到指令列（Command/Cmd/命令/指令）")

    rows: List[CmdRow] = []
    for row in values[1:]:
        row = list(row)
        if i_cmd >= len(row) or row[i_cmd] is None or str(row[i_cmd]).strip() == "":
            continue

        cmd = str(row[i_cmd]).strip()
        exp = str(row[i_exp]).strip() if i_exp is not None and i_exp < len(row) and row[i_exp] is not None else ""
        timeout_ms = 1200
        if i_to is not None and i_to < len(row) and row[i_to] is not None:
            try:
                timeout_ms = int(float(row[i_to]))
            except:
                timeout_ms = 1200

        match = "contains"
        if i_ma is not None and i_ma < len(row) and row[i_ma] is not None:
            ma = str(row[i_ma]).strip().lower()
            if ma in ("contains", "exact", "regex"):
                match = ma

        rows.append(CmdRow(command=cmd, expected=exp, timeout_ms=timeout_ms, match=match))
    return rows


def load_table(path: str) -> List[CmdRow]:
    if path.lower().endswith(".csv"):
        return load_csv(path)
    if path.lower().endswith(".xlsx"):
        return load_xlsx(path)
    raise ValueError("仅支持 .xlsx / .csv")


# -------------------- 串口读写/判定 --------------------
def read_response(ser: serial.Serial, total_timeout_ms: int, idle_gap_ms: int = 120) -> str:
    """
    读取响应：直到总超时，或收到数据后 idle_gap 内再无新数据。
    """
    deadline = time.time() + total_timeout_ms / 1000.0
    idle_deadline = None
    buf = bytearray()

    while time.time() < deadline:
        n = ser.in_waiting
        if n:
            data = ser.read(n)
            buf.extend(data)
            idle_deadline = time.time() + idle_gap_ms / 1000.0
        else:
            if idle_deadline is not None and time.time() >= idle_deadline:
                break
            time.sleep(0.01)

    return buf.decode("utf-8", errors="ignore").strip()


def match_expected(actual: str, expected: str, match: str, ignore_case: bool) -> bool:
    if expected is None:
        expected = ""
    if expected.strip() == "":
        return True  # 没填期望则不判定，默认 PASS

    a = actual or ""
    e = expected

    if ignore_case:
        a_cmp = a.lower()
        e_cmp = e.lower()
    else:
        a_cmp = a
        e_cmp = e

    if match == "exact":
        return a_cmp.strip() == e_cmp.strip()
    if match == "regex":
        flags = re.IGNORECASE if ignore_case else 0
        try:
            return re.search(e, a, flags=flags) is not None
        except re.error:
            return False
    # contains
    return e_cmp in a_cmp


# -------------------- 测试线程 --------------------
class Runner(QThread):
    sig_log = Signal(str)
    sig_row_result = Signal(int, str, bool)  # row_index, actual, pass
    sig_done = Signal()

    def __init__(self,
                 port: str,
                 baud: int,
                 append_crlf: bool,
                 idle_gap_ms: int,
                 ignore_case: bool,
                 rows: List[CmdRow]):
        super().__init__()
        self.port = port
        self.baud = baud
        self.append_crlf = append_crlf
        self.idle_gap_ms = idle_gap_ms
        self.ignore_case = ignore_case
        self.rows = rows
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            self.sig_log.emit(f"[RUN] Open {self.port} @ {self.baud}")
            with serial.Serial(self.port, self.baud, timeout=0.05) as ser:
                ser.reset_input_buffer()

                for idx, r in enumerate(self.rows):
                    if self._stop:
                        self.sig_log.emit("[RUN] Stopped by user.")
                        break

                    cmd = (r.command or "").strip()
                    if not cmd.endswith("!"):
                        cmd += "!"
                    payload = cmd + ("\r\n" if self.append_crlf else "")

                    self.sig_log.emit(f">>> {cmd}")
                    ser.write(payload.encode("ascii", errors="ignore"))

                    actual = read_response(ser, r.timeout_ms, idle_gap_ms=self.idle_gap_ms)
                    ok = match_expected(actual, r.expected, r.match, self.ignore_case)
                    self.sig_log.emit(f"<<< {actual}")
                    self.sig_row_result.emit(idx, actual, ok)

        except Exception as e:
            self.sig_log.emit(f"[ERR] {type(e).__name__}: {e}")
        finally:
            self.sig_done.emit()


# -------------------- GUI --------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serial Command Table Tester")
        self.resize(1100, 720)

        self.rows: List[CmdRow] = []
        self.runner: Optional[Runner] = None

        # ---- 顶部控件 ----
        self.cmb_port = QComboBox()
        self.btn_refresh = QPushButton("刷新端口")
        self.ed_baud = QLineEdit("115200")
        self.chk_crlf = QCheckBox("发送追加 CRLF")
        self.chk_ignore_case = QCheckBox("忽略大小写匹配")
        self.chk_ignore_case.setChecked(True)

        self.spin_idle = QSpinBox()
        self.spin_idle.setRange(10, 2000)
        self.spin_idle.setValue(120)
        self.spin_idle.setSuffix(" ms (回包空闲结束)")

        self.btn_import = QPushButton("导入表格")
        self.btn_run = QPushButton("运行测试")
        self.btn_stop = QPushButton("停止")
        self.btn_stop.setEnabled(False)

        # ---- 表格 ----
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["#", "Command", "Expected", "Timeout(ms)", "Actual", "Result"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        # ---- 日志 ----
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)

        # ---- 布局 ----
        top = QHBoxLayout()
        top.addWidget(QLabel("端口:"))
        top.addWidget(self.cmb_port, 2)
        top.addWidget(self.btn_refresh)
        top.addSpacing(10)
        top.addWidget(QLabel("波特率:"))
        top.addWidget(self.ed_baud)
        top.addSpacing(10)
        top.addWidget(self.chk_crlf)
        top.addWidget(self.chk_ignore_case)
        top.addWidget(self.spin_idle, 2)
        top.addStretch(1)
        top.addWidget(self.btn_import)
        top.addWidget(self.btn_run)
        top.addWidget(self.btn_stop)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.addLayout(top)
        layout.addWidget(self.table, 3)
        layout.addWidget(QLabel("运行日志："))
        layout.addWidget(self.log, 2)
        self.setCentralWidget(central)

        # ---- 信号 ----
        self.btn_refresh.clicked.connect(self.refresh_ports)
        self.btn_import.clicked.connect(self.import_table)
        self.btn_run.clicked.connect(self.start_run)
        self.btn_stop.clicked.connect(self.stop_run)

        self.refresh_ports()

    def append_log(self, s: str):
        self.log.appendPlainText(s)

    def refresh_ports(self):
        self.cmb_port.clear()
        ports = list_ports.comports()
        for p in ports:
            # 显示友好文本，但 data 存设备名
            self.cmb_port.addItem(f"{p.device}  ({p.description})", p.device)
        if self.cmb_port.count() == 0:
            self.cmb_port.addItem("无串口", "")
        self.append_log("[UI] Ports refreshed.")

    def import_table(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择指令表格", "", "Table (*.xlsx *.csv)")
        if not path:
            return
        try:
            self.rows = load_table(path)
            self.populate_table()
            self.append_log(f"[UI] Imported: {path}, rows={len(self.rows)}")
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"{type(e).__name__}: {e}")

    def populate_table(self):
        self.table.setRowCount(0)
        for i, r in enumerate(self.rows):
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table.setItem(i, 1, QTableWidgetItem(r.command))
            self.table.setItem(i, 2, QTableWidgetItem(r.expected))
            self.table.setItem(i, 3, QTableWidgetItem(str(r.timeout_ms)))
            self.table.setItem(i, 4, QTableWidgetItem(""))
            self.table.setItem(i, 5, QTableWidgetItem(""))

    def set_row_result(self, row: int, actual: str, ok: bool):
        self.table.item(row, 4).setText(actual)
        self.table.item(row, 5).setText("PASS" if ok else "FAIL")

        color = QColor(200, 255, 200) if ok else QColor(255, 210, 210)
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if item:
                item.setBackground(color)

    def start_run(self):
        if not self.rows:
            QMessageBox.warning(self, "提示", "请先导入表格（xlsx/csv）。")
            return

        port = self.cmb_port.currentData()
        if not port:
            QMessageBox.warning(self, "提示", "请选择有效串口。")
            return

        try:
            baud = int(self.ed_baud.text().strip())
        except:
            QMessageBox.warning(self, "提示", "波特率不是有效整数。")
            return

        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)

        self.runner = Runner(
            port=port,
            baud=baud,
            append_crlf=self.chk_crlf.isChecked(),
            idle_gap_ms=int(self.spin_idle.value()),
            ignore_case=self.chk_ignore_case.isChecked(),
            rows=self.rows
        )
        self.runner.sig_log.connect(self.append_log)
        self.runner.sig_row_result.connect(self.set_row_result)
        self.runner.sig_done.connect(self.on_done)
        self.runner.start()

    def stop_run(self):
        if self.runner:
            self.runner.stop()
        self.append_log("[UI] Stop requested.")

    def on_done(self):
        self.append_log("[RUN] Done.")
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.runner = None


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
