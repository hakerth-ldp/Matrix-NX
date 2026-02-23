import sys
import time
import csv
import queue
from dataclasses import dataclass
from collections import deque
from datetime import datetime, date
from pathlib import Path

import serial
from serial.tools import list_ports

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QSpinBox, QDoubleSpinBox,
    QProgressBar, QMessageBox, QTabWidget, QTableWidget, QTableWidgetItem,
    QPlainTextEdit, QLineEdit, QHeaderView
)

# ---------------------------
# Serial config
# ---------------------------
BAUDRATE = 115200
DATABITS = serial.EIGHTBITS
STOPBITS = serial.STOPBITS_ONE
PARITY = serial.PARITY_NONE

READ_TIMEOUT_S = 2.0
WRITE_TIMEOUT_S = 2.0
SERIAL_POLL_TIMEOUT_S = 0.15  # small per-read timeout; we manage overall deadlines ourselves

# ---------------------------
# Logging
# ---------------------------
LOG_DIR = Path(r"C:\Coherent")
LOG_PREFIX_NORMAL = "matrixNX"
LOG_PREFIX_RD = "matrixNX_RD"

# ---------------------------
# Commands
# ---------------------------
CMD_ALL = "All?"
CMD_PASS = "CALibration:PASSword Rocinante"
CMD_TEMPS = "SERVice:XALL? TEMPeratures"
CMD_OTHERS = "SERVice:XALL? OTHers"

# ---------------------------
# Field names
# ---------------------------
FIELD_NAMES_NORMAL = [
    "Status",
    "Warnings",
    "Faults",
    "Actual Housing Temperature",
    "Actual Laserdiode Temperature",
    "Actual SHG Temperature",
    "Actual THG Temperature",
    "Operation Hours",
    "Laserdiode Hours",
    "THG Crystal Hours",
    "Actual THG Spot Hours",
    "Actual THG Spot Number",
    "Actual THG Spot Status",
    "Scaled UV-Power",
]

FIELD_NAMES_RD = [
    # TEMPeratures (10)
    "Reso Temperature",
    "Vanadat Temperature",
    "Laserdiode Temperature",
    "Housing Temperature",
    "SHG Temperature",
    "THG Temperature",
    "Actual SHG Voltage",
    "Actual THG Voltage",
    "Actual SHG Current",
    "Actual THG Current",
    # OTHers (7)
    "Fan Output Drive",
    "Laserdiode Current",
    "System Status Flags",
    "Scaled UV-Power",
    "Raw UV-Power",
    "Operation Hours",
    "LD Hours",
]

TEMP_KEYS_ALL = ["Reso", "Vanadat", "Diode", "Housing", "SHG", "THG"]

TEMP_IDX_NORMAL = {
    "Housing": 3,
    "Diode": 4,
    "SHG": 5,
    "THG": 6,
}
TEMP_IDX_RD = {
    "Reso": 0,
    "Vanadat": 1,
    "Diode": 2,
    "Housing": 3,
    "SHG": 4,
    "THG": 5,
}

# ---------------------------
# Helpers
# ---------------------------
def float_to_de_str(x: float, decimals: int = 3) -> str:
    return f"{x:.{decimals}f}".replace(".", ",")

def timestamp_de() -> str:
    return datetime.now().strftime("%d.%m.%Y %H:%M:%S")

def try_format_value_de(v: str) -> str:
    s = (v or "").strip()
    if s == "":
        return ""
    if ":" in s:  # times like 45:03:14
        return s
    try:
        x = float(s)
        return float_to_de_str(x, 3)
    except ValueError:
        return s

def is_err_line(s: str) -> bool:
    s = (s or "").strip().upper()
    return s.startswith("ERR")

# ---------------------------
# Daily CSV loggers
# ---------------------------
class DailyCsvLoggerNormal:
    """Normal mode: timestamp + sec + first 8 values (Status..Operation Hours)"""
    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self._current_day: date | None = None
        self._fh = None
        self._writer = None

    def _open_for_today(self):
        today = date.today()
        if self._current_day == today and self._fh:
            return

        if self._fh:
            try:
                self._fh.flush()
                self._fh.close()
            except Exception:
                pass

        self._current_day = today
        fname = f"{LOG_PREFIX_NORMAL}_{today.isoformat()}.csv"
        fpath = self.base_dir / fname
        is_new = not fpath.exists()

        self._fh = open(fpath, "a", encoding="utf-8-sig", newline="")
        self._writer = csv.writer(self._fh, delimiter=";")

        if is_new:
            header = [
                "Zeitstempel",
                "SekSeitStart",
                "Status",
                "Warnings",
                "Faults",
                "T_Housing",
                "T_Diode",
                "T_SHG",
                "T_THG",
                "OperationHours",
            ]
            self._writer.writerow(header)
            self._fh.flush()

    def log_row(self, tstamp: str, sec_since_start: float, fields14: list[str]):
        self._open_for_today()

        status = fields14[0]
        warnings = fields14[1]
        faults = fields14[2]
        t_h = float(fields14[3])
        t_d = float(fields14[4])
        t_shg = float(fields14[5])
        t_thg = float(fields14[6])
        op_hours = fields14[7]

        row = [
            tstamp,
            float_to_de_str(sec_since_start, 3),
            status,
            warnings,
            faults,
            float_to_de_str(t_h, 3),
            float_to_de_str(t_d, 3),
            float_to_de_str(t_shg, 3),
            float_to_de_str(t_thg, 3),
            op_hours,
        ]
        self._writer.writerow(row)
        self._fh.flush()

    def close(self):
        if self._fh:
            try:
                self._fh.flush()
                self._fh.close()
            except Exception:
                pass
        self._fh = None
        self._writer = None
        self._current_day = None


class DailyCsvLoggerRD:
    """R+D mode: timestamp + sec + ALL 17 values (10 temps + 7 others)"""
    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self._current_day: date | None = None
        self._fh = None
        self._writer = None

    def _open_for_today(self):
        today = date.today()
        if self._current_day == today and self._fh:
            return

        if self._fh:
            try:
                self._fh.flush()
                self._fh.close()
            except Exception:
                pass

        self._current_day = today
        fname = f"{LOG_PREFIX_RD}_{today.isoformat()}.csv"
        fpath = self.base_dir / fname
        is_new = not fpath.exists()

        self._fh = open(fpath, "a", encoding="utf-8-sig", newline="")
        self._writer = csv.writer(self._fh, delimiter=";")

        if is_new:
            header = ["Zeitstempel", "SekSeitStart"] + FIELD_NAMES_RD
            self._writer.writerow(header)
            self._fh.flush()

    def log_row(self, tstamp: str, sec_since_start: float, values17: list[str]):
        self._open_for_today()
        row = [tstamp, float_to_de_str(sec_since_start, 3)] + [try_format_value_de(v) for v in values17]
        self._writer.writerow(row)
        self._fh.flush()

    def close(self):
        if self._fh:
            try:
                self._fh.flush()
                self._fh.close()
            except Exception:
                pass
        self._fh = None
        self._writer = None
        self._current_day = None


# ---------------------------
# Temperature stats
# ---------------------------
@dataclass
class TempStats:
    window: deque
    sum_: float
    mean: float
    dev: float

    bar_peak_pos: float
    bar_peak_neg: float
    t_bar_pos: float
    t_bar_neg: float

    latched_pos: float
    latched_neg: float

def init_tempstats(n: int = 50) -> TempStats:
    return TempStats(
        window=deque(maxlen=n),
        sum_=0.0,
        mean=0.0,
        dev=0.0,
        bar_peak_pos=0.0,
        bar_peak_neg=0.0,
        t_bar_pos=0.0,
        t_bar_neg=0.0,
        latched_pos=0.0,
        latched_neg=0.0,
    )

def update_tempstats(stats: TempStats, value: float, hold_s: float, now_mono: float) -> None:
    if len(stats.window) == stats.window.maxlen:
        stats.sum_ -= stats.window[0]
    stats.window.append(value)
    stats.sum_ += value

    stats.mean = stats.sum_ / max(1, len(stats.window))
    stats.dev = value - stats.mean

    if stats.dev > stats.latched_pos:
        stats.latched_pos = stats.dev
    if stats.dev < stats.latched_neg:
        stats.latched_neg = stats.dev

    if stats.dev > stats.bar_peak_pos:
        stats.bar_peak_pos = stats.dev
        stats.t_bar_pos = now_mono
    if stats.dev < stats.bar_peak_neg:
        stats.bar_peak_neg = stats.dev
        stats.t_bar_neg = now_mono

    if now_mono - stats.t_bar_pos > hold_s:
        stats.bar_peak_pos = max(stats.dev, 0.0)
        stats.t_bar_pos = now_mono
    if now_mono - stats.t_bar_neg > hold_s:
        stats.bar_peak_neg = min(stats.dev, 0.0)
        stats.t_bar_neg = now_mono


# ---------------------------
# Terminal history line edit (Up/Down)
# ---------------------------
class HistoryLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self._history: list[str] = []
        self._idx: int = -1

    def add_history(self, cmd: str, max_items: int = 50):
        cmd = (cmd or "").strip()
        if not cmd:
            return
        if self._history and self._history[-1] == cmd:
            self._idx = -1
            return
        self._history.append(cmd)
        if len(self._history) > max_items:
            self._history = self._history[-max_items:]
        self._idx = -1

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Up:
            if self._history:
                if self._idx == -1:
                    self._idx = len(self._history) - 1
                else:
                    self._idx = max(0, self._idx - 1)
                self.setText(self._history[self._idx])
                self.setCursorPosition(len(self.text()))
            return
        if event.key() == Qt.Key_Down:
            if self._history:
                if self._idx == -1:
                    return
                self._idx += 1
                if self._idx >= len(self._history):
                    self._idx = -1
                    self.clear()
                else:
                    self.setText(self._history[self._idx])
                    self.setCursorPosition(len(self.text()))
            return
        super().keyPressEvent(event)

# ---------------------------
# Serial worker thread
# ---------------------------
class SerialWorker(QThread):
    connected = Signal(str)
    disconnected = Signal(str)
    error = Signal(str)
    warn = Signal(str)
    data = Signal(object)
    term_line = Signal(str)

    def __init__(self):
        super().__init__()
        self._port_name = None
        self._interval_ms = 500
        self._stop = False
        self._ser = None
        self._mode = "normal"
        self._manual_q: queue.Queue[str] = queue.Queue()

    def configure(self, port_name: str, interval_ms: int, mode: str):
        self._port_name = port_name
        self._interval_ms = max(10, int(interval_ms))
        self._mode = mode

    def stop(self):
        self._stop = True

    def enqueue_manual(self, cmd: str):
        self._manual_q.put(cmd)

    def _close_serial(self):
        if self._ser:
            try:
                self._ser.close()
            except Exception:
                pass
            self._ser = None

    def _send_cmd(self, cmd: str):
        self._ser.write((cmd + "\r\n").encode("ascii", errors="replace"))

    def _readline_nonempty(self, deadline: float) -> str | None:
        while time.monotonic() < deadline and not self._stop:
            raw = self._ser.readline()
            if not raw:
                return None
            s = raw.decode(errors="replace").strip()
            if s != "":
                return s
        return None

    def _read_frame_data_plus_ok(self, expected_fields: int, overall_timeout_s: float = READ_TIMEOUT_S):
        deadline = time.monotonic() + overall_timeout_s
        data_parts = None
        ok_seen = False
        term = None

        while time.monotonic() < deadline and not self._stop:
            s = self._readline_nonempty(deadline)
            if s is None:
                break

            su = s.upper()
            if su == "OK" or is_err_line(su):
                term = s
                ok_seen = (su == "OK")
                if data_parts is not None or expected_fields == 0:
                    return data_parts, ok_seen, term
                continue

            parts = [p.strip() for p in s.split(",")]
            if expected_fields > 0 and len(parts) == expected_fields:
                data_parts = parts
                continue

        return data_parts, ok_seen, term

    def _read_manual_response(self, overall_timeout_s: float = READ_TIMEOUT_S):
        deadline = time.monotonic() + overall_timeout_s
        lines: list[str] = []

        first = self._readline_nonempty(deadline)
        if first is None:
            return lines
        lines.append(first)

        if first.upper() == "OK" or is_err_line(first):
            return lines

        idle_deadline = time.monotonic() + 0.35
        while time.monotonic() < deadline and not self._stop:
            raw = self._ser.readline()
            if raw:
                s = raw.decode(errors="replace").strip()
                if s:
                    lines.append(s)
                    idle_deadline = time.monotonic() + 0.35
                    if s.upper() == "OK" or is_err_line(s):
                        break
            if time.monotonic() > idle_deadline:
                break

        return lines

    def run(self):
        if not self._port_name:
            self.error.emit("Kein COM-Port gewählt.")
            return

        try:
            self._ser = serial.Serial(
                port=self._port_name,
                baudrate=BAUDRATE,
                bytesize=DATABITS,
                parity=PARITY,
                stopbits=STOPBITS,
                timeout=SERIAL_POLL_TIMEOUT_S,
                write_timeout=WRITE_TIMEOUT_S,
            )
            try:
                self._ser.reset_input_buffer()
                self._ser.reset_output_buffer()
            except Exception:
                pass

            self.connected.emit(f"Verbunden: {self._port_name} ({self._mode})")
        except Exception as e:
            self.error.emit(f"Verbindung fehlgeschlagen: {e}")
            self._close_serial()
            return

        if self._mode == "rd":
            try:
                self._send_cmd(CMD_PASS)
                self.term_line.emit(f">> {CMD_PASS}")
                _, okp, term = self._read_frame_data_plus_ok(expected_fields=0)
                if term:
                    self.term_line.emit(f"<< {term}")
                if not okp:
                    self.warn.emit("Warnung: Kein OK auf Passwort – versuche trotzdem weiter.")
            except Exception as e:
                self.warn.emit(f"Warnung: Passwort-Handshake fehlgeschlagen: {e}")

        interval_s = self._interval_ms / 1000.0
        next_cycle = time.monotonic()

        self._stop = False
        while not self._stop:
            try:
                cmd = self._manual_q.get_nowait()
            except queue.Empty:
                cmd = None

            if cmd:
                cmd = cmd.strip()
                if cmd:
                    try:
                        self._send_cmd(cmd)
                        self.term_line.emit(f">> {cmd}")
                        lines = self._read_manual_response()
                        if not lines:
                            self.term_line.emit("<< (keine Antwort / Timeout)")
                        else:
                            for ln in lines:
                                self.term_line.emit(f"<< {ln}")
                    except Exception as e:
                        self.term_line.emit(f"<< Fehler: {e}")
                continue

            now = time.monotonic()
            if now < next_cycle:
                time.sleep(min(0.05, next_cycle - now))
                continue

            if self._mode == "normal":
                try:
                    self._send_cmd(CMD_ALL)
                    parts, ok_seen, _ = self._read_frame_data_plus_ok(expected_fields=14)
                    if parts is not None:
                        if not ok_seen:
                            self.warn.emit("Warnung: OK fehlt (All?)")
                        self.data.emit({"mode": "normal", "values": parts, "ok": ok_seen})
                except Exception as e:
                    self.error.emit(f"Kommunikationsfehler: {e}")
                    break
            else:
                try:
                    self._send_cmd(CMD_TEMPS)
                    temps, ok_t, _ = self._read_frame_data_plus_ok(expected_fields=10)
                    if not ok_t:
                        self.warn.emit("Warnung: OK fehlt (TEMPeratures)")

                    self._send_cmd(CMD_OTHERS)
                    others, ok_o, _ = self._read_frame_data_plus_ok(expected_fields=7)
                    if not ok_o:
                        self.warn.emit("Warnung: OK fehlt (OTHers)")

                    if temps is None and others is None:
                        pass
                    else:
                        if temps is None:
                            temps = [""] * 10
                        if others is None:
                            others = [""] * 7
                        combined = temps + others
                        self.data.emit({"mode": "rd", "values": combined, "ok_t": ok_t, "ok_o": ok_o})
                except Exception as e:
                    self.error.emit(f"Kommunikationsfehler: {e}")
                    break

            next_cycle = max(next_cycle + interval_s, time.monotonic())

        self._close_serial()
        self.disconnected.emit("Getrennt")


# ---------------------------
# Main GUI
# ---------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Matrix NX RS232 Monitor")

        self.worker = SerialWorker()
        self.worker.connected.connect(self.on_connected)
        self.worker.disconnected.connect(self.on_disconnected)
        self.worker.error.connect(self.on_error)
        self.worker.warn.connect(self.on_warn)
        self.worker.data.connect(self.on_data)
        self.worker.term_line.connect(self.on_term_line)

        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        self.logger_normal = DailyCsvLoggerNormal(LOG_DIR)
        self.logger_rd = DailyCsvLoggerRD(LOG_DIR)

        self.t0_mono = None
        self.current_mode = None
        self.stats = {k: init_tempstats(50) for k in TEMP_KEYS_ALL}

        root = QWidget()
        self.setCentralWidget(root)
        main_layout = QVBoxLayout(root)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self._build_tab_connection()
        self._build_tab_live()
        self._build_tab_temperatures()
        self._build_tab_terminal()

        self.lbl_status = QLabel("Nicht verbunden")
        self.lbl_status.setTextInteractionFlags(Qt.TextSelectableByMouse)
        main_layout.addWidget(self.lbl_status)

        self.refresh_ports()
        self._set_reasonable_window_size()

    def _set_reasonable_window_size(self):
        try:
            geo = QGuiApplication.primaryScreen().availableGeometry()
            w = int(geo.width() * 0.72)
            h = int(geo.height() * 0.78)
            self.resize(max(860, w), max(600, h))
        except Exception:
            self.resize(920, 700)

    # ---------------------------
    # Tabs
    # ---------------------------
    def _build_tab_connection(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        row = QHBoxLayout()
        row.addWidget(QLabel("COM:"))
        self.cb_port = QComboBox()
        row.addWidget(self.cb_port)

        self.btn_refresh = QPushButton("Aktualisieren")
        row.addWidget(self.btn_refresh)

        row.addSpacing(10)
        row.addWidget(QLabel("Intervall (ms):"))
        self.sb_interval = QSpinBox()
        self.sb_interval.setRange(10, 600000)
        self.sb_interval.setValue(500)
        row.addWidget(self.sb_interval)

        row.addStretch(1)

        self.btn_connect = QPushButton("Verbinden")
        self.btn_connect_rd = QPushButton("Verbinden R+D")
        self.btn_disconnect = QPushButton("Trennen")
        self.btn_disconnect.setEnabled(False)

        row.addWidget(self.btn_connect)
        row.addWidget(self.btn_connect_rd)
        row.addWidget(self.btn_disconnect)

        layout.addLayout(row)
        self.tabs.addTab(tab, "Verbindung")

        self.btn_refresh.clicked.connect(self.refresh_ports)
        self.btn_connect.clicked.connect(self.connect_normal)
        self.btn_connect_rd.clicked.connect(self.connect_rd)
        self.btn_disconnect.clicked.connect(self.disconnect_port)

    def _build_tab_live(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        self.tbl_live = QTableWidget()
        self.tbl_live.setColumnCount(2)
        self.tbl_live.setHorizontalHeaderLabels(["Parameter", "Wert"])
        self.tbl_live.verticalHeader().setVisible(False)
        self.tbl_live.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_live.setSelectionMode(QTableWidget.NoSelection)
        self.tbl_live.setAlternatingRowColors(True)

        hdr = self.tbl_live.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)

        f = self.tbl_live.font()
        f.setPointSize(max(8, f.pointSize() - 1))
        self.tbl_live.setFont(f)

        layout.addWidget(self.tbl_live)
        self.tabs.addTab(tab, "Live")

    def _build_tab_temperatures(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel("Hold (s):"))
        self.sb_hold = QDoubleSpinBox()
        self.sb_hold.setRange(0.1, 60.0)
        self.sb_hold.setValue(5.0)
        self.sb_hold.setSingleStep(0.5)
        ctrl.addWidget(self.sb_hold)

        ctrl.addSpacing(10)
        ctrl.addWidget(QLabel("Skala (°C):"))
        self.sb_scale = QDoubleSpinBox()
        self.sb_scale.setRange(0.1, 100.0)
        self.sb_scale.setValue(5.0)
        self.sb_scale.setSingleStep(0.5)
        ctrl.addWidget(self.sb_scale)

        ctrl.addSpacing(10)
        self.btn_reset_max = QPushButton("Reset Max")
        ctrl.addWidget(self.btn_reset_max)

        ctrl.addStretch(1)
        layout.addLayout(ctrl)

        # ---- FIX: 7 columns, Peak+ and Peak- both stretch equally ----
        self.tbl_temp = QTableWidget()
        self.tbl_temp.setColumnCount(7)
        self.tbl_temp.setHorizontalHeaderLabels([
            "Temp", "Mean", "Dev", "Max +", "Max -", "Peak +", "Peak -"
        ])
        self.tbl_temp.setRowCount(len(TEMP_KEYS_ALL))
        self.tbl_temp.verticalHeader().setVisible(False)
        self.tbl_temp.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_temp.setSelectionMode(QTableWidget.NoSelection)
        self.tbl_temp.setAlternatingRowColors(True)

        f = self.tbl_temp.font()
        f.setPointSize(max(8, f.pointSize() - 1))
        self.tbl_temp.setFont(f)

        hdr = self.tbl_temp.horizontalHeader()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for c in (1, 2, 3, 4):
            hdr.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.Stretch)
        hdr.setSectionResizeMode(6, QHeaderView.Stretch)

        self.temp_widgets = {}
        for r, key in enumerate(TEMP_KEYS_ALL):
            self.tbl_temp.setItem(r, 0, QTableWidgetItem(key))
            for c in (1, 2, 3, 4):
                self.tbl_temp.setItem(r, c, QTableWidgetItem("-"))

            bp = QProgressBar()
            bn = QProgressBar()
            for b in (bp, bn):
                b.setRange(0, 1000)
                b.setFormat("")
                b.setTextVisible(False)
                b.setMaximumHeight(14)
                b.setMinimumWidth(220)  # optional: looks much nicer

            self.tbl_temp.setCellWidget(r, 5, bp)
            self.tbl_temp.setCellWidget(r, 6, bn)
            self.temp_widgets[key] = (bp, bn)

        layout.addWidget(self.tbl_temp)
        self.tabs.addTab(tab, "Temperaturen")

        self.btn_reset_max.clicked.connect(self.reset_maxima)

    def _build_tab_terminal(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        top = QHBoxLayout()
        top.addWidget(QLabel("Befehl:"))

        self.le_cmd = HistoryLineEdit()
        self.le_cmd.setPlaceholderText("Befehl eingeben (\\r\\n wird automatisch angehängt)")
        top.addWidget(self.le_cmd, 1)

        self.btn_send = QPushButton("Senden")
        top.addWidget(self.btn_send)

        self.btn_clear = QPushButton("Clear")
        top.addWidget(self.btn_clear)

        layout.addLayout(top)

        self.te_rx = QPlainTextEdit()
        self.te_rx.setReadOnly(True)
        self.te_rx.setLineWrapMode(QPlainTextEdit.NoWrap)
        fm = self.te_rx.fontMetrics()
        self.te_rx.setMinimumHeight(fm.lineSpacing() * 6)
        layout.addWidget(self.te_rx, 1)

        self.tabs.addTab(tab, "Terminal")

        self.btn_send.clicked.connect(self.send_terminal_cmd)
        self.btn_clear.clicked.connect(lambda: self.te_rx.clear())
        self.le_cmd.returnPressed.connect(self.send_terminal_cmd)

    # ---------------------------
    # Connection / control
    # ---------------------------
    def refresh_ports(self):
        self.cb_port.clear()
        ports = [p.device for p in list_ports.comports()]
        self.cb_port.addItems(ports)

    def _start_worker(self, mode: str):
        if self.worker.isRunning():
            return
        port = self.cb_port.currentText().strip()
        if not port:
            QMessageBox.warning(self, "Hinweis", "Bitte COM-Port auswählen.")
            return

        self.btn_connect.setEnabled(False)
        self.btn_connect_rd.setEnabled(False)
        self.btn_disconnect.setEnabled(True)

        self.t0_mono = time.monotonic()
        self.current_mode = mode
        self.reset_stats_all()
        self.setup_live_table(mode)

        self.worker.configure(port, self.sb_interval.value(), mode)
        self.worker.start()

    def connect_normal(self):
        self._start_worker("normal")

    def connect_rd(self):
        self._start_worker("rd")

    def disconnect_port(self):
        if self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(2000)

        self.btn_connect.setEnabled(True)
        self.btn_connect_rd.setEnabled(True)
        self.btn_disconnect.setEnabled(False)
        self.lbl_status.setText("Nicht verbunden")
        self.current_mode = None

    def closeEvent(self, event):
        try:
            self.disconnect_port()
        finally:
            self.logger_normal.close()
            self.logger_rd.close()
        super().closeEvent(event)

    # ---------------------------
    # Live table setup
    # ---------------------------
    def setup_live_table(self, mode: str):
        names = FIELD_NAMES_NORMAL if mode == "normal" else FIELD_NAMES_RD
        self.tbl_live.setRowCount(len(names))
        for r, name in enumerate(names):
            self.tbl_live.setItem(r, 0, QTableWidgetItem(name))
            self.tbl_live.setItem(r, 1, QTableWidgetItem("-"))

    # ---------------------------
    # Temperature stats helpers
    # ---------------------------
    def reset_stats_all(self):
        self.stats = {k: init_tempstats(50) for k in TEMP_KEYS_ALL}
        for r, key in enumerate(TEMP_KEYS_ALL):
            for c in (1, 2, 3, 4):
                self.tbl_temp.item(r, c).setText("-")
            bp, bn = self.temp_widgets[key]
            bp.setValue(0)
            bn.setValue(0)

    def reset_maxima(self):
        for key in TEMP_KEYS_ALL:
            self.stats[key].latched_pos = 0.0
            self.stats[key].latched_neg = 0.0
        self.update_temp_table_display(mode=self.current_mode, values=None)

    # ---------------------------
    # Terminal
    # ---------------------------
    def send_terminal_cmd(self):
        cmd = self.le_cmd.text().strip()
        if not cmd:
            return
        if not self.worker.isRunning():
            self.te_rx.appendPlainText("<< Nicht verbunden")
            return

        self.le_cmd.add_history(cmd)
        self.le_cmd.clear()
        self.worker.enqueue_manual(cmd)

    def on_term_line(self, s: str):
        self.te_rx.appendPlainText(s)
        cursor = self.te_rx.textCursor()
        cursor.movePosition(cursor.End)
        self.te_rx.setTextCursor(cursor)

    # ---------------------------
    # Worker callbacks
    # ---------------------------
    def on_connected(self, msg: str):
        self.lbl_status.setText(msg)

    def on_disconnected(self, msg: str):
        self.lbl_status.setText(msg)
        self.btn_connect.setEnabled(True)
        self.btn_connect_rd.setEnabled(True)
        self.btn_disconnect.setEnabled(False)

    def on_warn(self, msg: str):
        self.lbl_status.setText(msg)

    def on_error(self, msg: str):
        QMessageBox.critical(self, "Fehler", msg)
        self.lbl_status.setText("Fehler: " + msg)
        self.btn_connect.setEnabled(True)
        self.btn_connect_rd.setEnabled(True)
        self.btn_disconnect.setEnabled(False)

    def on_data(self, payload: object):
        mode = payload.get("mode")
        values = payload.get("values", [])

        if values:
            for r in range(self.tbl_live.rowCount()):
                if r < len(values):
                    self.tbl_live.item(r, 1).setText(values[r])

        if self.t0_mono is None:
            self.t0_mono = time.monotonic()
        sec = time.monotonic() - self.t0_mono

        try:
            if mode == "normal" and len(values) == 14:
                self.logger_normal.log_row(timestamp_de(), sec, values)
            elif mode == "rd" and len(values) == 17:
                self.logger_rd.log_row(timestamp_de(), sec, values)
        except Exception as e:
            self.lbl_status.setText(f"Logging-Fehler: {e}")

        self.update_temp_table_display(mode=mode, values=values)

    def update_temp_table_display(self, mode: str | None, values: list[str] | None):
        hold_s = float(self.sb_hold.value())
        scale = float(self.sb_scale.value())
        bar_max = 1000
        now = time.monotonic()

        if mode == "normal" and values and len(values) == 14:
            for key, idx in TEMP_IDX_NORMAL.items():
                try:
                    t = float(values[idx])
                except Exception:
                    continue
                update_tempstats(self.stats[key], t, hold_s, now)

        if mode == "rd" and values and len(values) == 17:
            for key, idx in TEMP_IDX_RD.items():
                try:
                    t = float(values[idx])
                except Exception:
                    continue
                update_tempstats(self.stats[key], t, hold_s, now)

        for r, key in enumerate(TEMP_KEYS_ALL):
            st = self.stats[key]

            if mode == "normal" and key in ["Reso", "Vanadat"]:
                self.tbl_temp.item(r, 1).setText("-")
                self.tbl_temp.item(r, 2).setText("-")
                bp, bn = self.temp_widgets[key]
                bp.setValue(0)
                bn.setValue(0)
                self.tbl_temp.item(r, 3).setText(float_to_de_str(st.latched_pos, 3))
                self.tbl_temp.item(r, 4).setText(float_to_de_str(st.latched_neg, 3))
                continue

            self.tbl_temp.item(r, 1).setText(float_to_de_str(st.mean, 3))
            self.tbl_temp.item(r, 2).setText(float_to_de_str(st.dev, 3))
            self.tbl_temp.item(r, 3).setText(float_to_de_str(st.latched_pos, 3))
            self.tbl_temp.item(r, 4).setText(float_to_de_str(st.latched_neg, 3))

            pos = max(st.bar_peak_pos, 0.0)
            neg = abs(min(st.bar_peak_neg, 0.0))
            pos_val = int(min(1.0, pos / max(scale, 1e-6)) * bar_max)
            neg_val = int(min(1.0, neg / max(scale, 1e-6)) * bar_max)

            bp, bn = self.temp_widgets[key]
            bp.setValue(pos_val)
            bn.setValue(neg_val)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
