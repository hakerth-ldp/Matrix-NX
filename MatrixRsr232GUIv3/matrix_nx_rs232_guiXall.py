
import sys
import time
import csv
from dataclasses import dataclass
from collections import deque
from datetime import datetime, date
from pathlib import Path

import serial
from serial.tools import list_ports

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QSpinBox, QDoubleSpinBox, QGroupBox,
    QProgressBar, QMessageBox, QScrollArea
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

# Temp index mapping (values list differs by mode)
TEMP_KEYS_ALL = ["Reso", "Vanadat", "Diode", "Housing", "SHG", "THG"]

TEMP_IDX_NORMAL = {
    "Housing": 3,
    "Diode": 4,
    "SHG": 5,
    "THG": 6,
}

# In RD combined list: first 10 are TEMPeratures
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
    """For logging: if it is a float-like value, write decimal comma; else keep as-is."""
    s = v.strip()
    if s == "":
        return ""
    # do not convert time-like strings "45:03:14"
    if ":" in s:
        return s
    try:
        x = float(s)
        return float_to_de_str(x, 3)
    except ValueError:
        return s

# ---------------------------
# Daily CSV loggers
# ---------------------------
class DailyCsvLoggerNormal:
    """Normal mode: keep your original behavior: timestamp + sec + first 8 values (Status..Operation Hours)"""
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
#  - Bars: peak hold for hold_s seconds (like before)
#  - Numbers: latched max/min deviation until reset button
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

    latched_pos: float   # max positive deviation since reset
    latched_neg: float   # most negative deviation since reset

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
    # sliding window
    if len(stats.window) == stats.window.maxlen:
        stats.sum_ -= stats.window[0]
    stats.window.append(value)
    stats.sum_ += value

    stats.mean = stats.sum_ / max(1, len(stats.window))
    stats.dev = value - stats.mean

    # latched max/min (until reset)
    if stats.dev > stats.latched_pos:
        stats.latched_pos = stats.dev
    if stats.dev < stats.latched_neg:
        stats.latched_neg = stats.dev

    # bar peaks with hold (as in your first version)
    if stats.dev > stats.bar_peak_pos:
        stats.bar_peak_pos = stats.dev
        stats.t_bar_pos = now_mono
    if stats.dev < stats.bar_peak_neg:
        stats.bar_peak_neg = stats.dev
        stats.t_bar_neg = now_mono

    # if hold expired, reset bar peaks baseline to current deviation
    if now_mono - stats.t_bar_pos > hold_s:
        stats.bar_peak_pos = max(stats.dev, 0.0)
        stats.t_bar_pos = now_mono
    if now_mono - stats.t_bar_neg > hold_s:
        stats.bar_peak_neg = min(stats.dev, 0.0)
        stats.t_bar_neg = now_mono

# ---------------------------
# Serial worker
# ---------------------------
class SerialWorker(QThread):
    connected = Signal(str)
    disconnected = Signal(str)
    error = Signal(str)
    warn = Signal(str)
    data = Signal(object)   # {"mode": "normal"/"rd", "values": [...], ...}

    def __init__(self):
        super().__init__()
        self._port_name = None
        self._interval_ms = 500
        self._stop = False
        self._ser = None
        self._mode = "normal"  # "normal" or "rd"

    def configure(self, port_name: str, interval_ms: int, mode: str):
        self._port_name = port_name
        self._interval_ms = max(10, int(interval_ms))
        self._mode = mode

    def stop(self):
        self._stop = True

    def _close_serial(self):
        if self._ser:
            try:
                self._ser.close()
            except Exception:
                pass
            self._ser = None

    def _send_cmd(self, cmd: str):
        self._ser.write((cmd + "\r\n").encode("ascii", errors="replace"))

    def _read_data_and_ok(self, expected_fields: int):
        """
        Reads until timeout:
        - tries to capture one data line with expected_fields
        - tries to see an OK line
        Returns (data_parts_or_None, ok_bool)
        Soft: if OK missing but data exists -> ok=False, data returned
        """
        deadline = time.monotonic() + READ_TIMEOUT_S
        data_parts = None
        ok = False

        while time.monotonic() < deadline and not self._stop:
            raw = self._ser.readline()
            if not raw:
                break
            s = raw.decode(errors="replace").strip()
            if not s:
                continue
            if s.upper() == "OK":
                ok = True
                if data_parts is not None:
                    return data_parts, True
                continue

            parts = [p.strip() for p in s.split(",")]
            if len(parts) == expected_fields:
                data_parts = parts
                if ok:
                    return data_parts, True
                # else keep reading for OK until timeout

        return data_parts, ok

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
                timeout=READ_TIMEOUT_S,
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

        # RD handshake
        if self._mode == "rd":
            try:
                self._send_cmd(CMD_PASS)
                _, okp = self._read_data_and_ok(expected_fields=0)  # password returns only OK typically
                # expected_fields=0 won't match; so treat OK only:
                # We'll do a dedicated OK wait:
            except Exception:
                okp = False

            # Better OK wait (soft)
            try:
                deadline = time.monotonic() + READ_TIMEOUT_S
                ok_seen = False
                while time.monotonic() < deadline and not self._stop:
                    raw = self._ser.readline()
                    if not raw:
                        break
                    s = raw.decode(errors="replace").strip()
                    if s.upper() == "OK":
                        ok_seen = True
                        break
                if not ok_seen:
                    self.warn.emit("Warnung: Kein OK auf Passwort empfangen – versuche trotzdem weiter.")
            except Exception:
                self.warn.emit("Warnung: Fehler beim Warten auf OK (Passwort) – versuche trotzdem weiter.")

        interval_s = self._interval_ms / 1000.0
        next_cycle = time.monotonic()

        self._stop = False
        while not self._stop:
            now = time.monotonic()
            if now < next_cycle:
                time.sleep(min(0.05, next_cycle - now))
                continue

            if self._mode == "normal":
                # All?
                try:
                    self._send_cmd(CMD_ALL)
                    parts, ok = self._read_data_and_ok(expected_fields=14)
                    if parts is not None:
                        if not ok:
                            self.warn.emit("Warnung: OK fehlt (All?)")
                        self.data.emit({"mode": "normal", "values": parts, "ok": ok})
                except Exception as e:
                    self.error.emit(f"Kommunikationsfehler: {e}")
                    break

            else:
                # RD cycle: TEMPeratures + OTHers
                temps = None
                others = None
                ok_t = True
                ok_o = True
                try:
                    self._send_cmd(CMD_TEMPS)
                    temps, ok_t = self._read_data_and_ok(expected_fields=10)

                    self._send_cmd(CMD_OTHERS)
                    others, ok_o = self._read_data_and_ok(expected_fields=7)

                    if temps is None and others is None:
                        # nothing received this cycle
                        pass
                    else:
                        if temps is None:
                            temps = [""] * 10
                        if others is None:
                            others = [""] * 7

                        if not ok_t:
                            self.warn.emit("Warnung: OK fehlt (TEMPeratures)")
                        if not ok_o:
                            self.warn.emit("Warnung: OK fehlt (OTHers)")

                        combined = temps + others  # 17 fields
                        self.data.emit({"mode": "rd", "values": combined, "ok_t": ok_t, "ok_o": ok_o})
                except Exception as e:
                    self.error.emit(f"Kommunikationsfehler: {e}")
                    break

            # 2A behavior: next cycle is scheduled, no overlap
            next_cycle = max(next_cycle + interval_s, time.monotonic())

        self._close_serial()
        self.disconnected.emit("Getrennt")

# ---------------------------
# GUI
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

        # loggers
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        self.logger_normal = DailyCsvLoggerNormal(LOG_DIR)
        self.logger_rd = DailyCsvLoggerRD(LOG_DIR)

        self.t0_mono = None
        self.current_mode = None

        self.stats = {k: init_tempstats(50) for k in TEMP_KEYS_ALL}

        # Build UI in a scroll area (fix for clipped bottom)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        scroll.setWidget(container)
        self.setCentralWidget(scroll)

        layout = QVBoxLayout(container)

        # Connection box
        conn_box = QGroupBox("Verbindung")
        conn_layout = QHBoxLayout(conn_box)

        self.cb_port = QComboBox()
        self.btn_refresh = QPushButton("Ports aktualisieren")

        self.btn_connect = QPushButton("Verbinden")
        self.btn_connect_rd = QPushButton("Verbinden R+D")
        self.btn_disconnect = QPushButton("Trennen")
        self.btn_disconnect.setEnabled(False)

        self.btn_reset_max = QPushButton("Reset Max-Abweichungen")
        self.btn_reset_max.setEnabled(True)

        conn_layout.addWidget(QLabel("COM:"))
        conn_layout.addWidget(self.cb_port)
        conn_layout.addWidget(self.btn_refresh)

        conn_layout.addSpacing(10)
        conn_layout.addWidget(QLabel("Intervall (ms):"))
        self.sb_interval = QSpinBox()
        self.sb_interval.setRange(10, 600000)
        self.sb_interval.setValue(500)
        conn_layout.addWidget(self.sb_interval)

        conn_layout.addSpacing(10)
        conn_layout.addWidget(self.btn_connect)
        conn_layout.addWidget(self.btn_connect_rd)
        conn_layout.addWidget(self.btn_disconnect)
        conn_layout.addWidget(self.btn_reset_max)

        self.lbl_conn = QLabel("Nicht verbunden")
        self.lbl_conn.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        layout.addWidget(conn_box)
        layout.addWidget(self.lbl_conn)

        # Live data box (dynamic fields)
        self.live_box = QGroupBox("Live-Daten (letzte Antwort)")
        self.live_grid = QGridLayout(self.live_box)
        self.name_labels = []
        self.value_labels = []

        # Create max rows (17) - names will be set per mode
        max_fields = max(len(FIELD_NAMES_NORMAL), len(FIELD_NAMES_RD))
        for i in range(max_fields):
            name = QLabel("-")
            val = QLabel("-")
            val.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.live_grid.addWidget(name, i, 0)
            self.live_grid.addWidget(val, i, 1)
            self.name_labels.append(name)
            self.value_labels.append(val)

        layout.addWidget(self.live_box)

        # Temperature analysis box
        temp_box = QGroupBox("Temperatur-Analyse (Mittelwert 50 | Balken=Hold | Zahlen=Max bis Reset)")
        temp_grid = QGridLayout(temp_box)

        temp_grid.addWidget(QLabel("Temp"), 0, 0)
        temp_grid.addWidget(QLabel("Hold (s)"), 0, 1)
        temp_grid.addWidget(QLabel("Mittelwert (°C)"), 0, 2)
        temp_grid.addWidget(QLabel("Abw. aktuell (°C)"), 0, 3)
        temp_grid.addWidget(QLabel("Max + (°C)"), 0, 4)
        temp_grid.addWidget(QLabel("Max - (°C)"), 0, 5)
        temp_grid.addWidget(QLabel("Peak + (Balken)"), 0, 6)
        temp_grid.addWidget(QLabel("Peak - (Balken)"), 0, 7)

        # Common bar scale
        temp_grid.addWidget(QLabel("Balken-Skala (°C):"), 1, 0)
        self.sb_scale = QDoubleSpinBox()
        self.sb_scale.setRange(0.1, 100.0)
        self.sb_scale.setValue(5.0)
        self.sb_scale.setSingleStep(0.5)
        temp_grid.addWidget(self.sb_scale, 1, 1)
        temp_grid.addWidget(QLabel("→ Balken zeigen |Abweichung| bis zur Skala."), 1, 2, 1, 6)

        self.hold_spin = {}
        self.mean_lbl = {}
        self.dev_lbl = {}
        self.maxp_lbl = {}
        self.maxn_lbl = {}
        self.bar_pos = {}
        self.bar_neg = {}

        start_row = 2
        for r, key in enumerate(TEMP_KEYS_ALL, start=start_row):
            temp_grid.addWidget(QLabel(key), r, 0)

            hs = QDoubleSpinBox()
            hs.setRange(0.1, 60.0)
            hs.setValue(5.0)
            hs.setSingleStep(0.5)
            self.hold_spin[key] = hs
            temp_grid.addWidget(hs, r, 1)

            ml = QLabel("-")
            dl = QLabel("-")
            mp = QLabel("0,000")
            mn = QLabel("0,000")
            self.mean_lbl[key] = ml
            self.dev_lbl[key] = dl
            self.maxp_lbl[key] = mp
            self.maxn_lbl[key] = mn

            temp_grid.addWidget(ml, r, 2)
            temp_grid.addWidget(dl, r, 3)
            temp_grid.addWidget(mp, r, 4)
            temp_grid.addWidget(mn, r, 5)

            bp = QProgressBar()
            bn = QProgressBar()
            bp.setRange(0, 1000)
            bn.setRange(0, 1000)
            bp.setFormat("")  # nur Balken
            bn.setFormat("")
            self.bar_pos[key] = bp
            self.bar_neg[key] = bn
            temp_grid.addWidget(bp, r, 6)
            temp_grid.addWidget(bn, r, 7)

        layout.addWidget(temp_box)

        # Events
        self.btn_refresh.clicked.connect(self.refresh_ports)
        self.btn_connect.clicked.connect(self.connect_normal)
        self.btn_connect_rd.clicked.connect(self.connect_rd)
        self.btn_disconnect.clicked.connect(self.disconnect_port)
        self.btn_reset_max.clicked.connect(self.reset_maxima)

        self.refresh_ports()

        # Initial window size: a bit smaller; scroll handles the rest
        self.resize(950, 850)

    # ---------------------------
    # UI helpers
    # ---------------------------
    def refresh_ports(self):
        self.cb_port.clear()
        ports = [p.device for p in list_ports.comports()]
        self.cb_port.addItems(ports)

    def set_mode_and_live_fields(self, mode: str):
        self.current_mode = mode
        if mode == "normal":
            names = FIELD_NAMES_NORMAL
        else:
            names = FIELD_NAMES_RD

        for i in range(len(self.name_labels)):
            if i < len(names):
                self.name_labels[i].setText(names[i] + ":")
                self.value_labels[i].setText("-")
                self.name_labels[i].show()
                self.value_labels[i].show()
            else:
                self.name_labels[i].hide()
                self.value_labels[i].hide()

    def reset_stats_all(self):
        self.stats = {k: init_tempstats(50) for k in TEMP_KEYS_ALL}
        for k in TEMP_KEYS_ALL:
            self.mean_lbl[k].setText("-")
            self.dev_lbl[k].setText("-")
            self.maxp_lbl[k].setText("0,000")
            self.maxn_lbl[k].setText("0,000")
            self.bar_pos[k].setValue(0)
            self.bar_neg[k].setValue(0)

    def reset_maxima(self):
        for k in TEMP_KEYS_ALL:
            self.stats[k].latched_pos = 0.0
            self.stats[k].latched_neg = 0.0
            self.maxp_lbl[k].setText("0,000")
            self.maxn_lbl[k].setText("0,000")

    # ---------------------------
    # Connect / Disconnect
    # ---------------------------
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
        self.set_mode_and_live_fields(mode)
        self.reset_stats_all()

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
        self.lbl_conn.setText("Nicht verbunden")
        self.current_mode = None

    def closeEvent(self, event):
        try:
            self.disconnect_port()
        finally:
            self.logger_normal.close()
            self.logger_rd.close()
        super().closeEvent(event)

    # ---------------------------
    # Worker callbacks
    # ---------------------------
    def on_connected(self, msg: str):
        self.lbl_conn.setText(msg)

    def on_disconnected(self, msg: str):
        self.lbl_conn.setText(msg)
        self.btn_connect.setEnabled(True)
        self.btn_connect_rd.setEnabled(True)
        self.btn_disconnect.setEnabled(False)

    def on_warn(self, msg: str):
        # show soft warning in status line
        self.lbl_conn.setText(msg)

    def on_error(self, msg: str):
        QMessageBox.critical(self, "Fehler", msg)
        self.lbl_conn.setText("Fehler: " + msg)
        self.btn_connect.setEnabled(True)
        self.btn_connect_rd.setEnabled(True)
        self.btn_disconnect.setEnabled(False)

    def on_data(self, payload: object):
        mode = payload.get("mode")
        values = payload.get("values", [])

        # update live grid
        for i, v in enumerate(values):
            if i < len(self.value_labels):
                self.value_labels[i].setText(v)

        # logging
        if self.t0_mono is None:
            self.t0_mono = time.monotonic()
        sec = time.monotonic() - self.t0_mono

        try:
            if mode == "normal" and len(values) == 14:
                self.logger_normal.log_row(timestamp_de(), sec, values)
            elif mode == "rd" and len(values) == 17:
                self.logger_rd.log_row(timestamp_de(), sec, values)
        except Exception as e:
            self.lbl_conn.setText(f"Logging-Fehler: {e}")

        # temperature analysis
        now = time.monotonic()
        scale = float(self.sb_scale.value())
        bar_max = 1000

        if mode == "normal":
            # update only available temps (4)
            for key, idx in TEMP_IDX_NORMAL.items():
                try:
                    t = float(values[idx])
                except Exception:
                    continue
                hold = float(self.hold_spin[key].value())
                st = self.stats[key]
                update_tempstats(st, t, hold, now)
                self._update_temp_row(key, st, scale, bar_max)

            # temps not present in normal -> show "-" but keep last stats
            for key in ["Reso", "Vanadat"]:
                self.mean_lbl[key].setText("-")
                self.dev_lbl[key].setText("-")
                # keep max labels as they are (user wants reset-driven)
                self.bar_pos[key].setValue(0)
                self.bar_neg[key].setValue(0)

        elif mode == "rd":
            # update 6 temperatures from RD
            for key, idx in TEMP_IDX_RD.items():
                try:
                    t = float(values[idx])
                except Exception:
                    continue
                hold = float(self.hold_spin[key].value())
                st = self.stats[key]
                update_tempstats(st, t, hold, now)
                self._update_temp_row(key, st, scale, bar_max)

    def _update_temp_row(self, key: str, st: TempStats, scale: float, bar_max: int):
        # mean/dev
        self.mean_lbl[key].setText(float_to_de_str(st.mean, 3))
        self.dev_lbl[key].setText(float_to_de_str(st.dev, 3))

        # latched maxima numbers (until reset)
        self.maxp_lbl[key].setText(float_to_de_str(st.latched_pos, 3))
        self.maxn_lbl[key].setText(float_to_de_str(st.latched_neg, 3))

        # bars use held peaks
        pos = max(st.bar_peak_pos, 0.0)
        neg = abs(min(st.bar_peak_neg, 0.0))

        pos_val = int(min(1.0, pos / max(scale, 1e-6)) * bar_max)
        neg_val = int(min(1.0, neg / max(scale, 1e-6)) * bar_max)

        self.bar_pos[key].setValue(pos_val)
        self.bar_neg[key].setValue(neg_val)

# ---------------------------
# main
# ---------------------------
def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
