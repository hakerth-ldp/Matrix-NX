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
    QProgressBar, QMessageBox
)

# ---------------------------
# Config / constants
# ---------------------------
BAUDRATE = 115200
DATABITS = serial.EIGHTBITS
STOPBITS = serial.STOPBITS_ONE
PARITY = serial.PARITY_NONE
READ_TIMEOUT_S = 2.0
WRITE_TIMEOUT_S = 2.0

LOG_DIR = Path(r"C:\Coherent")
LOG_PREFIX = "matrixNX"

FIELD_NAMES = [
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

TEMP_FIELD_IDX = {
    "Housing": 3,
    "Diode": 4,
    "SHG": 5,
    "THG": 6,
}

# ---------------------------
# Helper: German-friendly formatting
# ---------------------------
def float_to_de_str(x: float, decimals: int = 3) -> str:
    s = f"{x:.{decimals}f}"
    return s.replace(".", ",")

def timestamp_de() -> str:
    return datetime.now().strftime("%d.%m.%Y %H:%M:%S")

# ---------------------------
# Logger (one file per day)
# ---------------------------
class DailyCsvLogger:
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

        # Close old file
        if self._fh:
            try:
                self._fh.flush()
                self._fh.close()
            except Exception:
                pass

        self._current_day = today
        fname = f"{LOG_PREFIX}_{today.isoformat()}.csv"
        fpath = self.base_dir / fname
        is_new = not fpath.exists()

        # newline="" is important for CSV on Windows
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
        """
        Logs: timestamp + seconds + first 8 values (Status..Operation Hours),
        but temps are numeric strings -> convert to DE decimal comma.
        """
        self._open_for_today()

        # first 8 fields: 0..7
        status = fields14[0]
        warnings = fields14[1]
        faults = fields14[2]

        # temps (3..6)
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

# ---------------------------
# Rolling mean + peak hold
# ---------------------------
@dataclass
class TempStats:
    window: deque
    sum_: float
    mean: float
    dev: float
    peak_pos: float
    peak_neg: float
    t_peak_pos: float
    t_peak_neg: float

def init_tempstats(n: int = 50) -> TempStats:
    return TempStats(
        window=deque(maxlen=n),
        sum_=0.0,
        mean=0.0,
        dev=0.0,
        peak_pos=0.0,
        peak_neg=0.0,
        t_peak_pos=0.0,
        t_peak_neg=0.0,
    )

def update_tempstats(stats: TempStats, value: float, hold_s: float, now_mono: float) -> TempStats:
    # update sliding window + sum
    if len(stats.window) == stats.window.maxlen:
        old = stats.window[0]
        stats.sum_ -= old
    stats.window.append(value)
    stats.sum_ += value

    stats.mean = stats.sum_ / max(1, len(stats.window))
    stats.dev = value - stats.mean

    # update peaks (pos/neg) and hold behavior
    if stats.dev > stats.peak_pos:
        stats.peak_pos = stats.dev
        stats.t_peak_pos = now_mono

    if stats.dev < stats.peak_neg:
        stats.peak_neg = stats.dev
        stats.t_peak_neg = now_mono

    # If hold expired, reset peak display baseline to current deviation
    if now_mono - stats.t_peak_pos > hold_s:
        stats.peak_pos = max(stats.dev, 0.0)
        stats.t_peak_pos = now_mono

    if now_mono - stats.t_peak_neg > hold_s:
        stats.peak_neg = min(stats.dev, 0.0)
        stats.t_peak_neg = now_mono

    return stats

# ---------------------------
# Serial worker thread
# ---------------------------
class SerialWorker(QThread):
    connected = Signal(str)
    disconnected = Signal(str)
    error = Signal(str)
    data = Signal(list)  # 14 fields as list[str]
    timeout = Signal()

    def __init__(self):
        super().__init__()
        self._port_name = None
        self._interval_ms = 500
        self._stop = False
        self._ser = None

    def configure(self, port_name: str, interval_ms: int):
        self._port_name = port_name
        self._interval_ms = max(10, int(interval_ms))

    def stop(self):
        self._stop = True

    def _close_serial(self):
        if self._ser:
            try:
                self._ser.close()
            except Exception:
                pass
            self._ser = None

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
            # Clean start
            try:
                self._ser.reset_input_buffer()
                self._ser.reset_output_buffer()
            except Exception:
                pass

            self.connected.emit(f"Verbunden: {self._port_name}")
        except Exception as e:
            self.error.emit(f"Verbindung fehlgeschlagen: {e}")
            self._close_serial()
            return

        interval_s = self._interval_ms / 1000.0
        next_send = time.monotonic()

        self._stop = False
        while not self._stop:
            now = time.monotonic()
            if now < next_send:
                time.sleep(min(0.05, next_send - now))
                continue

            # Send command
            try:
                self._ser.write(b"All?\r\n")
            except Exception as e:
                self.error.emit(f"Senden fehlgeschlagen: {e}")
                break

            # Read first line (values)
            try:
                line1 = self._ser.readline()  # up to \n or timeout
            except Exception as e:
                self.error.emit(f"Lesen fehlgeschlagen: {e}")
                break

            if not line1:
                self.timeout.emit()
                # schedule next send (2A behavior)
                next_send = max(next_send + interval_s, time.monotonic())
                continue

            line1_s = line1.decode(errors="replace").strip()
            # read second line (OK), but don't depend on it too hard
            try:
                line2 = self._ser.readline()
                if line2:
                    _ = line2.decode(errors="replace").strip()
            except Exception:
                pass

            parts = [p.strip() for p in line1_s.split(",")]
            if len(parts) != 14:
                # sometimes devices may echo extra text; ignore bad frames
                # (no timeout logging requested)
                next_send = max(next_send + interval_s, time.monotonic())
                continue

            self.data.emit(parts)

            # schedule next send (2A behavior)
            next_send = max(next_send + interval_s, time.monotonic())

        self._close_serial()
        self.disconnected.emit("Getrennt")

# ---------------------------
# GUI
# ---------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Matrix NX RS232 Monitor (All?)")

        self.worker = SerialWorker()
        self.worker.connected.connect(self.on_connected)
        self.worker.disconnected.connect(self.on_disconnected)
        self.worker.error.connect(self.on_error)
        self.worker.data.connect(self.on_data)
        self.worker.timeout.connect(self.on_timeout)

        self.logger = DailyCsvLogger(LOG_DIR)
        self.t0_mono = None

        self.stats = {
            "Housing": init_tempstats(50),
            "Diode": init_tempstats(50),
            "SHG": init_tempstats(50),
            "THG": init_tempstats(50),
        }

        # UI
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        # Connection box
        conn_box = QGroupBox("Verbindung")
        conn_layout = QHBoxLayout(conn_box)

        self.cb_port = QComboBox()
        self.btn_refresh = QPushButton("Ports aktualisieren")
        self.btn_connect = QPushButton("Verbinden")
        self.btn_disconnect = QPushButton("Trennen")
        self.btn_disconnect.setEnabled(False)

        conn_layout.addWidget(QLabel("COM:"))
        conn_layout.addWidget(self.cb_port)
        conn_layout.addWidget(self.btn_refresh)
        conn_layout.addSpacing(10)
        conn_layout.addWidget(QLabel("Intervall (ms):"))
        self.sb_interval = QSpinBox()
        self.sb_interval.setRange(10, 600000)
        self.sb_interval.setValue(500)
        conn_layout.addWidget(self.sb_interval)

        conn_layout.addWidget(self.btn_connect)
        conn_layout.addWidget(self.btn_disconnect)

        self.lbl_conn = QLabel("Nicht verbunden")
        self.lbl_conn.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        layout.addWidget(conn_box)
        layout.addWidget(self.lbl_conn)

        # Live values
        live_box = QGroupBox("Live-Daten (letzte Antwort)")
        live_grid = QGridLayout(live_box)
        self.live_labels = []
        for i, name in enumerate(FIELD_NAMES):
            live_grid.addWidget(QLabel(name + ":"), i, 0)
            val = QLabel("-")
            val.setTextInteractionFlags(Qt.TextSelectableByMouse)
            live_grid.addWidget(val, i, 1)
            self.live_labels.append(val)
        layout.addWidget(live_box)

        # Temperature analysis
        temp_box = QGroupBox("Temperatur-Analyse (Mittelwert über 50 Messungen + Abweichung/Peaks)")
        temp_grid = QGridLayout(temp_box)

        temp_grid.addWidget(QLabel("Temp"), 0, 0)
        temp_grid.addWidget(QLabel("Hold (s)"), 0, 1)
        temp_grid.addWidget(QLabel("Mittelwert (°C)"), 0, 2)
        temp_grid.addWidget(QLabel("Abw. aktuell (°C)"), 0, 3)
        temp_grid.addWidget(QLabel("Peak + (Balken)"), 0, 4)
        temp_grid.addWidget(QLabel("Peak - (Balken)"), 0, 5)

        self.hold_spin = {}
        self.mean_lbl = {}
        self.dev_lbl = {}
        self.bar_pos = {}
        self.bar_neg = {}

        # Common bar scale
        scale_row = 1
        temp_grid.addWidget(QLabel("Balken-Skala (°C):"), scale_row, 0)
        self.sb_scale = QDoubleSpinBox()
        self.sb_scale.setRange(0.1, 100.0)
        self.sb_scale.setValue(5.0)
        self.sb_scale.setSingleStep(0.5)
        temp_grid.addWidget(self.sb_scale, scale_row, 1)
        temp_grid.addWidget(QLabel("→ Peak-Balken zeigen |Abweichung| bis zur Skala."), scale_row, 2, 1, 4)

        start_row = 2
        for r, key in enumerate(["Housing", "Diode", "SHG", "THG"], start=start_row):
            temp_grid.addWidget(QLabel(key), r, 0)

            hs = QDoubleSpinBox()
            hs.setRange(0.1, 60.0)
            hs.setValue(5.0)
            hs.setSingleStep(0.5)
            self.hold_spin[key] = hs
            temp_grid.addWidget(hs, r, 1)

            ml = QLabel("-")
            dl = QLabel("-")
            self.mean_lbl[key] = ml
            self.dev_lbl[key] = dl
            temp_grid.addWidget(ml, r, 2)
            temp_grid.addWidget(dl, r, 3)

            bp = QProgressBar()
            bn = QProgressBar()
            bp.setRange(0, 1000)
            bn.setRange(0, 1000)
            bp.setFormat("")  # "nur Balken"
            bn.setFormat("")
            self.bar_pos[key] = bp
            self.bar_neg[key] = bn
            temp_grid.addWidget(bp, r, 4)
            temp_grid.addWidget(bn, r, 5)

        layout.addWidget(temp_box)

        # Events
        self.btn_refresh.clicked.connect(self.refresh_ports)
        self.btn_connect.clicked.connect(self.connect_port)
        self.btn_disconnect.clicked.connect(self.disconnect_port)

        self.refresh_ports()

    def refresh_ports(self):
        self.cb_port.clear()
        ports = [p.device for p in list_ports.comports()]
        self.cb_port.addItems(ports)

    def connect_port(self):
        if self.worker.isRunning():
            return
        port = self.cb_port.currentText().strip()
        if not port:
            QMessageBox.warning(self, "Hinweis", "Bitte COM-Port auswählen.")
            return

        self.btn_connect.setEnabled(False)
        self.btn_disconnect.setEnabled(True)

        self.t0_mono = time.monotonic()
        self.worker.configure(port, self.sb_interval.value())
        self.worker.start()

    def disconnect_port(self):
        if self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(2000)
        self.btn_connect.setEnabled(True)
        self.btn_disconnect.setEnabled(False)
        self.lbl_conn.setText("Nicht verbunden")

    def closeEvent(self, event):
        try:
            self.disconnect_port()
        finally:
            self.logger.close()
        super().closeEvent(event)

    def on_connected(self, msg: str):
        self.lbl_conn.setText(msg)

    def on_disconnected(self, msg: str):
        self.lbl_conn.setText(msg)
        self.btn_connect.setEnabled(True)
        self.btn_disconnect.setEnabled(False)

    def on_error(self, msg: str):
        QMessageBox.critical(self, "Fehler", msg)
        self.lbl_conn.setText("Fehler: " + msg)
        self.btn_connect.setEnabled(True)
        self.btn_disconnect.setEnabled(False)

    def on_timeout(self):
        # kein Logging gewünscht, nur optional Anzeige
        self.lbl_conn.setText(f"Verbunden (Timeout beim Lesen, {READ_TIMEOUT_S:.1f}s)")

    def on_data(self, fields: list[str]):
        # Update live labels
        for i, v in enumerate(fields):
            self.live_labels[i].setText(v)

        # Logging (first 8 values)
        if self.t0_mono is None:
            self.t0_mono = time.monotonic()
        sec = time.monotonic() - self.t0_mono
        try:
            self.logger.log_row(timestamp_de(), sec, fields)
        except Exception as e:
            # logging errors should not kill acquisition
            self.lbl_conn.setText(f"Logging-Fehler: {e}")

        # Temperature analysis
        now = time.monotonic()
        scale = float(self.sb_scale.value())
        bar_max = 1000  # progressbar internal range

        for key, idx in TEMP_FIELD_IDX.items():
            try:
                t = float(fields[idx])
            except ValueError:
                continue

            hold = float(self.hold_spin[key].value())
            st = self.stats[key]
            update_tempstats(st, t, hold, now)

            self.mean_lbl[key].setText(float_to_de_str(st.mean, 3))
            self.dev_lbl[key].setText(float_to_de_str(st.dev, 3))

            # display held peaks as bars (only bars)
            pos = max(st.peak_pos, 0.0)
            neg = abs(min(st.peak_neg, 0.0))

            # scale to progressbar
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
    w.resize(900, 900)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
