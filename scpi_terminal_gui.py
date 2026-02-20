import csv
import re
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import serial
    from serial.tools import list_ports
except ImportError:  # pragma: no cover
    serial = None
    list_ports = None

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None


XALL_GROUP_FIELDS = {
    "TEMPERATURES": [
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
    ],
    "STEPPER": [
        "Actual Stepper Position",
        "Actual Spot Number",
        "Spot Hours",
        "Spot Hours Remain",
        "Crystal Hours",
        "Spot Status",
        "Spot Warnings",
        "Spot Faults",
    ],
    "OTHERS": [
        "Fan Output Drive",
        "Laserdiode Current",
        "System Status Flags",
        "Scaled UV-Power",
        "Raw UV-Power",
        "Operation Hours",
        "LD Hours",
    ],
}

ALL_FIELDS = [
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

BYTESIZE_OPTIONS = ("5", "6", "7", "8")
PARITY_OPTIONS = ("None", "Even", "Odd", "Mark", "Space")
STOPBITS_OPTIONS = ("1", "1.5", "2")


class SCPITerminalApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SCPI Serial Terminal")
        self.root.geometry("1240x800")

        self.ser = None
        self.reader_thread = None
        self.stop_reader = threading.Event()
        self.pause_reader = threading.Event()
        self.serial_lock = threading.Lock()

        self.monitor_thread = None
        self.monitor_stop = threading.Event()
        self.monitor_rows: list[dict[str, float | str]] = []
        self.monitor_csv_file = None
        self.monitor_set_shg = 0.0
        self.monitor_set_thg = 0.0
        self.monitor_dev_shg: list[float] = []
        self.monitor_dev_thg: list[float] = []
        self.monitor_timestamps: list[float] = []
        self.monitor_active = False

        self.response_lines: list[str] = []
        self.current_params: dict[str, str] = {}
        self.baseline_params: dict[str, str] = {}

        self.scpi_definitions: dict[str, dict[str, str]] = {}
        self.selected_response_template = ""
        self.all_xall_field_names: list[str] = []
        self.last_sent_command = ""

        self._build_ui()
        self._refresh_ports()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="COM Port:").grid(row=0, column=0, sticky=tk.W)
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(top, textvariable=self.port_var, width=16, state="readonly")
        self.port_combo.grid(row=0, column=1, padx=5)

        ttk.Label(top, text="Baudrate:").grid(row=0, column=2, sticky=tk.W)
        self.baud_var = tk.StringVar(value="9600")
        self.baud_combo = ttk.Combobox(top, textvariable=self.baud_var, width=10, values=["9600", "19200", "38400", "57600", "115200"], state="readonly")
        self.baud_combo.grid(row=0, column=3, padx=5)

        ttk.Label(top, text="Datenbits:").grid(row=0, column=4, sticky=tk.W)
        self.bytesize_var = tk.StringVar(value="8")
        self.bytesize_combo = ttk.Combobox(top, textvariable=self.bytesize_var, width=6, values=BYTESIZE_OPTIONS, state="readonly")
        self.bytesize_combo.grid(row=0, column=5, padx=5)

        ttk.Label(top, text="Parität:").grid(row=0, column=6, sticky=tk.W)
        self.parity_var = tk.StringVar(value="None")
        self.parity_combo = ttk.Combobox(top, textvariable=self.parity_var, width=8, values=PARITY_OPTIONS, state="readonly")
        self.parity_combo.grid(row=0, column=7, padx=5)

        ttk.Label(top, text="Stopbits:").grid(row=0, column=8, sticky=tk.W)
        self.stopbits_var = tk.StringVar(value="1")
        self.stopbits_combo = ttk.Combobox(top, textvariable=self.stopbits_var, width=6, values=STOPBITS_OPTIONS, state="readonly")
        self.stopbits_combo.grid(row=0, column=9, padx=5)

        ttk.Label(top, text="Timeout (s):").grid(row=0, column=10, sticky=tk.W)
        self.timeout_var = tk.StringVar(value="1")
        ttk.Entry(top, textvariable=self.timeout_var, width=8).grid(row=0, column=11, padx=5)

        ttk.Button(top, text="Ports neu laden", command=self._refresh_ports).grid(row=0, column=12, padx=8)
        self.connect_btn = ttk.Button(top, text="Verbinden", command=self._toggle_connection)
        self.connect_btn.grid(row=0, column=13, padx=8)

        self.status_var = tk.StringVar(value="Nicht verbunden")
        ttk.Label(top, textvariable=self.status_var, foreground="#444").grid(row=0, column=14, sticky=tk.W, padx=4)

        self.monitor_btn_top = ttk.Button(
            top,
            text="monitor Temp Stage SHG/THG",
            command=self.toggle_temp_monitor,
        )
        self.monitor_btn_top.grid(row=0, column=15, padx=6)

        catalog = ttk.LabelFrame(self.root, text="SCPI Katalog", padding=10)
        catalog.pack(fill=tk.X, padx=10, pady=(0, 8))

        ttk.Button(catalog, text="SCPI Katalog laden", command=self.load_scpi_catalog).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(catalog, text="ALL/XALL Feldliste laden", command=self.load_all_xall_field_map).grid(row=0, column=1, padx=(0, 8))

        ttk.Label(catalog, text="Befehl:").grid(row=0, column=2, sticky=tk.W)
        self.catalog_cmd_var = tk.StringVar()
        self.catalog_cmd_combo = ttk.Combobox(catalog, textvariable=self.catalog_cmd_var, state="readonly", width=50)
        self.catalog_cmd_combo.grid(row=0, column=3, padx=6, sticky=tk.W)
        self.catalog_cmd_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_catalog_command_selected())
        ttk.Button(catalog, text="Befehl übernehmen", command=self.use_selected_catalog_command).grid(row=0, column=4, padx=6)

        self.template_var = tk.StringVar(value="Response-Template: -")
        self.command_info_var = tk.StringVar(value="Kategorie/Info: -")
        self.field_map_var = tk.StringVar(value="ALL/XALL Feldliste: nicht geladen")
        ttk.Label(catalog, textvariable=self.template_var).grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(6, 0))
        ttk.Label(catalog, textvariable=self.command_info_var).grid(row=2, column=0, columnspan=5, sticky=tk.W)
        ttk.Label(catalog, textvariable=self.field_map_var).grid(row=3, column=0, columnspan=5, sticky=tk.W)

        cmd_frame = ttk.LabelFrame(self.root, text="SCPI Terminal", padding=10)
        cmd_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        cmd_row = ttk.Frame(cmd_frame)
        cmd_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(cmd_row, text="Command:").pack(side=tk.LEFT)
        self.command_var = tk.StringVar()
        cmd_entry = ttk.Entry(cmd_row, textvariable=self.command_var)
        cmd_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        cmd_entry.bind("<Return>", lambda _e: self.send_command())

        ttk.Button(cmd_row, text="Senden", command=self.send_command).pack(side=tk.LEFT)
        ttk.Button(cmd_row, text="*IDN?", command=lambda: self._preset_send("*IDN?")).pack(side=tk.LEFT, padx=5)
        ttk.Button(cmd_row, text="ALL?", command=lambda: self._preset_send("ALL?")).pack(side=tk.LEFT, padx=5)
        ttk.Button(cmd_row, text="SERVice:XALL? TEMPeratures", command=lambda: self._preset_send("SERVice:XALL? TEMPeratures")).pack(side=tk.LEFT)

        self.output = tk.Text(cmd_frame, height=18, wrap=tk.NONE)
        self.output.pack(fill=tk.BOTH, expand=True)
        scroll_y = ttk.Scrollbar(self.output, orient=tk.VERTICAL, command=self.output.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.output.configure(yscrollcommand=scroll_y.set)

        bottom = ttk.Frame(self.root, padding=10)
        bottom.pack(fill=tk.X)

        self.monitor_btn_bottom = ttk.Button(
            bottom,
            text="monitor Temp Stage SHG/THG",
            command=self.toggle_temp_monitor,
        )
        self.monitor_btn_bottom.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(bottom, text="Antworten als TXT speichern", command=self.save_txt).pack(side=tk.LEFT)
        ttk.Button(bottom, text="Antworten als CSV speichern", command=self.save_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom, text="Log löschen", command=self.clear_output).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Button(bottom, text="Aktuelle Parameter aus letztem Kommando", command=self.extract_current_params).pack(side=tk.LEFT)
        ttk.Button(bottom, text="Referenzdatei laden", command=self.load_baseline_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom, text="Parameter vergleichen", command=self.compare_params).pack(side=tk.LEFT)

        monitor_frame = ttk.LabelFrame(self.root, text="Temp Monitor Δ zu Set-Wert", padding=10)
        monitor_frame.pack(fill=tk.X, padx=10, pady=(0, 8))

        monitor_header = ttk.Frame(monitor_frame)
        monitor_header.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            monitor_header,
            text="SHG/THG Live-Monitor (100 ms Polling)",
        ).pack(side=tk.LEFT)

        self.monitor_btn = ttk.Button(
            monitor_header,
            text="monitor Temp Stage SHG/THG",
            command=self.toggle_temp_monitor,
        )
        self.monitor_btn.pack(side=tk.RIGHT)

        self.monitor_canvas = tk.Canvas(monitor_frame, height=170, background="#101010", highlightthickness=0)
        self.monitor_canvas.pack(fill=tk.X, expand=True)
        self.monitor_info_var = tk.StringVar(value="Monitor aus")
        ttk.Label(monitor_frame, textvariable=self.monitor_info_var).pack(anchor=tk.W, pady=(6, 0))

        cmp_frame = ttk.LabelFrame(self.root, text="Vergleich", padding=10)
        cmp_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.compare_tree = ttk.Treeview(cmp_frame, columns=("param", "current", "baseline", "status"), show="headings", height=8)
        for col, text, width in [("param", "Parameter", 260), ("current", "Aktuell", 280), ("baseline", "Referenz", 280), ("status", "Status", 150)]:
            self.compare_tree.heading(col, text=text)
            self.compare_tree.column(col, width=width)
        self.compare_tree.pack(fill=tk.BOTH, expand=True)

        self.compare_tree.tag_configure("ok", background="#e7f8e7")
        self.compare_tree.tag_configure("diff", background="#ffe8e8")
        self.compare_tree.tag_configure("missing", background="#fff8d9")

    def _refresh_ports(self) -> None:
        if list_ports is None:
            self.port_combo["values"] = []
            self.status_var.set("pyserial nicht installiert")
            return

        ports = [p.device for p in list_ports.comports()]
        self.port_combo["values"] = ports
        self.port_var.set(ports[0] if ports else "")
        self.status_var.set("Portliste aktualisiert" if ports else "Keine COM Ports gefunden")

    def _toggle_connection(self) -> None:
        if self.ser and self.ser.is_open:
            self._disconnect()
        else:
            self._connect()

    def _connect(self) -> None:
        if serial is None:
            messagebox.showerror("Fehler", "pyserial ist nicht installiert.\nInstalliere: pip install pyserial")
            return

        port = self.port_var.get().strip()
        if not port:
            messagebox.showwarning("Hinweis", "Bitte COM Port auswählen")
            return

        try:
            baud = int(self.baud_var.get())
            timeout = float(self.timeout_var.get())
            bytesize = self._parse_bytesize(self.bytesize_var.get())
            parity = self._parse_parity(self.parity_var.get())
            stopbits = self._parse_stopbits(self.stopbits_var.get())
            self.ser = serial.Serial(
                port=port,
                baudrate=baud,
                timeout=timeout,
                bytesize=bytesize,
                parity=parity,
                stopbits=stopbits,
            )
        except Exception as exc:
            messagebox.showerror("Verbindungsfehler", str(exc))
            return

        self.stop_reader.clear()
        self.reader_thread = threading.Thread(target=self._reader_loop, daemon=True)
        self.reader_thread.start()
        self.connect_btn.configure(text="Trennen")
        self.status_var.set(f"Verbunden: {port} @ {baud}, {self._format_framing()}")
        self._append_output("[INFO] Verbindung hergestellt")

    def _format_framing(self) -> str:
        parity = self.parity_var.get().strip()
        parity_letter = "N" if parity == "None" else parity[:1].upper()
        return f"{self.bytesize_var.get().strip()}{parity_letter}{self.stopbits_var.get().strip()}"

    @staticmethod
    def _parse_bytesize(value: str) -> int:
        mapping = {
            "5": serial.FIVEBITS,
            "6": serial.SIXBITS,
            "7": serial.SEVENBITS,
            "8": serial.EIGHTBITS,
        }
        return mapping.get(value.strip(), serial.EIGHTBITS)

    @staticmethod
    def _parse_parity(value: str) -> str:
        mapping = {
            "none": serial.PARITY_NONE,
            "even": serial.PARITY_EVEN,
            "odd": serial.PARITY_ODD,
            "mark": serial.PARITY_MARK,
            "space": serial.PARITY_SPACE,
        }
        return mapping.get(value.strip().lower(), serial.PARITY_NONE)

    @staticmethod
    def _parse_stopbits(value: str) -> float:
        mapping = {
            "1": serial.STOPBITS_ONE,
            "1.5": serial.STOPBITS_ONE_POINT_FIVE,
            "2": serial.STOPBITS_TWO,
        }
        return mapping.get(value.strip(), serial.STOPBITS_ONE)

    def _query_serial(self, cmd: str) -> str:
        if not self.ser or not self.ser.is_open:
            raise RuntimeError("Nicht verbunden")

        with self.serial_lock:
            self.pause_reader.set()
            try:
                self.ser.reset_input_buffer()
                self.ser.write((cmd + "\n").encode())
                self._safe_append_output(f"TX> {cmd}")
                raw = self.ser.readline()
                text = raw.decode(errors="replace").strip() if raw else ""
                self._safe_append_output(f"RX> {text if text else '<timeout>'}")
                return text
            finally:
                self.pause_reader.clear()

    def _safe_append_output(self, line: str) -> None:
        if threading.current_thread() is threading.main_thread():
            self._append_output(line)
        else:
            self.root.after(0, self._append_output, line)

    @staticmethod
    def _parse_float_response(value: str) -> float:
        if not value:
            raise ValueError("Leere Antwort")
        match = re.search(r"[-+]?\d+(?:[\.,]\d+)?", value)
        if not match:
            raise ValueError(f"Kein Zahlenwert in Antwort: {value}")
        return float(match.group(0).replace(",", "."))

    def toggle_temp_monitor(self) -> None:
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.stop_temp_monitor()
        else:
            self.start_temp_monitor()

    def start_temp_monitor(self) -> None:
        if not self.ser or not self.ser.is_open:
            messagebox.showwarning("Nicht verbunden", "Bitte zuerst verbinden")
            return

        path = filedialog.asksaveasfilename(
            title="Monitor CSV speichern",
            defaultextension=".csv",
            initialfile=f"temp_monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return

        try:
            set_shg = self._parse_float_response(self._query_serial("SOURce:TEMPerature:LEVel:SET? TEMP_STAGE_SHG"))
            set_thg = self._parse_float_response(self._query_serial("SOURce:TEMPerature:LEVel:SET? TEMP_STAGE_THG"))
        except Exception as exc:
            messagebox.showerror("Monitor Fehler", f"Set-Werte konnten nicht gelesen werden:\n{exc}")
            return

        self.monitor_csv_file = path
        self.monitor_set_shg = set_shg
        self.monitor_set_thg = set_thg
        self.monitor_rows.clear()
        self.monitor_dev_shg.clear()
        self.monitor_dev_thg.clear()
        self.monitor_timestamps.clear()
        self.monitor_stop.clear()
        self.monitor_active = True

        self._set_monitor_button_text("Stop monitor Temp Stage SHG/THG")
        self.monitor_info_var.set(f"Monitor läuft | Set SHG={set_shg:.3f} | Set THG={set_thg:.3f}")
        self._append_output(f"[INFO] Temp Monitor gestartet -> {path}")

        self.monitor_thread = threading.Thread(target=self._temp_monitor_loop, daemon=True)
        self.monitor_thread.start()

    def stop_temp_monitor(self) -> None:
        if not self.monitor_active:
            return

        self.monitor_stop.set()
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=1.5)

        if self.monitor_rows and self.monitor_csv_file:
            self._write_monitor_csv(self.monitor_csv_file)
            self._append_output(f"[INFO] Temp Monitor CSV gespeichert: {self.monitor_csv_file}")

        self._set_monitor_button_text("monitor Temp Stage SHG/THG")
        self.monitor_info_var.set("Monitor aus")
        self.monitor_active = False

    def _set_monitor_button_text(self, text: str) -> None:
        self.monitor_btn.configure(text=text)
        self.monitor_btn_top.configure(text=text)
        self.monitor_btn_bottom.configure(text=text)

    def _temp_monitor_loop(self) -> None:
        start_ts = time.time()
        interval_s = 0.1

        while not self.monitor_stop.is_set() and self.ser and self.ser.is_open:
            ts = time.time()
            elapsed = ts - start_ts
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

            try:
                act_shg = self._parse_float_response(self._query_serial("SOURce:TEMPerature:ACTual? TEMP_STAGE_SHG"))
                act_thg = self._parse_float_response(self._query_serial("SOURce:TEMPerature:ACTual? TEMP_STAGE_THG"))
            except Exception as exc:
                self.root.after(0, self._append_output, f"[ERROR] Temp Monitor: {exc}")
                break

            dev_shg = act_shg - self.monitor_set_shg
            dev_thg = act_thg - self.monitor_set_thg

            self.monitor_rows.append(
                {
                    "timestamp": stamp,
                    "elapsed_s": elapsed,
                    "set_shg": self.monitor_set_shg,
                    "actual_shg": act_shg,
                    "deviation_shg": dev_shg,
                    "set_thg": self.monitor_set_thg,
                    "actual_thg": act_thg,
                    "deviation_thg": dev_thg,
                }
            )
            self.monitor_timestamps.append(elapsed)
            self.monitor_dev_shg.append(dev_shg)
            self.monitor_dev_thg.append(dev_thg)

            self.root.after(0, self._update_monitor_plot)

            next_tick = ts + interval_s
            remaining = next_tick - time.time()
            if remaining > 0:
                time.sleep(remaining)

        self.root.after(0, self.stop_temp_monitor)

    def _write_monitor_csv(self, path: str) -> None:
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "timestamp",
                    "elapsed_s",
                    "set_shg",
                    "actual_shg",
                    "deviation_shg",
                    "set_thg",
                    "actual_thg",
                    "deviation_thg",
                ],
                delimiter=";",
            )
            writer.writeheader()
            writer.writerows(self.monitor_rows)

    def _update_monitor_plot(self) -> None:
        canvas = self.monitor_canvas
        canvas.delete("all")

        width = max(canvas.winfo_width(), 200)
        height = max(canvas.winfo_height(), 120)
        pad = 24
        left, top, right, bottom = pad, pad, width - pad, height - pad

        canvas.create_rectangle(left, top, right, bottom, outline="#666")
        mid_y = (top + bottom) / 2
        canvas.create_line(left, mid_y, right, mid_y, fill="#555", dash=(4, 4))
        canvas.create_text(left + 4, mid_y - 8, text="Δ=0", fill="#bbb", anchor=tk.W)

        if len(self.monitor_timestamps) < 2:
            canvas.create_text(left + 4, top + 4, text="warte auf Daten...", fill="#ccc", anchor=tk.NW)
            return

        max_abs = max(
            max(abs(v) for v in self.monitor_dev_shg),
            max(abs(v) for v in self.monitor_dev_thg),
            0.01,
        )
        x_min = self.monitor_timestamps[0]
        x_max = self.monitor_timestamps[-1]
        x_range = max(x_max - x_min, 0.1)

        def to_xy(xv: float, yv: float) -> tuple[float, float]:
            x = left + ((xv - x_min) / x_range) * (right - left)
            y = mid_y - (yv / max_abs) * ((bottom - top) / 2)
            return x, y

        points_shg = [coord for x, y in (to_xy(xv, yv) for xv, yv in zip(self.monitor_timestamps, self.monitor_dev_shg)) for coord in (x, y)]
        points_thg = [coord for x, y in (to_xy(xv, yv) for xv, yv in zip(self.monitor_timestamps, self.monitor_dev_thg)) for coord in (x, y)]

        canvas.create_line(*points_shg, fill="#2ecc71", width=2, smooth=True)
        canvas.create_line(*points_thg, fill="#ff8c42", width=2, smooth=True)
        canvas.create_text(right - 120, top + 10, text="SHG Δ", fill="#2ecc71", anchor=tk.W)
        canvas.create_text(right - 120, top + 24, text="THG Δ", fill="#ff8c42", anchor=tk.W)
        canvas.create_text(
            left + 4,
            top + 4,
            text=f"±{max_abs:.3f}",
            fill="#bbb",
            anchor=tk.NW,
        )

    def _disconnect(self) -> None:
        if self.monitor_active:
            self.stop_temp_monitor()

        self.stop_reader.set()
        if self.reader_thread and self.reader_thread.is_alive():
            self.reader_thread.join(timeout=1)

        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass

        self.connect_btn.configure(text="Verbinden")
        self.status_var.set("Nicht verbunden")
        self._append_output("[INFO] Verbindung getrennt")

    def _reader_loop(self) -> None:
        while not self.stop_reader.is_set() and self.ser and self.ser.is_open:
            if self.pause_reader.is_set():
                time.sleep(0.01)
                continue
            try:
                with self.serial_lock:
                    raw = self.ser.readline()
                if raw:
                    text = raw.decode(errors="replace").strip()
                    if text:
                        self.root.after(0, self._append_output, f"RX> {text}")
                else:
                    time.sleep(0.02)
            except Exception as exc:
                self.root.after(0, self._append_output, f"[ERROR] Lesefehler: {exc}")
                break

    def _preset_send(self, cmd: str) -> None:
        self.command_var.set(cmd)
        self.send_command()

    def send_command(self) -> None:
        cmd = self.command_var.get().strip()
        if not cmd:
            return
        if not self.ser or not self.ser.is_open:
            messagebox.showwarning("Nicht verbunden", "Bitte zuerst verbinden")
            return

        try:
            self.ser.write((cmd + "\n").encode())
            self.last_sent_command = cmd
            self._append_output(f"TX> {cmd}")
            self.command_var.set("")
        except Exception as exc:
            messagebox.showerror("Senden fehlgeschlagen", str(exc))

    def _append_output(self, line: str) -> None:
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        full = f"{stamp} {line}"
        self.response_lines.append(full)
        self.output.insert(tk.END, full + "\n")
        self.output.see(tk.END)

    def clear_output(self) -> None:
        self.response_lines.clear()
        self.output.delete("1.0", tk.END)

    def save_txt(self) -> None:
        if not self.response_lines:
            messagebox.showinfo("Hinweis", "Keine Daten zum Speichern")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if path:
            Path(path).write_text("\n".join(self.response_lines), encoding="utf-8")
            messagebox.showinfo("Gespeichert", f"TXT gespeichert:\n{path}")

    def save_csv(self) -> None:
        if not self.response_lines:
            messagebox.showinfo("Hinweis", "Keine Daten zum Speichern")
            return

        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return

        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["timestamp", "direction", "message"])
            for line in self.response_lines:
                ts = line[:19]
                rest = line[20:] if len(line) > 20 else ""
                if rest.startswith("TX>"):
                    direction = "TX"
                    msg = rest[4:]
                elif rest.startswith("RX>"):
                    direction = "RX"
                    msg = rest[4:]
                else:
                    direction = "INFO"
                    msg = rest
                writer.writerow([ts, direction, msg])

        messagebox.showinfo("Gespeichert", f"CSV gespeichert:\n{path}")

    def load_scpi_catalog(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Katalog", "*.xlsx *.xlsm *.csv *.tsv *.txt"), ("Alle", "*.*")])
        if not path:
            return
        try:
            rows = self._load_catalog_rows(path)
            parsed = self._parse_scpi_rows(rows)
            self.scpi_definitions = parsed
            labels = sorted(parsed.keys())
            self.catalog_cmd_combo["values"] = labels
            if labels:
                self.catalog_cmd_var.set(labels[0])
                self._on_catalog_command_selected()
                messagebox.showinfo("SCPI geladen", f"{len(labels)} Befehle gefunden")
            else:
                messagebox.showwarning("Hinweis", "Keine SCPI-Befehle im Katalog erkannt")
        except Exception as exc:
            messagebox.showerror("Katalog Fehler", str(exc))

    def _load_catalog_rows(self, path: str) -> list[dict[str, str]]:
        ext = Path(path).suffix.lower()
        if ext in {".xlsx", ".xlsm"}:
            return self._load_catalog_rows_xlsx(path)
        return self._load_catalog_rows_delimited(path)

    def _load_catalog_rows_xlsx(self, path: str) -> list[dict[str, str]]:
        if load_workbook is None:
            raise RuntimeError("openpyxl fehlt. Bitte installieren: pip install openpyxl")

        wb = load_workbook(path, data_only=True, read_only=True)
        rows: list[dict[str, str]] = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            raw_rows = list(ws.iter_rows(min_row=1, max_row=6000, values_only=True))
            if not raw_rows:
                continue
            header_idx, header = self._detect_header(raw_rows)
            normalized = [self._norm_key(h) for h in header]
            for row in raw_rows[header_idx + 1 :]:
                values = ["" if v is None else str(v).strip() for v in row]
                item = {normalized[i]: values[i] if i < len(values) else "" for i in range(len(normalized))}
                if any(v for v in item.values()):
                    rows.append(item)
        return rows

    def _load_catalog_rows_delimited(self, path: str) -> list[dict[str, str]]:
        text = Path(path).read_text(encoding="utf-8", errors="replace")
        sample = text.splitlines()[0] if text.splitlines() else ""
        delimiter = "\t" if "\t" in sample else ";" if ";" in sample else ","
        reader = csv.DictReader(text.splitlines(), delimiter=delimiter)
        rows = []
        for row in reader:
            normalized = {self._norm_key(k): ("" if v is None else str(v).strip()) for k, v in row.items() if k}
            rows.append(normalized)
        return rows

    def _parse_scpi_rows(self, rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
        result: dict[str, dict[str, str]] = {}
        for row in rows:
            command = self._pick(row, ["scpi command", "command", "befehl", "scpi", "cmd"])
            if not command:
                continue

            category = self._pick(row, ["category", "kategorie", "group"]) or "-"
            response = self._pick(row, ["response", "antwort", "resp", "rückgabe"])
            instance = self._pick(row, ["instance"])
            parameter = self._pick(row, ["parameter"])
            unit = self._pick(row, ["unit", "einheit"])
            description = self._pick(row, ["description", "beschreibung"]) or ""

            label = f"{category} | {command}"
            result[label] = {
                "category": category,
                "command": command,
                "response": response,
                "instance": instance,
                "parameter": parameter,
                "unit": unit,
                "description": description,
            }
        return result

    def load_all_xall_field_map(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Feldliste", "*.txt *.csv"), ("Alle", "*.*")])
        if not path:
            return
        content = Path(path).read_text(encoding="utf-8", errors="replace").splitlines()
        fields = []
        for line in content:
            token = line.strip().strip(";")
            if token and token.lower() not in {"field", "parameter", "name"}:
                fields.append(token)
        self.all_xall_field_names = fields
        self.field_map_var.set(f"ALL/XALL Feldliste: {len(fields)} Einträge geladen")

    def _on_catalog_command_selected(self) -> None:
        selected = self.catalog_cmd_var.get().strip()
        entry = self.scpi_definitions.get(selected, {})
        response = entry.get("response", "")
        self.selected_response_template = response
        self.template_var.set(f"Response-Template: {response if response else '-'}")
        info = f"Kategorie: {entry.get('category', '-')} | Instance: {entry.get('instance', '-')} | Unit: {entry.get('unit', '-') }"
        self.command_info_var.set(info)

    def use_selected_catalog_command(self) -> None:
        selected = self.catalog_cmd_var.get().strip()
        if not selected or selected not in self.scpi_definitions:
            messagebox.showwarning("Hinweis", "Bitte zuerst einen Katalog-Befehl auswählen")
            return
        self.command_var.set(self.scpi_definitions[selected]["command"])

    @staticmethod
    def _norm_key(value: str) -> str:
        return str(value).strip().lower() if value else ""

    def _pick(self, row: dict[str, str], variants: list[str]) -> str:
        for key in variants:
            norm = self._norm_key(key)
            if norm in row and row[norm].strip():
                return row[norm].strip()
        return ""

    @staticmethod
    def _detect_header(rows: list[tuple]) -> tuple[int, list[str]]:
        best_idx = 0
        best_score = -1
        best_header: list[str] = []
        keys = ("category", "command", "befehl", "parameter", "response", "antwort", "unit", "description")
        for idx, row in enumerate(rows[:25]):
            header = ["" if v is None else str(v).strip().lower() for v in row]
            score = sum(1 for c in header if any(k in c for k in keys))
            if score > best_score:
                best_idx, best_score, best_header = idx, score, header
        return best_idx, best_header

    @staticmethod
    def _strip_prefix(line: str) -> str:
        trimmed = line[20:] if len(line) > 20 else line
        if trimmed.startswith("RX>"):
            return trimmed[4:].strip()
        if trimmed.startswith("TX>"):
            return ""
        return trimmed.strip()

    def _extract_last_response_block(self) -> list[str]:
        tx_indexes = [i for i, l in enumerate(self.response_lines) if " TX>" in l]
        if not tx_indexes:
            return [self._strip_prefix(l) for l in self.response_lines if self._strip_prefix(l)]

        start = tx_indexes[-1]
        result = []
        for line in self.response_lines[start + 1 :]:
            if " TX>" in line:
                break
            stripped = self._strip_prefix(line)
            if stripped:
                result.append(stripped)
        return result

    @staticmethod
    def _parse_key_values(lines: list[str]) -> dict[str, str]:
        result: dict[str, str] = {}
        for raw in lines:
            line = raw.strip()
            for segment in re.split(r"[;|]", line):
                segment = segment.strip()
                if not segment:
                    continue
                for sep in ("=", ":", ","):
                    if sep in segment:
                        left, right = segment.split(sep, 1)
                        key = left.strip()
                        value = right.strip()
                        if key:
                            result[key] = value
                        break
        return result

    @staticmethod
    def _extract_template_fields(template: str) -> list[str]:
        if not template:
            return []
        fields = re.findall(r"[A-Za-z][A-Za-z0-9_\-/]*", template)
        return [f for f in fields if len(f) > 1][:250]

    @staticmethod
    def _split_values(lines: list[str]) -> list[str]:
        values = []
        for line in lines:
            for token in re.split(r"[,;]", line):
                token = token.strip()
                if token:
                    values.append(token)
        return values

    def _xall_fields_from_command(self, command_text: str) -> list[str]:
        upper = command_text.upper().strip()
        if "XALL?" not in upper:
            return []
        for group, fields in XALL_GROUP_FIELDS.items():
            if group in upper:
                return fields
        return []

    def _parse_all_xall_response(self, lines: list[str], command_text: str) -> dict[str, str]:
        values = self._split_values(lines)
        if not values:
            return {}

        upper = command_text.upper().strip()
        fields: list[str] = []
        if upper.endswith("ALL?"):
            fields = ALL_FIELDS
        elif "XALL?" in upper:
            fields = self._xall_fields_from_command(command_text)

        if not fields and self.all_xall_field_names:
            fields = self.all_xall_field_names

        if not fields:
            return {}

        parsed: dict[str, str] = {}
        for idx, field in enumerate(fields):
            parsed[field] = values[idx] if idx < len(values) else ""
        return parsed

    def _parse_response_by_template(self, lines: list[str], template: str) -> dict[str, str]:
        command_text = self.last_sent_command.strip().upper()
        if command_text.endswith("ALL?") or "XALL?" in command_text:
            parsed = self._parse_all_xall_response(lines, command_text)
            if parsed:
                return parsed

        keys = self._extract_template_fields(template)
        if not keys:
            return self._parse_key_values(lines)

        values = self._split_values(lines)
        if len(values) < len(keys):
            return self._parse_key_values(lines)

        return {key: values[i] for i, key in enumerate(keys)}

    def extract_current_params(self) -> None:
        response_block = self._extract_last_response_block()
        if not response_block:
            messagebox.showwarning("Hinweis", "Keine Antwortdaten gefunden")
            return
        parsed = self._parse_response_by_template(response_block, self.selected_response_template)
        if not parsed:
            parsed = {"RAW_RESPONSE": " | ".join(response_block)}
        self.current_params = parsed
        messagebox.showinfo("Parameter", f"Aktuelle Parameter erkannt: {len(parsed)}")

    def load_baseline_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Vergleichsdateien", "*.txt *.csv *.xlsx"), ("Alle", "*.*")])
        if not path:
            return

        try:
            ext = Path(path).suffix.lower()
            parsed: dict[str, str]
            if ext == ".csv":
                parsed = self._parse_csv_file(path)
            elif ext in {".xlsx", ".xlsm"}:
                parsed = self._parse_excel_baseline(path)
            else:
                content = Path(path).read_text(encoding="utf-8", errors="replace").splitlines()
                parsed = self._parse_key_values(content)

            if not parsed:
                parsed = {"RAW_RESPONSE": Path(path).read_text(encoding="utf-8", errors="replace").strip()}

            self.baseline_params = parsed
            messagebox.showinfo("Referenz geladen", f"Datei: {path}\nParameter: {len(parsed)}")
        except Exception as exc:
            messagebox.showerror("Fehler", str(exc))

    @staticmethod
    def _parse_csv_file(path: str) -> dict[str, str]:
        result: dict[str, str] = {}
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            reader = csv.reader(f, delimiter=";")
            for row in reader:
                if len(row) >= 2 and row[0].strip().lower() not in {"timestamp", "direction", "message"}:
                    key = row[0].strip()
                    value = row[1].strip() if len(row) > 1 else ""
                    if key:
                        result[key] = value
        return result

    def _parse_excel_baseline(self, path: str) -> dict[str, str]:
        if load_workbook is None:
            raise RuntimeError("openpyxl fehlt für Excel Vergleichsdateien")
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        lines = []
        for row in ws.iter_rows(min_row=1, max_row=600, values_only=True):
            vals = ["" if x is None else str(x).strip() for x in row if str(x).strip()]
            if vals:
                lines.append("=".join(vals[:2]) if len(vals) >= 2 else vals[0])
        return self._parse_key_values(lines)

    def compare_params(self) -> None:
        if not self.current_params:
            messagebox.showwarning("Hinweis", "Bitte zuerst aktuelle Parameter aus letzter Antwort extrahieren")
            return
        if not self.baseline_params:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine Referenzdatei laden")
            return

        for item in self.compare_tree.get_children():
            self.compare_tree.delete(item)

        keys = sorted(set(self.current_params) | set(self.baseline_params))
        diff_count = 0
        for key in keys:
            current = self.current_params.get(key)
            baseline = self.baseline_params.get(key)

            if current is None:
                status, tag = "Fehlt aktuell", "missing"
                diff_count += 1
            elif baseline is None:
                status, tag = "Fehlt Referenz", "missing"
                diff_count += 1
            elif current == baseline:
                status, tag = "OK", "ok"
            else:
                status, tag = "Abweichung", "diff"
                diff_count += 1

            self.compare_tree.insert("", tk.END, values=(key, current or "", baseline or "", status), tags=(tag,))

        self.status_var.set(f"Vergleich abgeschlossen: {diff_count} Abweichung(en)")


def main() -> None:
    root = tk.Tk()
    app = SCPITerminalApp(root)

    def on_close() -> None:
        app._disconnect()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
