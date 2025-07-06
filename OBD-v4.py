import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports
import threading
import time
from datetime import datetime
import csv
import os
import pytz
import re
import json
from obd import OBD, commands

# -------------------- Config Storage --------------------
CONFIG_FILE = "veh_config.json"

def load_config():
    """Loads vehicle configuration from a JSON file."""
    if os.path.isfile(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {"VEH_NO": "BBJ-91", "VEH_TYPE": "Chery Tiggo 8 Pro", "YR_MFR": "2023"}

def save_config(config):
    """Saves vehicle configuration to a JSON file."""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

# -------------------- Global Settings --------------------
config = load_config()
TIMEZONE = "Asia/Karachi"
CSV_FILENAME = "obd_dataset.csv"
UNITS = "metric"
# MODIFIED: Changed refresh rate to 2 seconds
REFRESH_RATE = 2.0
BAUDRATE = 38400

# -------------------- Headers --------------------
headers = [
    "Timestamp", "Coolant Temp", "Engine RPM", "Throttle Pos", "Engine Load",
    "Speed", "Fuel Level", "Intake Temp", "Ambient Temp"
]

# -------------------- Utility Functions --------------------
def extract_mac_from_hwid(hwid: str) -> str:
    """Extracts MAC address from a device HWID string."""
    match = re.search(r'&([0-9A-F]{12})', hwid.upper())
    if match:
        mac_raw = match.group(1)
        return ":".join(mac_raw[i:i+2] for i in range(0, 12, 2))
    return "00:00:00:00:00:00"

def list_paired_bluetooth_ports():
    """
    Filters for COM ports with a valid MAC address
    and returns a list of tuples (port_name, mac_address).
    """
    all_ports = serial.tools.list_ports.comports()
    paired_ports = []
    for port in all_ports:
        if "Bluetooth" in port.description:
            mac = extract_mac_from_hwid(port.hwid)
            if mac != "00:00:00:00:00:00":
                paired_ports.append((port.device, mac))
    return paired_ports

def get_formatted_value(response):
    """Formats OBD response values, handling nulls and units."""
    if response is None or response.is_null() or response.value is None:
        return "N/A"
    
    value = response.value.magnitude
    unit_label = response.unit
    
    if isinstance(value, float):
        return f"{value:.1f} {unit_label}"
    return f"{value} {unit_label}"

def initialize_csv(file_path):
    """Creates the CSV file with headers if it doesn't exist."""
    if not os.path.isfile(file_path):
        with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["Timestamp", "Vehicle No", "Vehicle Type", "Year"] + headers[1:])

def append_row_to_csv(file_path, row):
    """Appends a single data row to the CSV file."""
    with open(file_path, 'a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(row)

# -------------------- GUI Class --------------------
class OBDLoggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OBD-II Logger GUI v11.0")

        self.connection = None
        self.running = False
        
        self.port_map = {}

        # StringVars for entry boxes
        self.veh_no = tk.StringVar(value=config.get("VEH_NO"))
        self.veh_type = tk.StringVar(value=config.get("VEH_TYPE"))
        self.yr_mfr = tk.StringVar(value=config.get("YR_MFR"))

        self.column_vars = {h: tk.BooleanVar(value=True) for h in headers}
        self.last_sort = {'col': None, 'rev': False}

        self.display_veh_no = tk.StringVar(value="Vehicle No: --")
        self.display_veh_type = tk.StringVar(value="Vehicle Type: --")
        self.display_yr_mfr = tk.StringVar(value="Year: --")

        self.create_widgets()
        self.update_ports_list()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """Creates and places all the widgets in the application window."""
        # --- Vehicle Info Frame (Entry boxes) ---
        frm_info = ttk.Frame(self.root)
        frm_info.pack(padx=10, pady=(10,5), fill='x')

        ttk.Label(frm_info, text="Vehicle No:").grid(row=0, column=0, padx=2, pady=2, sticky='w')
        ttk.Entry(frm_info, textvariable=self.veh_no, width=15).grid(row=0, column=1)
        ttk.Label(frm_info, text="Vehicle Type:").grid(row=0, column=2, padx=(10, 2), pady=2, sticky='w')
        ttk.Entry(frm_info, textvariable=self.veh_type, width=25).grid(row=0, column=3)
        ttk.Label(frm_info, text="Year:").grid(row=0, column=4, padx=(10, 2), pady=2, sticky='w')
        ttk.Entry(frm_info, textvariable=self.yr_mfr, width=8).grid(row=0, column=5)
        
        frm_cols = ttk.LabelFrame(self.root, text="Show/Hide Columns")
        frm_cols.pack(padx=10, pady=5, fill='x')
        for i, col in enumerate(headers):
            cb = ttk.Checkbutton(frm_cols, text=col, variable=self.column_vars[col], command=self.update_visible_columns)
            cb.grid(row=0, column=i, padx=5, sticky='w')

        # --- Controls Frame ---
        frm_controls = ttk.Frame(self.root)
        frm_controls.pack(padx=10, pady=5, fill='x')
        
        ttk.Label(frm_controls, text="Paired Port:").pack(side='left', padx=(0,5))
        self.port_combo = ttk.Combobox(frm_controls, state="readonly", width=25)
        self.port_combo.pack(side='left', padx=(0,5))
        self.refresh_btn = ttk.Button(frm_controls, text="Refresh", command=self.update_ports_list, width=8)
        self.refresh_btn.pack(side='left', padx=(0, 10))
        self.connect_btn = ttk.Button(frm_controls, text="Connect", command=self.toggle_connection, width=12)
        self.connect_btn.pack(side='left', padx=(0, 5))
        self.monitor_btn = ttk.Button(frm_controls, text="Start Monitoring", command=self.toggle_monitoring, state='disabled', width=18)
        self.monitor_btn.pack(side='left')
        self.status_label = ttk.Label(frm_controls, text="Status: Disconnected", foreground="red", font=('Helvetica', 10, 'bold'))
        self.status_label.pack(side='left', padx=(10, 0))

        # --- Frame for displaying vehicle info labels ---
        frm_display_info = ttk.Frame(self.root)
        frm_display_info.pack(padx=10, pady=(5, 2), fill='x')
        style = ttk.Style()
        style.configure("Display.TLabel", font=('Helvetica', 9, 'bold'))
        ttk.Label(frm_display_info, textvariable=self.display_veh_no, style="Display.TLabel").pack(side='left')
        ttk.Label(frm_display_info, textvariable=self.display_veh_type, style="Display.TLabel").pack(side='left', padx=30)
        ttk.Label(frm_display_info, textvariable=self.display_yr_mfr, style="Display.TLabel").pack(side='left')
        
        # --- Treeview (Data Table) ---
        self.tree = ttk.Treeview(self.root, columns=headers, show='headings', height=12)
        for h in headers:
            self.tree.heading(h, text=h, command=lambda _h=h: self.sort_column(_h, False))
            self.tree.column(h, width=110, anchor='center')
        self.tree.pack(fill='both', expand=True, padx=10, pady=(0,5))
        self.update_visible_columns()

        # --- Log Output with scrollbar ---
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        log_scrollbar = ttk.Scrollbar(log_frame)
        log_scrollbar.pack(side='right', fill='y')
        self.log_output = tk.Text(log_frame, height=4, bg="black", fg="lime green", 
                                  wrap='word', font=("Courier New", 9), 
                                  yscrollcommand=log_scrollbar.set)
        self.log_output.pack(fill='both', expand=True)
        log_scrollbar.config(command=self.log_output.yview)

    def update_visible_columns(self):
        """Updates the treeview to show only the columns selected via checkbox."""
        visible_cols = [h for h, var in self.column_vars.items() if var.get()]
        self.tree["displaycolumns"] = visible_cols

    def sort_column(self, col, reverse):
        """Sorts a treeview column when its header is clicked."""
        try:
            def get_sort_key(value_str):
                if isinstance(value_str, str):
                    match = re.match(r"(-?\d+\.?\d*)", value_str)
                    if match:
                        return float(match.group(1))
                return value_str
            data = [(get_sort_key(self.tree.set(k, col)), k) for k in self.tree.get_children('')]
        except tk.TclError:
            self.log(f"Cannot sort by '{col}' as it is currently hidden.")
            return
        data.sort(reverse=reverse)
        for index, (val, k) in enumerate(data):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    def update_ports_list(self):
        """Scans for ports, formats them with MAC, and updates the combobox."""
        self.log("Scanning for paired Bluetooth devices by MAC address...")
        ports_with_mac = list_paired_bluetooth_ports()
        self.port_map = {f"{port} (MAC: {mac})": port for port, mac in ports_with_mac}
        display_values = list(self.port_map.keys())
        self.port_combo['values'] = display_values
        if display_values:
            self.port_combo.current(0)
            self.log(f"Found paired ports: {', '.join(display_values)}")
        else:
            self.log("No paired Bluetooth devices with valid MAC addresses found.")
            self.port_combo.set('')

    def log(self, msg):
        """Thread-safe method to log messages to the GUI's text widget."""
        self.root.after(0, self._log_message, msg)

    def _log_message(self, msg):
        self.log_output.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')}: {msg}\n")
        self.log_output.see(tk.END)

    def toggle_connection(self):
        """Toggles the connection state."""
        if self.connection and self.connection.is_connected():
            self.disconnect()
        else:
            self.connect()
    
    def connect(self):
        """Starts the connection process using the selected COM port."""
        display_string = self.port_combo.get()
        if not display_string:
            self.log("⚠️ Please select a paired port first.")
            return
        port = self.port_map.get(display_string)
        if not port:
            self.log(f"Error: Could not find port for '{display_string}'.")
            return
        self.display_veh_no.set(f"Vehicle No: {self.veh_no.get()}")
        self.display_veh_type.set(f"Vehicle Type: {self.veh_type.get()}")
        self.display_yr_mfr.set(f"Year: {self.yr_mfr.get()}")
        self.log(f"🚀 Starting connection process for {port}...")
        self.connect_btn.config(state='disabled')
        self.monitor_btn.config(state='disabled')
        self.port_combo.config(state='disabled')
        self.refresh_btn.config(state='disabled')
        self.status_label.config(text=f"Status: Connecting to {port}...", foreground="orange")
        threading.Thread(target=self.attempt_connection_on_port, args=(port,), daemon=True).start()

    def attempt_connection_on_port(self, port):
        """Attempts to connect to a specific port, with retries. This runs in a thread."""
        for attempt in range(1, 4):
            self.log(f"🧪 Trying connection on {port} (Attempt {attempt}/3)...")
            try:
                conn = OBD(portstr=port, baudrate=BAUDRATE, fast=False, timeout=5)
                if conn.is_connected():
                    self.log(f"✅ Connection successful on {port}!")
                    # ADDED: 5-second stabilization delay
                    self.log("⏳ Stabilizing connection, please wait 5 seconds...")
                    time.sleep(5)
                    self.connection = conn
                    self.root.after(0, self.update_ui_on_connect, port)
                    return
                else:
                    self.log(f"❌ {port} did not respond as an OBD device.")
                    conn.close()
            except Exception as e:
                self.log(f"⛔ Error on {port}: {str(e).strip()}")
            # ADDED: 2-second delay between attempts
            time.sleep(2)
            
        self.log(f"🚫 All connection attempts failed for {port}.")
        self.root.after(0, self.update_ui_on_fail, "Connection Failed")

    def update_ui_on_connect(self, port):
        """Updates the GUI to reflect a successful connection."""
        self.status_label.config(text=f"Status: Connected on {port}", foreground="green")
        self.connect_btn.config(text="Disconnect", state='normal')
        self.monitor_btn.config(state='normal')

    def update_ui_on_fail(self, reason):
        """Updates the GUI to reflect a failed connection attempt."""
        self.status_label.config(text=f"Status: {reason}", foreground="red")
        self.connect_btn.config(text="Connect", state='normal')
        self.port_combo.config(state='readonly')
        self.refresh_btn.config(state='normal')
        self.connection = None

    def disconnect(self):
        """Disconnects from the OBD adapter and resets the UI."""
        self.log("🔌 Disconnecting...")
        if self.running:
            self.toggle_monitoring()
        if self.connection:
            self.connection.close()
        self.connection = None
        self.display_veh_no.set("Vehicle No: --")
        self.display_veh_type.set("Vehicle Type: --")
        self.display_yr_mfr.set("Year: --")
        self.status_label.config(text="Status: Disconnected", foreground="red")
        self.connect_btn.config(text="Connect", state='normal')
        self.monitor_btn.config(state='disabled')
        self.port_combo.config(state='readonly')
        self.refresh_btn.config(state='normal')
        self.log("Disconnected successfully.")

    def toggle_monitoring(self):
        """Starts or stops the data monitoring loop."""
        if self.running:
            self.running = False
            self.monitor_btn.config(text="Start Monitoring")
            self.log("⏹️ Monitoring stopped.")
        else:
            self.running = True
            self.monitor_btn.config(text="Stop Monitoring")
            self.log("▶️ Monitoring started...")
            initialize_csv(CSV_FILENAME)
            threading.Thread(target=self.monitor_data, daemon=True).start()

    def monitor_data(self):
        """The main data-gathering loop. Runs in a thread."""
        while self.running:
            try:
                if not self.connection or not self.connection.is_connected():
                    self.log("⚠️ Connection lost! Stopping monitoring.")
                    self.root.after(0, self.disconnect)
                    break

                now = datetime.now(pytz.timezone(TIMEZONE))
                timestamp = now.isoformat()
                
                data_points = {
                    "Coolant Temp": get_formatted_value(self.connection.query(commands.COOLANT_TEMP)),
                    "Engine RPM": get_formatted_value(self.connection.query(commands.RPM)),
                    "Throttle Pos": get_formatted_value(self.connection.query(commands.THROTTLE_POS)),
                    "Engine Load": get_formatted_value(self.connection.query(commands.ENGINE_LOAD)),
                    "Speed": get_formatted_value(self.connection.query(commands.SPEED)),
                    "Fuel Level": get_formatted_value(self.connection.query(commands.FUEL_LEVEL)),
                    "Intake Temp": get_formatted_value(self.connection.query(commands.INTAKE_TEMP)),
                    "Ambient Temp": get_formatted_value(self.connection.query(commands.AMBIANT_AIR_TEMP)),
                }
                
                row_data = [data_points.get(h, "N/A") for h in headers[1:]]
                
                self.root.after(0, self.update_gui_and_csv, timestamp, row_data)
                
                # The refresh rate (delay) is controlled by the REFRESH_RATE global variable
                time.sleep(REFRESH_RATE)
            
            except Exception as e:
                self.log(f"⛔ ERROR in monitoring loop: {e}")
                self.log("⏹️ Halting monitoring due to error.")
                self.root.after(0, self.toggle_monitoring)
                break
        
        self.root.after(0, lambda: self.monitor_btn.config(text="Start Monitoring"))
        self.running = False
        
    def update_gui_and_csv(self, timestamp, row_data):
        """Safely updates the GUI table and writes to the CSV from the main thread."""
        display_ts = timestamp.split('T')[1].split('+')[0].split('.')[0]
        if self.tree.winfo_exists():
            self.tree.insert('', 'end', values=[display_ts] + row_data)
            self.tree.yview_moveto(1)
        full_row = [timestamp, self.veh_no.get(), self.veh_type.get(), self.yr_mfr.get()] + row_data
        append_row_to_csv(CSV_FILENAME, full_row)

    def on_closing(self):
        """Handles window close event, saving config and disconnecting."""
        self.log("Exiting application...")
        current_config = {
            "VEH_NO": self.veh_no.get(),
            "VEH_TYPE": self.veh_type.get(),
            "YR_MFR": self.yr_mfr.get()
        }
        save_config(current_config)
        if self.connection and self.connection.is_connected():
            threading.Thread(target=self.connection.close, daemon=True).start()
        self.root.destroy()

# -------------------- Main Entry Point --------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = OBDLoggerApp(root)
    root.mainloop()