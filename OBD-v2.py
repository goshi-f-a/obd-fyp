"""
=======================================================================
    OBD-II Live Data Logger for Paired Bluetooth ELM327 Adapters
=======================================================================
    File Name: OBD-v2.py
    
📄 Description:
    This script automatically detects and connects to a **paired Bluetooth OBD-II adapter** 
    (like ELM327) via a serial COM port, retrieves real-time vehicle diagnostics data from 
    the ECU, and displays it in a live terminal table. The data is also optionally saved to 
    a CSV file upon exit.

🔧 Key Features:
    - Scans only **Bluetooth COM ports** with valid MAC addresses.
    - Detects and connects to the first **responsive OBD-II adapter**.
    - Displays live data (Coolant Temp, RPM, Speed, Throttle, etc.) in tabular format.
    - Optionally saves log to CSV upon termination.
    - Supports metric and imperial unit systems.

🛠️ Requirements:
    - Python 3.7+
    - `python-OBD` (for OBD communication)
    - `pyserial` (for COM port detection and access)

📦 Install dependencies:
    pip install obd pyserial

🚗 Compatible with:
    - ELM327 Bluetooth adapters (Standard OBD-II Protocols)
    - Tested on Windows (adjustable for Linux/macOS if needed)

🧪 Sample OBD-II Parameters Queried:
    - Coolant Temperature
    - Engine RPM
    - Throttle Position
    - Engine Load
    - Vehicle Speed
    - Fuel Level
    - Intake Air Temp
    - Ambient Air Temp

📁 Output:
    - Real-time terminal table
    - CSV file with timestamped readings (optional on exit)

📅 Author: Farooque Azam
🗓️ Last Modified: 2025-07-05

=======================================================================
"""

import time
from datetime import datetime
import csv
import re
import serial
import serial.tools.list_ports
from obd import OBD, commands

# ------------------------- Configuration -------------------------

PROTOCOL = "6"           # OBD-II protocol (e.g., "6" = ISO 15765-4 CAN)
BAUDRATE = 38400         # Common for Bluetooth ELM327
REFRESH_RATE = 1.0       # Time between readings (in seconds)
UNITS = "metric"         # Use "imperial" for °F and mph

# ------------------------- Data Setup ----------------------------

data_log = []  # Collected sensor readings

# Headers for CSV and display
headers = [
    "Timestamp", "Coolant Temp", "Engine RPM",
    "Throttle Pos", "Engine Load", "Speed",
    "Fuel Level", "Intake Temp", "Ambient Temp"
]

# ------------------------- Utility Functions ---------------------

def clear_console():
    """Clears the terminal output using ANSI escape codes."""
    print("\033[H\033[J", end="")

def extract_mac_from_hwid(hwid: str) -> str:
    """
    Extracts MAC address from HWID string.

    Parameters:
        hwid (str): Hardware ID string from serial.tools.list_ports

    Returns:
        str: MAC address formatted as XX:XX:XX:XX:XX:XX or default 00:... if not found
    """
    match = re.search(r'&([0-9A-F]{12})', hwid.upper())
    if match:
        mac_raw = match.group(1)
        mac = ":".join(mac_raw[i:i+2] for i in range(0, 12, 2))
        return mac
    return "00:00:00:00:00:00"

def list_paired_bluetooth_ports():
    """
    Lists only paired Bluetooth serial ports with valid MACs.

    Returns:
        list of tuples: [(COMx, MAC), ...]
    """
    ports = serial.tools.list_ports.comports()
    filtered = []
    for port in ports:
        if "Bluetooth" in port.description:
            mac = extract_mac_from_hwid(port.hwid)
            if mac != "00:00:00:00:00:00":
                filtered.append((port.device, mac))
    return filtered

def get_formatted_value(response, unit):
    """
    Parses and formats sensor values from OBD response.

    Parameters:
        response (obd.OBDResponse): OBD sensor response
        unit (str): Expected unit (e.g., °C, %, km/h)

    Returns:
        str or None: Formatted value with unit
    """
    if response.is_null():
        return None
    value = response.value.magnitude

    # Unit conversion if needed
    if unit == "°C" and UNITS == "imperial":
        value = value * 9/5 + 32
        unit = "°F"
    elif unit == "km/h" and UNITS == "imperial":
        value = value * 0.621371
        unit = "mph"

    return f"{value:.1f} {unit}" if isinstance(value, float) else f"{value} {unit}"

def display_table(data, port):
    """
    Clears and displays live OBD data in tabular format.

    Parameters:
        data (list): List of data rows
        port (str): Connected COM port for display
    """
    clear_console()
    print(f"=== OBD-II MONITOR [{UNITS.upper()}] ===")
    print(f"Connected via {port} | Protocol: {connection.protocol_name()}")
    print(f"Refresh Rate: {REFRESH_RATE}s | Press Ctrl+C to stop\n")
    print("|".join(f"{h:^15}" for h in headers))
    print("-" * (15 * len(headers) + len(headers) - 1))
    for row in data:
        print("|".join(f"{str(x):^15}" for x in row))

def save_to_csv(data):
    """
    Saves the data log to a timestamped CSV file.

    Parameters:
        data (list): Collected OBD data rows
    """
    filename = f"obd_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    try:
        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            writer.writerows(data)
        print(f"\n✅ Data saved to {filename}")
    except Exception as e:
        print(f"\n❌ Error saving CSV: {e}")

# ------------------------- Connection Logic ----------------------

def connect_with_retry(protocol, baudrate=38400, retries=5, delay=2, timeout=3):
    """
    Scans paired Bluetooth COM ports and attempts to connect to OBD-II.

    Returns:
        tuple: (OBD connection object, port name) if successful, else (None, None)
    """
    print("🔄 Waiting before first attempt...")
    time.sleep(5)

    print("🔍 Scanning for paired Bluetooth OBD-II devices...")
    available_ports = list_paired_bluetooth_ports()

    if not available_ports:
        print("🚫 No paired Bluetooth OBD-II ports found.")
        return None, None

    print("🧭 Found paired ports:")
    for port, mac in available_ports:
        print(f"  - {port} (MAC: {mac})")

    for port, mac in available_ports:
        print(f"\n🔍 Pre-checking {port} ({mac})...")
        try:
            time.sleep(2)
            s = serial.Serial(port=port, baudrate=baudrate, timeout=timeout)
            time.sleep(2)
            s.close()
            time.sleep(10)  # Allow the device to settle
        except Exception as e:
            print(f"⚠️ Skipping {port}: {e}")
            continue

        for attempt in range(1, retries + 1):
            print(f"🧪 Trying OBD connection on {port} (Attempt {attempt}/{retries})...")
            try:
                connection = OBD(
                    portstr=port,
                    baudrate=baudrate,
                    protocol=protocol,
                    fast=False,
                    timeout=timeout
                )
                if connection.is_connected():
                    print(f"✅ Connected successfully on {port}")
                    return connection, port
                else:
                    print(f"❌ {port} did not respond as OBD.")
                    connection.close()
            except Exception as e:
                print(f"⛔ Error on {port}: {e}")
            time.sleep(delay)

    print("🚫 No responsive OBD-II adapter found.")
    return None, None

# ------------------------- Main Monitoring -----------------------

def main():
    global connection
    connection, detected_port = connect_with_retry(PROTOCOL, BAUDRATE)

    if not connection or not connection.is_connected():
        print("❌ Could not establish a stable connection to the vehicle.")
        return

    print("✅ OBD-II adapter connected.")
    print("⏳ Waiting 2 seconds to stabilize...")
    time.sleep(2)

    try:
        while True:
            timestamp = datetime.now().strftime('%H:%M:%S')
            row = [
                timestamp,
                get_formatted_value(connection.query(commands.COOLANT_TEMP), '°C'),
                get_formatted_value(connection.query(commands.RPM), ''),
                get_formatted_value(connection.query(commands.THROTTLE_POS), '%'),
                get_formatted_value(connection.query(commands.ENGINE_LOAD), '%'),
                get_formatted_value(connection.query(commands.SPEED), 'km/h'),
                get_formatted_value(connection.query(commands.FUEL_LEVEL), '%'),
                get_formatted_value(connection.query(commands.INTAKE_TEMP), '°C'),
                get_formatted_value(connection.query(commands.AMBIANT_AIR_TEMP), '°C')
            ]
            data_log.append(row)
            display_table(data_log, detected_port)
            time.sleep(REFRESH_RATE)

    except KeyboardInterrupt:
        print("\n⏹️ Monitoring stopped by user.")
        if data_log:
            save = input("💾 Save data to CSV? (y/n): ").strip().lower()
            if save == 'y':
                save_to_csv(data_log)

    finally:
        connection.close()
        print("🔌 OBD connection closed.")

# ------------------------- Entry Point ---------------------------

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n🛑 Program interrupted by user. Exiting gracefully...")
