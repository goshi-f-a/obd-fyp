"""
=======================================================================
    OBD-II Live Data Logger — Logs VEH_NO, VEH_TYPE, YR_MFR & Timestamp
=======================================================================
    File Name: OBD-v3.py

📄 Description:
    This Python script connects to a Bluetooth-based ELM327 OBD-II adapter,
    collects real-time vehicle telemetry data, and logs it to a single CSV file
    for all vehicles. The CSV log includes vehicle metadata such as vehicle
    number, type, and year of manufacture along with a timezone-aware timestamp.

🔧 Features:
    - Automatically scans and connects to Bluetooth serial ports.
    - Queries ECU for key engine and environmental parameters.
    - Appends readings row-wise to a common CSV file in UTF-8 (Excel-safe).
    - Prints live tabular data to terminal.
    - Supports both metric and imperial units.

📁 Output:
    - A unified dataset: `obd_dataset.csv`
    - UTF-8 with BOM encoding for Excel compatibility

🛠 Dependencies:
    - python-OBD
    - pyserial
    - pytz

🧪 Logged OBD-II Parameters:
    - Coolant Temperature
    - Engine RPM
    - Throttle Position
    - Engine Load
    - Vehicle Speed
    - Fuel Level
    - Intake Air Temperature
    - Ambient Air Temperature

📅 Author: Farooque Azam
🗓️ Last Modified: 2025-07-06
"""

import time
from datetime import datetime
import csv
import re
import os
import serial
import serial.tools.list_ports
import pytz
from obd import OBD, commands

# ------------------------- Configuration -------------------------

PROTOCOL = "6"           # OBD-II protocol (ISO 15765-4 CAN)
BAUDRATE = 38400         # Bluetooth ELM327 default baudrate
REFRESH_RATE = 1.0       # Time (in seconds) between samples
UNITS = "metric"         # "metric" for °C/km/h, "imperial" for °F/mph

VEH_NO    = "BBJ-91"                 # Vehicle registration number
VEH_TYPE  = "Chery Tiggo 8 Pro"      # Model/make
YR_MFR    = "2023"                   # Year of manufacture
TIMEZONE  = "Asia/Karachi"           # Local timezone for timestamp

CSV_FILENAME = "obd_dataset.csv"     # Common dataset file

# ------------------------- Data Setup ----------------------------

headers = [
    "Timestamp", "Vehicle No", "Vehicle Type", "Year",
    "Coolant Temp", "Engine RPM", "Throttle Pos", "Engine Load",
    "Speed", "Fuel Level", "Intake Temp", "Ambient Temp"
]

# ------------------------- Utility Functions ---------------------

def clear_console():
    """Clears the terminal using ANSI escape sequences."""
    print("\033[H\033[J", end="")

def extract_mac_from_hwid(hwid: str) -> str:
    """
    Extracts MAC address from a device HWID string.
    Returns formatted MAC address or default if not found.
    """
    match = re.search(r'&([0-9A-F]{12})', hwid.upper())
    if match:
        mac_raw = match.group(1)
        mac = ":".join(mac_raw[i:i+2] for i in range(0, 12, 2))
        return mac
    return "00:00:00:00:00:00"

def list_paired_bluetooth_ports():
    """
    Lists COM ports with Bluetooth devices having valid MAC addresses.
    Returns a list of tuples: (port, MAC)
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
    Parses and formats the OBD-II response to include appropriate units.
    Automatically converts to imperial if selected.
    """
    if response.is_null():
        return None
    value = response.value.magnitude

    if unit == "°C" and UNITS == "imperial":
        value = value * 9/5 + 32
        unit = "°F"
    elif unit == "km/h" and UNITS == "imperial":
        value = value * 0.621371
        unit = "mph"

    return f"{value:.1f} {unit}" if isinstance(value, float) else f"{value} {unit}"

def display_table(data, port):
    """
    Displays a live table of sensor data in the terminal.
    Shows current COM port and connection protocol.
    """
    clear_console()
    print(f"=== OBD-II MONITOR [{UNITS.upper()}] ===")
    print(f"Connected via {port} | Protocol: {connection.protocol_name()}")
    print(f"Refresh Rate: {REFRESH_RATE}s | Press Ctrl+C to stop\n")
    print("|".join(f"{h:^15}" for h in headers))
    print("-" * (15 * len(headers) + len(headers) - 1))
    for row in data:
        print("|".join(f"{str(x):^15}" for x in row))

def initialize_csv(file_path):
    """
    Initializes the CSV file if it doesn't exist.
    Writes header row only once.
    """
    if not os.path.isfile(file_path):
        with open(file_path, 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerow(headers)

def append_row_to_csv(file_path, row):
    """
    Appends a single row of sensor data to the CSV file using UTF-8 BOM encoding.
    """
    with open(file_path, 'a', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        writer.writerow(row)

# ------------------------- Connection Logic ----------------------

def connect_with_retry(protocol, baudrate=38400, retries=5, delay=2, timeout=3):
    """
    Tries to connect to all paired Bluetooth COM ports.
    Returns a working OBD connection and its port, or None.
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
            time.sleep(10)
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
    """
    Entry point for live data monitoring and CSV logging.
    """
    global connection
    connection, detected_port = connect_with_retry(PROTOCOL, BAUDRATE)

    if not connection or not connection.is_connected():
        print("❌ Could not establish a stable connection to the vehicle.")
        return

    print("✅ OBD-II adapter connected.")
    print("⏳ Waiting 2 seconds to stabilize...")
    time.sleep(2)

    initialize_csv(CSV_FILENAME)

    try:
        while True:
            now = datetime.now(pytz.timezone(TIMEZONE))
            timestamp = now.isoformat()

            row = [
                timestamp,
                VEH_NO,
                VEH_TYPE,
                YR_MFR,
                get_formatted_value(connection.query(commands.COOLANT_TEMP), '°C'),
                get_formatted_value(connection.query(commands.RPM), ''),
                get_formatted_value(connection.query(commands.THROTTLE_POS), '%'),
                get_formatted_value(connection.query(commands.ENGINE_LOAD), '%'),
                get_formatted_value(connection.query(commands.SPEED), 'km/h'),
                get_formatted_value(connection.query(commands.FUEL_LEVEL), '%'),
                get_formatted_value(connection.query(commands.INTAKE_TEMP), '°C'),
                get_formatted_value(connection.query(commands.AMBIANT_AIR_TEMP), '°C')
            ]

            append_row_to_csv(CSV_FILENAME, row)
            display_table([row], detected_port)
            time.sleep(REFRESH_RATE)

    except KeyboardInterrupt:
        print("\n⏹️ Monitoring stopped by user.")

    finally:
        connection.close()
        print("🔌 OBD connection closed.")

# ------------------------- Entry Point ---------------------------

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n🛑 Program interrupted by user. Exiting gracefully...")
