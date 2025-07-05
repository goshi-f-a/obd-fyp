import time                      # For delays and time tracking
from obd import OBD, commands    # OBD-II library for communication with vehicle
from datetime import datetime    # For timestamps in logs
import csv                       # For saving collected data to CSV

# ------------------------- Configuration Section -------------------------

PORT = "COM4"             # The COM port connected to the OBD-II adapter
PROTOCOL = "6"            # OBD-II protocol number (e.g., "6" = ISO 15765-4 CAN)
BAUDRATE = 38400          # Communication speed (matches adapter default)
REFRESH_RATE = 1.0        # Interval between OBD queries (in seconds)
UNITS = "metric"          # Measurement system: "metric" (°C, km/h) or "imperial" (°F, mph)

# ------------------------- Data and Display Setup -------------------------

data_log = []  # List to store logged data rows

# CSV and display table headers
headers = [
    "Timestamp", "Coolant Temp", "Engine RPM",
    "Throttle Pos", "Engine Load", "Speed",
    "Fuel Level", "Intake Temp", "Ambient Temp"
]

def clear_console():
    """
    Clears the terminal screen using ANSI escape codes.
    Ensures the display table appears clean and updated each cycle.
    """
    print("\033[H\033[J", end="")

def get_formatted_value(response, unit):
    """
    Extracts and formats sensor value from the OBD response.
    
    Parameters:
        response (obd.OBDResponse): The result of an OBD query.
        unit (str): The expected unit for the sensor value.
    
    Returns:
        str or None: Formatted sensor value with unit, or None if no data.
    """
    if response.is_null():
        return None  # No data received or sensor unsupported

    value = response.value.magnitude  # Extract raw numerical value

    # Convert units to imperial if requested
    if unit == "°C" and UNITS == "imperial":
        value = value * 9/5 + 32
        unit = "°F"
    elif unit == "km/h" and UNITS == "imperial":
        value = value * 0.621371
        unit = "mph"

    return f"{value:.1f} {unit}" if isinstance(value, float) else f"{value} {unit}"

def display_table(data):
    """
    Displays logged OBD data in a tabular format on the terminal.
    
    Parameters:
        data (list): List of data rows to display.
    """
    clear_console()
    print(f"=== OBD-II MONITOR [{UNITS.upper()}] ===")
    print(f"Connected via {PORT} | Protocol: {connection.protocol_name()}")
    print(f"Reading interval: {REFRESH_RATE} sec | Press Ctrl+C to exit\n")
    
    # Print table header
    print("|".join(f"{h:^15}" for h in headers))
    print("-" * (15 * len(headers) + len(headers) - 1))

    # Print each row of logged values
    for row in data:
        print("|".join(f"{str(x):^15}" for x in row))

def save_to_csv(data):
    """
    Saves collected data to a timestamped CSV file.
    
    Parameters:
        data (list): List of data rows to save.
    """
    filename = f"obd_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    try:
        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            writer.writerows(data)
        print(f"\n✅ Data saved successfully to {filename}")
    except Exception as e:
        print(f"\n❌ Error saving file: {str(e)}")

def connect_with_retry(port, baudrate, protocol, retries=5, delay=2):
    """
    Attempts to establish a connection to the OBD-II adapter with retries.

    Parameters:
        port (str): COM port identifier.
        baudrate (int): Serial communication baud rate.
        protocol (str): OBD-II protocol string.
        retries (int): Number of retry attempts.
        delay (int): Delay in seconds between retries.

    Returns:
        obd.OBD or None: Connected OBD instance or None if failed.
    """
    print("🔄 Waiting before first attempt...")
    time.sleep(5)  # Allow OS and hardware to stabilize

    for attempt in range(retries):
        try:
            print(f"🧪 Attempt {attempt + 1}: Connecting to {port}...")
            connection = OBD(portstr=port, baudrate=baudrate, protocol=protocol, fast=False)
            if connection.is_connected():
                print(f"✅ Connected on attempt {attempt + 1}")
                return connection
            else:
                print(f"❌ Attempt {attempt + 1} failed: Not connected")
                connection.close()
        except Exception as e:
            print(f"⛔ Attempt {attempt + 1} error: {e}")
        time.sleep(delay)

    return None  # All attempts failed

# ----------------------------- Main Program ------------------------------

def main():
    """
    Main function to run the live OBD-II monitoring loop.
    Connects to the adapter, polls sensor data, displays in real time,
    and allows CSV export on exit.
    """
    global connection

    # Attempt to establish connection
    connection = connect_with_retry(PORT, BAUDRATE, PROTOCOL)

    if not connection or not connection.is_connected():
        print("❌ Could not establish a stable connection to the vehicle.")
        return

    print("✅ Successfully connected to OBD-II adapter.")
    print("⏳ Waiting 2 seconds for adapter to stabilize before querying data...")
    time.sleep(2)

    try:
        while True:
            # Timestamp for this reading
            timestamp = datetime.now().strftime('%H:%M:%S')

            # Query all sensor values from ECU
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

            data_log.append(row)           # Append new row
            display_table(data_log)       # Display live data
            time.sleep(REFRESH_RATE)      # Wait before next reading

    except KeyboardInterrupt:
        # Handle Ctrl+C to stop monitoring
        print("\n⏹️ Monitoring stopped by user.")
        if data_log:
            save = input("💾 Save data to CSV? (y/n): ").lower()
            if save == 'y':
                save_to_csv(data_log)

    finally:
        # Cleanly close the OBD connection
        connection.close()
        print("🔌 OBD connection closed.")

# ----------------------------- Entry Point -------------------------------

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n🛑 Program interrupted by user. Exiting gracefully...")
