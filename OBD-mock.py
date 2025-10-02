import pandas as pd
import time
from datetime import datetime

# Load sample data
df = pd.read_csv("obd_log_20250705_155130.csv")

# How often to emit a row (in seconds)
REFRESH_RATE = 1.0

def send_to_firebase(row_dict):
    """
    Placeholder function to upload a row to Firebase.
    Replace this logic with actual Firebase integration.
    """
    print(f"🔥 Sending to Firebase: {row_dict}")
    # TODO: Implement actual Firebase call here

def simulate_obd_data():
    """
    Simulates OBD-II data by emitting one row at a time
    from the sample CSV file.
    """
    print("🚗 Starting OBD data simulation...")
    for _, row in df.iterrows():
        row_dict = row.to_dict()
        row_dict["Simulated_Timestamp"] = datetime.now().strftime('%H:%M:%S')
        send_to_firebase(row_dict)
        time.sleep(REFRESH_RATE)
    print("✅ Simulation complete.")

if __name__ == "__main__":
    try:
        simulate_obd_data()
    except KeyboardInterrupt:
        print("\n🛑 Simulation stopped by user.")




