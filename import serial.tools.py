import serial.tools.list_ports
import re

def extract_mac_from_hwid(hwid: str) -> str:
    """
    Extracts a 12-digit MAC-like Bluetooth address from the HWID string.
    Returns it in XX:XX:XX:XX:XX:XX format if found.
    """
    match = re.search(r'&([0-9A-F]{12})', hwid.upper())
    if match:
        mac_raw = match.group(1)
        mac = ":".join(mac_raw[i:i+2] for i in range(0, 12, 2))
        return mac
    return "Unknown"

def list_ports_with_mac():
    """
    Lists all Bluetooth COM ports and attempts to extract MAC addresses.
    """
    ports = serial.tools.list_ports.comports()
    print("📡 Scanning for Bluetooth Serial Ports with MAC addresses:\n")
    
    found = False
    for port in ports:
        if "Bluetooth" in port.description or "Serial over Bluetooth" in port.description:
            found = True
            mac = extract_mac_from_hwid(port.hwid)
            print(f"🔌 Port:        {port.device}")
            print(f"📋 Description: {port.description}")
            print(f"🆔 HWID:        {port.hwid}")
            print(f"🔗 MAC Address: {mac}")
            print("-" * 60)

    if not found:
        print("🚫 No Bluetooth serial ports found.")

if __name__ == "__main__":
    list_ports_with_mac()
