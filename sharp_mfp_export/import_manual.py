
import sys
from pathlib import Path
from sharp_mfp_export import sync_csv_to_db, init_db

# File found: MX-M5050_0502286400_20260205.csv
CSV_PATH = Path(r"c:\Users\ZZY0721\OneDrive - ESCOLA SAO PAULO\桌面\工作资料\coding\espprinter\sharp_mfp_export\exports\joblog\MX-M5050_0502286400_20260205.csv")
PRINTER_IP = "http://10.32.48.154"

def run_import():
    print(f"Importing {CSV_PATH} for {PRINTER_IP}...")
    if not CSV_PATH.exists():
        print("File not found!")
        return
    
    # Ensure tables exist
    init_db()
    
    try:
        count = sync_csv_to_db(CSV_PATH, PRINTER_IP)
        print(f"Success! Inserted/Ignored {count} rows.")
    except Exception as e:
        print(f"Import failed: {e}")

if __name__ == "__main__":
    run_import()
