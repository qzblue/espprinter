
import requests
import pymysql
import os
import sys
from sharp_mfp_export import get_db_connection

PRINTER_IP = "http://10.32.48.154"

def check_access():
    print(f"Checking access to {PRINTER_IP}...")
    try:
        resp = requests.get(PRINTER_IP, timeout=3)
        print(f"Code: {resp.status_code}")
        print("Printer seems online.")
    except Exception as e:
        print(f"Failed to connect: {e}")

def check_db():
    print("Checking DB for .154 logs...")
    conn = get_db_connection()
    with conn.cursor() as cursor:
        cursor.execute("SELECT COUNT(*) as c FROM job_logs WHERE printer_addr LIKE '%154%'")
        row = cursor.fetchone()
        print(f"Job Logs count for 154: {row['c']}")
        
        cursor.execute("SELECT printer_addr, count(*) as c FROM user_counts GROUP BY printer_addr")
        rows = cursor.fetchall()
        print("User Counts per printer:")
        for r in rows:
            print(f"  {r['printer_addr']}: {r['c']}")

if __name__ == "__main__":
    check_access()
    check_db()
