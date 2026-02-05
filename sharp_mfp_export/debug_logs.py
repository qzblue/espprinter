
import pymysql
import os
from sharp_mfp_export import get_db_connection, SharpMFP, USERNAME, PASSWORD, OUT_DIR

def check_update_logs():
    print("=== Latest Update Logs ===")
    conn = get_db_connection()
    with conn.cursor() as cursor:
        cursor.execute("SELECT * FROM update_logs ORDER BY id DESC LIMIT 5")
        rows = cursor.fetchall()
        for r in rows:
            print(f"[{r['start_time']}] Source: {r['trigger_source']}, Status: {r['status']}")
            print(f"  Message: {r['message']}")
    conn.close()

def debug_154_download():
    print("\n=== Debugging .154 Download ===")
    base = "http://10.32.48.154"
    client = SharpMFP(base, USERNAME, PASSWORD)
    
    try:
        print(f"Logging in to {base}...")
        client.login()
        print("Login success.")
        
        jl_dir = OUT_DIR / "joblog_debug"
        os.makedirs(jl_dir, exist_ok=True)
        
        print("Downloading Job Log...")
        csv_path = client.export_joblog(jl_dir)
        print(f"Download result: {csv_path}")
        
        if csv_path:
            text = csv_path.read_text(encoding="big5", errors="replace")
            print(f"CSV Content Preview (First 500 chars):\n{text[:500]}")
            print(f"Total lines: {len(text.splitlines())}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_update_logs()
    debug_154_download()
