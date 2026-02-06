
from sharp_mfp_export import get_db_connection

def add_indices():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            print("Adding index on start_time...")
            try:
                cursor.execute("CREATE INDEX idx_start_time ON job_logs(start_time)")
                print("OK.")
            except Exception as e:
                print(f"Skipped (maybe exists): {e}")

            print("Adding index on user_name...")
            try:
                cursor.execute("CREATE INDEX idx_user_name ON job_logs(user_name)")
                print("OK.")
            except Exception as e:
                print(f"Skipped (maybe exists): {e}")
                
            print("Adding index on login_name...")
            try:
                cursor.execute("CREATE INDEX idx_login_name ON job_logs(login_name)")
                print("OK.")
            except Exception as e:
                print(f"Skipped (maybe exists): {e}")

    finally:
        conn.close()

if __name__ == "__main__":
    add_indices()
