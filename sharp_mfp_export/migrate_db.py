
from sharp_mfp_export import get_db_connection

def migrate():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            # Attempt to add new columns if they don't exist (Migration)
            alter_sqls = [
                "ALTER TABLE job_logs ADD COLUMN IF NOT EXISTS file_name VARCHAR(255)",
                "ALTER TABLE job_logs ADD COLUMN IF NOT EXISTS scan_type VARCHAR(100)",
                "ALTER TABLE job_logs ADD COLUMN IF NOT EXISTS destination VARCHAR(255)"
            ]
            print("Applying migration...")
            for asql in alter_sqls:
                try:
                    cursor.execute(asql)
                    print(f"Executed: {asql}")
                except Exception as e:
                    # If syntax error (e.g. MariaDB < 10.2 for IF NOT EXISTS), try without it but catch duplicate column
                    print(f"Failed with IF NOT EXISTS, trying standard add: {e}")
                    try:
                        clean_sql = asql.replace("ADD COLUMN IF NOT EXISTS", "ADD COLUMN")
                        cursor.execute(clean_sql)
                        print(f"Executed fallback: {clean_sql}")
                    except Exception as e2:
                        print(f"Ignored error (column likely exists): {e2}")
            print("Migration finished.")
    finally:
        conn.close()

if __name__ == "__main__":
    migrate()
