
import pymysql
import os
from sharp_mfp_export import get_db_connection

try:
    conn = get_db_connection()
    with conn.cursor() as cursor:
        cursor.execute("SELECT count(*) as total FROM job_logs")
        total = cursor.fetchone()
        
        cursor.execute("SELECT count(*) as has_file FROM job_logs WHERE file_name IS NOT NULL AND file_name != ''")
        has_file = cursor.fetchone()
        
        print(f"Total: {total['total']}, With File: {has_file['has_file']}")
        
        cursor.execute("SELECT file_name, mode FROM job_logs WHERE file_name IS NOT NULL AND file_name != '' LIMIT 5")
        rows = cursor.fetchall()
        print("Sample with file:", rows)
        
        if params := (rows):
             pass
        else:
             print("No rows with file_name found.")
    conn.close()
except Exception as e:
    print(f"Error: {e}")
