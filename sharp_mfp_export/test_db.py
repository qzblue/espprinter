import pymysql

config = {
    'host': '10.32.65.22',
    'user': 'printer',
    'password': 'HDtAHFahLsdkNazm',
    'database': 'printer',
    'connect_timeout': 5
}

print(f"Connecting to {config['host']}...")
try:
    conn = pymysql.connect(**config)
    print("Connection successful!")
    with conn.cursor() as cursor:
        cursor.execute("SELECT VERSION()")
        version = cursor.fetchone()
        print(f"Database version: {version[0]}")
    conn.close()
except Exception as e:
    print(f"Connection failed: {e}")
