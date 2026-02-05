# sharp_mfp_export.py
# Python 3.9+ recommended

import argparse
import csv
import os
import re
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import requests


import pymysql
import pymysql.cursors

# ========= 配置區（改這裡就好） =========
# 優先讀取環境變數 SHARP_PRINTERS (逗號分隔)
env_printers = os.getenv("SHARP_PRINTERS")
if env_printers:
    PRINTERS = [p.strip() for p in env_printers.split(",") if p.strip()]
else:
    PRINTERS = [
        "http://10.64.48.120",
        "http://10.96.48.109",
        "http://10.32.48.155"
    ]

# 列印機別名 (可根據 IP 顯示易讀名稱)
# 格式: { "http://10.64.48.120": "小學3樓大教員室列印機", ... }
PRINTER_ALIASES = {
    "http://10.64.48.120": "小學3樓大教員室列印機",
    "http://10.96.48.109": "中學3樓大教員室列印機",
    "http://10.32.48.155": "幼稚園教員室列印機"
}

# 數據庫配置
DB_CONFIG = {
    'host': os.getenv("DB_HOST", "10.32.65.22"),
    'user': os.getenv("DB_USER", "printer"),
    'password': os.getenv("DB_PASS", "HDtAHFahLsdkNazm"),
    'database': os.getenv("DB_NAME", "printer"),
    'charset': 'utf8mb4',
    'cursorclass': pymysql.cursors.DictCursor,
    'connect_timeout': 5,
    'autocommit': True
}
# =====================================

def get_db_connection():
    return pymysql.connect(**DB_CONFIG)

def init_db():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            # unique key: printer + job_id + start_time
            # start_time is critical to distinguish same job id recycled or from different days?
            # job_id itself might repeat after 9999 or similar?
            # User said "data has duplicates", "incremental logs".
            # We trust that (printer, job_id, start_time) is unique enough.
            # If Job ID is strictly unique per printer forever, (printer, job_id) is enough.
            # But safer to include time.
            sql = """
            CREATE TABLE IF NOT EXISTS job_logs (
                id INT AUTO_INCREMENT PRIMARY KEY,
                printer_addr VARCHAR(100) NOT NULL,
                job_id VARCHAR(50),
                account_job_id VARCHAR(50),
                mode VARCHAR(50),
                user_name VARCHAR(100),
                login_name VARCHAR(100),
                computer_name VARCHAR(100),
                start_time DATETIME,
                end_time DATETIME,
                bw_pages INT DEFAULT 0,
                color_pages INT DEFAULT 0,
                total_pages INT DEFAULT 0,
                file_name VARCHAR(255),
                scan_type VARCHAR(100),
                destination VARCHAR(255),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE KEY unique_job (printer_addr, job_id, start_time) 
            );
            """
            cursor.execute(sql)

            # Attempt to add new columns if they don't exist (Migration)
            alter_sqls = [
                "ALTER TABLE job_logs ADD COLUMN IF NOT EXISTS file_name VARCHAR(255)",
                "ALTER TABLE job_logs ADD COLUMN IF NOT EXISTS scan_type VARCHAR(100)",
                "ALTER TABLE job_logs ADD COLUMN IF NOT EXISTS destination VARCHAR(255)"
            ]
            for asql in alter_sqls:
                try:
                    cursor.execute(asql)
                except Exception:
                    # Ignore if column exists
                    pass
            
            # Explicitly check and add columns if missing (Robust check)
            cursor.execute("SHOW COLUMNS FROM job_logs")
            existing_cols = {row['Field'] for row in cursor.fetchall()}
            
            robust_adds = []
            if 'file_name' not in existing_cols: robust_adds.append("ADD COLUMN file_name VARCHAR(255)")
            if 'scan_type' not in existing_cols: robust_adds.append("ADD COLUMN scan_type VARCHAR(100)")
            if 'destination' not in existing_cols: robust_adds.append("ADD COLUMN destination VARCHAR(255)")
            
            for stmt in robust_adds:
                try:
                    cursor.execute(f"ALTER TABLE job_logs {stmt}")
                except Exception as e:
                    print(f"Schema update error: {e}")
            
            # User Count Table
            # User Count Table
            sql_uc = """
            CREATE TABLE IF NOT EXISTS user_counts (
                id INT AUTO_INCREMENT PRIMARY KEY,
                printer_addr VARCHAR(100) NOT NULL,
                user_name VARCHAR(100),
                print_bw INT DEFAULT 0,
                print_color INT DEFAULT 0,
                copy_bw INT DEFAULT 0,
                copy_color INT DEFAULT 0,
                other_usage INT DEFAULT 0,
                total_pages INT DEFAULT 0,
                snapshot_time DATETIME DEFAULT CURRENT_TIMESTAMP,
                INDEX idx_printer_latest (printer_addr, snapshot_time)
            );
            """
            cursor.execute(sql_uc)
            
            # Update Logs Table
            sql_log = """
            CREATE TABLE IF NOT EXISTS update_logs (
                id INT AUTO_INCREMENT PRIMARY KEY,
                trigger_source VARCHAR(20) NOT NULL,
                status VARCHAR(20) NOT NULL,
                start_time DATETIME DEFAULT CURRENT_TIMESTAMP,
                end_time DATETIME,
                message TEXT,
                INDEX idx_start_time (start_time)
            );
            """
            cursor.execute(sql_log)
    finally:
        conn.close()

def sync_csv_to_db(path: Path, printer_addr: str) -> int:
    """Read CSV path, parse it, and upsert into DB."""
    entries = _joblog_entries_from_csv_raw(path)
    if not entries:
        return 0

    conn = get_db_connection()
    inserted = 0
    try:
        with conn.cursor() as cursor:
            sql = """
            INSERT INTO job_logs (
                printer_addr, job_id, account_job_id, mode, 
                user_name, login_name, computer_name, 
                start_time, end_time, bw_pages, color_pages, total_pages,
                file_name, scan_type, destination
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                file_name = VALUES(file_name),
                scan_type = VALUES(scan_type),
                destination = VALUES(destination),
                bw_pages = VALUES(bw_pages),
                color_pages = VALUES(color_pages),
                total_pages = VALUES(total_pages)
            """
            
            # Prepare batch
            values = []
            for e in entries:
                if not e.get("start"):
                    continue

                values.append((
                    printer_addr,
                    e.get("job_id"),
                    e.get("account_job_id"),
                    e.get("mode"),
                    e.get("user"),
                    e.get("login"),
                    e.get("computer"),
                    e.get("start"),
                    e.get("end"),
                    e.get("bw", 0),
                    e.get("color", 0),
                    e.get("pages", 0),
                    e.get("file_name"),
                    e.get("scan_type"),
                    e.get("destination")
                ))
            
            if values:
                inserted = cursor.executemany(sql, values)
    finally:
        conn.close()
    
    return inserted


def sync_usercount_to_db(path: Path, printer_addr: str) -> int:
    """Parse usercount CSV and insert snapshot."""
    rows = _read_csv_rows_raw(path)
    if not rows:
        return 0

    conn = get_db_connection()
    inserted = 0
    
    # Try to parse timestamp from filename: uc_TAG_TIMESTAMP.csv
    # format: YYYYMMDD-HHMMSS
    timestamp = datetime.now()
    try:
        ts_str = path.stem.split("_")[-1]
        timestamp = datetime.strptime(ts_str, "%Y%m%d-%H%M%S")
    except (ValueError, IndexError):
        pass

    try:
        with conn.cursor() as cursor:
            sql = """
            INSERT INTO user_counts (
                printer_addr, user_name, 
                print_bw, print_color, copy_bw, copy_color, other_usage, total_pages,
                snapshot_time
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            values = []
            for row in rows:
                user_name = normalize_name(row.get("用戶名稱"), "N/A")
                usage = collect_usercount_usage(row)
                
                # Mapping keys from collect_usercount_usage (based on USAGE_CATEGORY_CONFIG labels mostly)
                # But collect_usercount_usage uses raw part before "已使用"
                # Ex: "印表機:黑白"
                
                print_bw = usage.get("印表機:黑白", 0)
                print_color = usage.get("印表機:全彩", 0)
                copy_bw = usage.get("影印:黑白", 0)
                copy_color = usage.get("影印:全彩", 0)
                
                # Sum known categories to find 'other'
                known_sum = print_bw + print_color + copy_bw + copy_color
                total = sum(usage.values())
                other = total - known_sum
                
                if total == 0:
                    continue

                values.append((
                    printer_addr, user_name,
                    print_bw, print_color, copy_bw, copy_color, other, total,
                    timestamp
                ))

            if values:
                inserted = cursor.executemany(sql, values)
    finally:
        conn.close()
    return inserted


def fetch_latest_user_counts(
    printer_addr: str,
    user_filter: Optional[str] = None,
    show_zero: bool = False,
    limit: int = 0,
    offset: int = 0
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Fetch the latest snapshot for a printer.
    Returns (results, total_count).
    """
    conn = get_db_connection()
    results = []
    total = 0
    try:
        with conn.cursor() as cursor:
            # Subquery to find latest time for this printer
            sql_base = """
            FROM user_counts 
            WHERE printer_addr = %s
            AND snapshot_time = (
                SELECT MAX(snapshot_time) FROM user_counts WHERE printer_addr = %s
            )
            """
            params_base = [printer_addr, printer_addr]
            
            if user_filter:
                sql_base += " AND user_name LIKE %s"
                params_base.append(f"%{user_filter}%")
                
            if not show_zero:
                sql_base += " AND total_pages > 0"

            # Get total count
            count_sql = f"SELECT COUNT(*) as cnt {sql_base}"
            cursor.execute(count_sql, params_base)
            total = cursor.fetchone()['cnt']
            
            if total == 0:
                return [], 0
                
            # Get Data
            sql = f"SELECT * {sql_base} ORDER BY total_pages DESC"
            params = list(params_base)
            
            if limit > 0:
                sql += " LIMIT %s OFFSET %s"
                params.extend([limit, offset])
            
            cursor.execute(sql, params)
            rows = cursor.fetchall()
            
            for r in rows:
                # Reconstruct usage dict for webapp compatibility
                # "usage_list" format: [{"label": "...", "pages": 123}, ...]
                # USAGE_CATEGORY_CONFIG labels: "印表機:黑白" etc.
                
                usage_list = []
                if r['print_bw'] > 0: usage_list.append({"label": "印表機:黑白", "pages": r['print_bw']})
                if r['print_color'] > 0: usage_list.append({"label": "印表機:全彩", "pages": r['print_color']})
                if r['copy_bw'] > 0: usage_list.append({"label": "影印:黑白", "pages": r['copy_bw']})
                if r['copy_color'] > 0: usage_list.append({"label": "影印:全彩", "pages": r['copy_color']})
                if r['other_usage'] > 0: usage_list.append({"label": "其他", "pages": r['other_usage']})

                # Need "usage": {"印表機:黑白": 123} map too?
                # webapp uses `usage_list` for display and `usage` dict for sorting/export sometimes.
                # Let's provide both.
                usage_dict = {item["label"]: item["pages"] for item in usage_list}
                
                results.append({
                    "name": r['user_name'],
                    "total": r['total_pages'],
                    "usage": usage_dict,
                    "usage_list": usage_list,
                    "snapshot_time": r['snapshot_time']
                })
    finally:
        conn.close()
    return results, total


def _build_job_logs_where_clause(
    printer_addr: Optional[str] = None,
    user_kw: Optional[str] = None,
    mode_kw: Optional[str] = None,
    computer_kw: Optional[str] = None,
    start_dt: Optional[datetime] = None,
    end_dt: Optional[datetime] = None,
    filename_kw: Optional[str] = None
) -> Tuple[str, List[Any]]:
    sql = " WHERE 1=1"
    params = []
    
    if printer_addr and printer_addr != 'all':
        sql += " AND printer_addr = %s"
        params.append(printer_addr)
    
    if user_kw:
        # Enhanced user search: search by username OR LDAP display name
        from ldap_service import search_usernames_by_display_name
        
        # Get usernames matching display name query
        ldap_matches = search_usernames_by_display_name(user_kw)
        
        if ldap_matches:
            # Search by: (username LIKE query) OR (username IN ldap_matches)
            placeholders = ", ".join(["%s"] * len(ldap_matches))
            sql += f" AND (user_name LIKE %s OR login_name LIKE %s OR user_name IN ({placeholders}) OR login_name IN ({placeholders}))"
            kw = f"%{user_kw}%"
            params.extend([kw, kw] + list(ldap_matches) + list(ldap_matches))
        else:
            # No LDAP matches, just use keyword search
            sql += " AND (user_name LIKE %s OR login_name LIKE %s)"
            kw = f"%{user_kw}%"
            params.extend([kw, kw])
        
    if mode_kw:
        sql += " AND mode LIKE %s"
        params.append(f"%{mode_kw}%")
        
    if computer_kw:
        sql += " AND computer_name LIKE %s"
        params.append(f"%{computer_kw}%")
    
    if filename_kw:
        sql += " AND (file_name LIKE %s OR scan_type LIKE %s OR destination LIKE %s)"
        kw = f"%{filename_kw}%"
        params.extend([kw, kw, kw])
        
    if start_dt:
        sql += " AND start_time >= %s"
        params.append(start_dt)
    
    if end_dt:
        sql += " AND start_time <= %s"
        params.append(end_dt)
        
    return sql, params


def fetch_aggregated_users_paginated(
    page: int,
    per_page: int,
    printer_addr: Optional[str] = None,
    user_kw: Optional[str] = None,
    mode_kw: Optional[str] = None,
    computer_kw: Optional[str] = None,
    start_dt: Optional[datetime] = None,
    end_dt: Optional[datetime] = None,
    filename_kw: Optional[str] = None
) -> Tuple[List[Dict[str, str]], int]:
    """
    Returns (user_list, total_users).
    user_list is [{"user": "...", "login": "..."}, ...]
    Ordered by total_pages DESC (top users first).
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            where_sql, params = _build_job_logs_where_clause(
                printer_addr, user_kw, mode_kw, computer_kw, start_dt, end_dt, filename_kw
            )
            
            # Count total unique users
            count_sql = f"SELECT COUNT(DISTINCT user_name, login_name) as cnt FROM job_logs {where_sql}"
            cursor.execute(count_sql, params)
            total = cursor.fetchone()['cnt']
            
            if total == 0:
                return [], 0
            
            # Fetch paginated users
            # We group by user/login and sort by SUM(total_pages) desc
            sql = f"""
            SELECT user_name, login_name, SUM(total_pages) as page_sum
            FROM job_logs
            {where_sql}
            GROUP BY user_name, login_name
            ORDER BY page_sum DESC
            """
            
            query_params = list(params)
            if per_page > 0:
                offset = (page - 1) * per_page
                sql += " LIMIT %s OFFSET %s"
                query_params.extend([per_page, offset])
            
            cursor.execute(sql, query_params)
            rows = cursor.fetchall()
            
            users = []
            for r in rows:
                users.append({
                    "user": r['user_name'] or "",
                    "login": r['login_name'] or ""
                })
                
            return users, total
    finally:
        conn.close()


def fetch_job_logs_by_users(
    users: List[Dict[str, str]],
    printer_addr: Optional[str] = None,
    mode_kw: Optional[str] = None,
    computer_kw: Optional[str] = None,
    start_dt: Optional[datetime] = None,
    end_dt: Optional[datetime] = None,
    filename_kw: Optional[str] = None
) -> List[Dict[str, Any]]:
    """Fetch all logs for specific users (for the current page view)."""
    if not users:
        return []
        
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            # Build base WHERE from filters (skipping user_kw as we filter by specific users)
            where_sql, params = _build_job_logs_where_clause(
                printer_addr, None, mode_kw, computer_kw, start_dt, end_dt, filename_kw
            )
            
            # Add user list filter
            # (user_name = u1 AND login_name = l1) OR ...
            user_conditions = []
            for u in users:
                user_conditions.append("(user_name <=> %s AND login_name <=> %s)")
                params.extend([u['user'], u['login']])
            
            if user_conditions:
                where_sql += " AND (" + " OR ".join(user_conditions) + ")"
            
            sql = f"SELECT * FROM job_logs {where_sql} ORDER BY start_time DESC"
            cursor.execute(sql, params)
            rows = cursor.fetchall()
    finally:
        conn.close()
        
    return _convert_db_rows_to_api(rows)


def _convert_db_rows_to_api(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    results = []
    for r in rows:
        user_name = r['user_name'] or ""
        login_name = r['login_name'] or "N/A"
        results.append({
            "job_id": r['job_id'],
            "account_job_id": r['account_job_id'],
            "mode": r['mode'],
            "computer": r['computer_name'],
            "user": user_name,
            "login": login_name,
            "start": r['start_time'],
            "end": r['end_time'],
            "bw": r['bw_pages'],
            "color": r['color_pages'],
            "pages": r['total_pages'],
            "user_display": normalize_name(user_name, "未知"),
            "user_key": normalize_name(user_name, "未知").lower(),
            "login_display": normalize_name(login_name, "N/A"),
            "login_key": normalize_name(login_name, "N/A").lower(),
            "printer": r['printer_addr'],
            "file_name": r.get('file_name'),
            "scan_type": r.get('scan_type'),
            "destination": r.get('destination')
        })
    return results


# 建議用環境變數放帳密，不要硬寫在檔案裡
# Windows CMD:
#   set SHARP_USER=admin
#   set SHARP_PASS=admin
USERNAME = os.getenv("SHARP_USER", "admin")
PASSWORD = os.getenv("SHARP_PASS", "admin")

# User Count 下載參數（你抓包用的是 usernum=85, del=0）
USERNUM = 85
USERCOUNT_DELETE_AFTER_SAVE = 0  # 0=不刪

# Job Log 下載參數（照你抓包）
JOBLOG_DOWNLOAD_PARAMS = {
    "format": "0",
    "order": "1",
    "selectItem": "1101111111101111111111111111111111111111111101111111111111111111",
    "date": "0",
    "delAfterSave": "0",
}

# Job Log Save：你抓包那堆 checkbox，我用「清單」方式維護
JOBLOG_CHECKBOX_ON = [
    1, 62, 3, 4, 5, 6, 63, 64, 65, 66,
    7, 8, 9, 58, 10, 11, 12, 13, 14, 15,
    16, 17, 18, 19, 20, 21, 73, 74, 22, 23,
    24, 25, 26, 27, 28, 29, 30, 67, 68, 31,
    51, 52, 32, 33, 35, 36, 37, 53, 38, 39,
    70, 40, 41, 72, 49, 50, 54, 55, 56, 57,
    71
]
JOBLOG_GGT_SELECT_116 = "59"  # 你抓包是 59


OUT_DIR = Path("./exports")
TIMEOUT = 30
RETRY = 2
SLEEP_BETWEEN_PRINTERS = 0.5
CSV_ENCODING = "big5"
CSV_ERRORS = "replace"
# =====================================


def now_ts() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def host_tag(base: str) -> str:
    u = urlparse(base)
    return u.netloc.replace(":", "_")


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def extract_hidden_value(html: str, name: str) -> str:
    """
    Extract <input name="X" value="..."> from HTML.
    Works for hidden/text inputs (Sharp pages typically use this for tokens).
    """
    m = re.search(rf'name="{re.escape(name)}"\s+value="([^"]*)"', html)
    return m.group(1) if m else ""


def request_with_retry(fn, *args, **kwargs):
    last = None
    for i in range(RETRY + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last = e
            if i < RETRY:
                time.sleep(1.0 + i)
            else:
                raise last


def safe_int(value: Optional[str]) -> int:
    if value is None:
        return 0
    cleaned = value.strip().replace(",", "")
    if not cleaned or cleaned.lower() in {"n/a", "na", "無限制"}:
        return 0
    try:
        return int(cleaned)
    except ValueError:
        try:
            return int(float(cleaned))
        except ValueError:
            return 0


def normalize_name(value: Optional[str], fallback: str = "未知") -> str:
    if value is None:
        return fallback
    cleaned = value.strip()
    return cleaned or fallback


def normalize_key(value: Optional[str], fallback: str = "未知") -> str:
    return normalize_name(value, fallback).lower()


def parse_time_value(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    cleaned = value.strip()
    if not cleaned or cleaned.upper() == "N/A":
        return None
    for fmt in (
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue
    return None


def format_dt(value: Optional[datetime]) -> str:
    return value.strftime("%Y-%m-%d %H:%M:%S") if value else "未知"


def resolve_printers(spec: Optional[str]) -> List[str]:
    if not spec or spec.lower() == "all":
        return PRINTERS
    printers: List[str] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if not part.startswith("http"):
            part = f"http://{part}"
        printers.append(part.rstrip("/"))
    return printers or PRINTERS


def latest_export_file(kind: str, printer: str) -> Optional[Path]:
    directory = OUT_DIR / ("usercount" if kind == "usercount" else "joblog")
    if not directory.exists():
        return None
    tag = host_tag(printer)
    prefix = "uc" if kind == "usercount" else "joblog"
    pattern = f"{prefix}_{tag}_*.csv"
    matches = sorted(directory.glob(pattern))
    return matches[-1] if matches else None



# In-memory cache: key -> (mtime, data)
_FILE_CACHE: Dict[str, Tuple[float, Any]] = {}

def _smart_load(path: Path, loader_func, cache_key_prefix: str = "") -> Any:
    key = f"{cache_key_prefix}:{path.absolute()}"
    try:
        stat = path.stat()
        mtime = stat.st_mtime
        size = stat.st_size
    except OSError:
        # File not found or error, just try loading (will likely fail)
        return loader_func(path)



    # Check cache
    cached = _FILE_CACHE.get(key)
    if cached:
        c_mtime, c_size, c_data = cached
        if c_mtime == mtime and c_size == size:
            return c_data

    # Load and cache
    data = loader_func(path)
    _FILE_CACHE[key] = (mtime, size, data)
    return data


def _read_csv_rows_raw(path: Path) -> List[Dict[str, str]]:
    with open(path, encoding=CSV_ENCODING, errors=CSV_ERRORS, newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader)

def read_csv_rows(path: Path) -> List[Dict[str, str]]:
    return _smart_load(path, _read_csv_rows_raw, "csv_rows")


class SharpMFP:
    def __init__(self, base: str, username: str, password: str):
        self.base = base.rstrip("/")
        self.username = username
        self.password = password
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": "Mozilla/5.0"})

    # ---------- Auth ----------
    def login(self) -> None:
        login_url = f"{self.base}/login.html?/main.html"

        # 1) GET login page to obtain fresh token2
        r = self.s.get(login_url, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()

        token2 = extract_hidden_value(r.text, "token2")
        if not token2:
            raise RuntimeError("login token2 not found (UI/firmware changed or blocked)")

        payload = {
            "ggt_textbox(10002)": self.username,   # username
            "ggt_textbox(10003)": self.password,   # password
            "ggt_select(10004)": "0",
            "action": "loginbtn",
            "token2": token2,
            "ordinate": "0",
            "ggt_hidden(10008)": "0",
        }

        r = self.s.post(login_url, data=payload, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()

        # 2) verify not redirected back to login
        t = self.s.get(f"{self.base}/main.html", timeout=TIMEOUT, allow_redirects=True)
        if "login.html" in t.url:
            raise RuntimeError(f"login failed (redirected to {t.url})")

    # ---------- User Count ----------
    def export_user_count(self, out_dir: Path) -> Path:
        """
        Flow:
          GET  /account_usercountlist_save.html  -> token1/token2
          POST /account_usercountlist_save.html  action=countsavebtn
          GET  /account_count_save.html?usernum=..&del=..
        """
        ensure_dir(out_dir)

        save_page = f"{self.base}/account_usercountlist_save.html"
        r = self.s.get(save_page, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()

        token1 = extract_hidden_value(r.text, "token1")
        token2 = extract_hidden_value(r.text, "token2")
        if not token1 or not token2:
            raise RuntimeError("usercount token1/token2 not found")

        data = {
            "action": "countsavebtn",
            "token1": token1,
            "token2": token2,
            "ordinate": "",
        }
        r = self.s.post(save_page, data=data, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()

        dl = f"{self.base}/account_count_save.html"
        r = self.s.get(
            dl,
            params={"usernum": str(USERNUM), "del": str(USERCOUNT_DELETE_AFTER_SAVE)},
            timeout=TIMEOUT,
            allow_redirects=True,
        )
        r.raise_for_status()

        fn = out_dir / f"uc_{host_tag(self.base)}_{now_ts()}.csv"
        fn.write_bytes(r.content)
        return fn

    # ---------- Job Log ----------
    def export_joblog(self, out_dir: Path) -> Path:
        """
        Flow:
          GET  /sysmgt_joblog_save.html -> token1/token2
          POST /sysmgt_joblog_save.html action=jobsavebtn + checkbox options
          GET  /joblog_download.html?...
        """
        ensure_dir(out_dir)

        page = f"{self.base}/sysmgt_joblog_save.html"
        r = self.s.get(page, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()

        token1 = extract_hidden_value(r.text, "token1")
        token2 = extract_hidden_value(r.text, "token2")
        if not token1 or not token2:
            raise RuntimeError("joblog token1/token2 not found")

        data = {
            "ggt_select(116)": JOBLOG_GGT_SELECT_116,
            "action": "jobsavebtn",
            "token1": token1,
            "token2": token2,
        }

        # 勾選欄位
        for i in JOBLOG_CHECKBOX_ON:
            data[f"ggt_checkbox({i})"] = "1"

        r = self.s.post(page, data=data, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        dl = f"{self.base}/joblog_download.html"
        r = self.s.get(dl, params=JOBLOG_DOWNLOAD_PARAMS, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()

        fn = out_dir / f"joblog_{host_tag(self.base)}_{now_ts()}.csv"
        fn.write_bytes(r.content)
        return fn


def cleanup_old_exports() -> None:
    """每個分類, 每台機器只保留最新的 2 個檔案"""
    for kind in ["usercount", "joblog"]:
        directory = OUT_DIR / kind
        if not directory.exists():
            continue

        # Group by printer tag
        # filename format: prefix_tag_timestamp.csv
        # e.g., uc_10_64_48_120_20230101-120000.csv
        files_by_tag: Dict[str, List[Path]] = defaultdict(list)
        
        for p in directory.glob("*.csv"):
            if not p.is_file():
                continue
            # extract tag: uc_<THIS_PART>_timestamp.csv
            parts = p.stem.split("_")
            if len(parts) < 3:
                continue
            # prefix is parts[0], timestamp is parts[-1], middle is tag
            tag = "_".join(parts[1:-1])
            files_by_tag[tag].append(p)

        for tag, files in files_by_tag.items():
            # Sort by name (timestamp) desc
            files.sort(key=lambda x: x.name, reverse=True)
            
            # Keep top 2, delete rest
            to_delete = files[2:]
            if to_delete:
                print(f"清理舊檔案 ({kind} - {tag}): 刪除 {len(to_delete)} 個")
                for f in to_delete:
                    try:
                        f.unlink()
                        print(f"  [DEL] {f.name}")
                    except OSError as e:
                        print(f"  [ERR] {f.name}: {e}")


def download_exports(printers: Optional[List[str]] = None) -> None:
    # Ensure DB table exists
    init_db()

    uc_dir = OUT_DIR / "usercount"
    jl_dir = OUT_DIR / "joblog"
    ensure_dir(uc_dir)
    ensure_dir(jl_dir)

    selected = printers or PRINTERS
    for base in selected:
        print(f"\n== {base} ==")
        client = SharpMFP(base, USERNAME, PASSWORD)

        try:
            request_with_retry(client.login)
            uc = request_with_retry(client.export_user_count, uc_dir)
            print("OK UC    :", uc)
            
            # Sync User Count to DB
            if uc:
                uc_count = sync_usercount_to_db(uc, base)
                print(f"DB Sync UC: Inserted {uc_count} rows")

            jl = request_with_retry(client.export_joblog, jl_dir)
            print("OK JOBLOG:", jl)
            
            # Sync to DB
            if jl:
                count = sync_csv_to_db(jl, base)
                print(f"DB Sync  : Inserted/Ignored {count} rows")

        except Exception as e:
            print("FAIL:", e)

        time.sleep(SLEEP_BETWEEN_PRINTERS)
    
    # Run cleanup after all downloads
    cleanup_old_exports()


def fetch_job_logs(
    printer_addr: Optional[str] = None,
    user_kw: Optional[str] = None,
    mode_kw: Optional[str] = None,
    computer_kw: Optional[str] = None,
    start_dt: Optional[datetime] = None,
    end_dt: Optional[datetime] = None
) -> List[Dict[str, Any]]:
    """Query job logs from MySQL"""
    conn = get_db_connection()
    rows = []
    try:
        with conn.cursor() as cursor:
            sql = "SELECT * FROM job_logs WHERE 1=1"
            params = []
            
            if printer_addr and printer_addr != 'all':
                sql += " AND printer_addr = %s"
                params.append(printer_addr)
            
            if user_kw:
                # user_name OR login_name
                sql += " AND (user_name LIKE %s OR login_name LIKE %s)"
                kw = f"%{user_kw}%"
                params.extend([kw, kw])
                
            if mode_kw:
                sql += " AND mode LIKE %s"
                params.append(f"%{mode_kw}%")
                
            if computer_kw:
                sql += " AND computer_name LIKE %s"
                params.append(f"%{computer_kw}%")
                
            if start_dt:
                sql += " AND start_time >= %s"
                params.append(start_dt)
            
            if end_dt:
                sql += " AND start_time <= %s"
                params.append(end_dt)
            
            # Order by start_time DESC
            sql += " ORDER BY start_time DESC"
            
            cursor.execute(sql, params)
            rows = cursor.fetchall()
    finally:
        conn.close()
        
    # Convert to format expected by webapp (shim)
    results = []
    for r in rows:
        user_name = r['user_name'] or ""
        login_name = r['login_name'] or "N/A"
        results.append({
            "job_id": r['job_id'],
            "account_job_id": r['account_job_id'],
            "mode": r['mode'],
            "computer": r['computer_name'],
            "user": user_name,
            "login": login_name,
            "start": r['start_time'],
            "end": r['end_time'],
            "bw": r['bw_pages'],
            "color": r['color_pages'],
            # calculated fields
            "pages": r['total_pages'],
            "user_display": normalize_name(user_name, "未知"),
            "user_key": normalize_name(user_name, "未知").lower(),
            "login_display": normalize_name(login_name, "N/A"),
            "login_key": normalize_name(login_name, "N/A").lower(),
            # for grouping if needed
            "printer": r['printer_addr']
        })
    return results


def parse_cli_time(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    parsed = parse_time_value(value)
    if not parsed:
        raise ValueError(f"無法解析時間格式: {value}")
    return parsed


def parse_month_range(value: str) -> Tuple[datetime, datetime]:
    try:
        start = datetime.strptime(value.strip() + "-01", "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"月份格式應為 YYYY-MM，例如 2025-11 (收到 {value})") from exc

    if start.month == 12:
        next_month = start.replace(year=start.year + 1, month=1, day=1)
    else:
        next_month = start.replace(month=start.month + 1, day=1)

    start = start.replace(hour=0, minute=0, second=0, microsecond=0)
    end = next_month - timedelta(seconds=1)
    return start, end


def parse_week_range(value: str) -> Tuple[datetime, datetime]:
    m = re.fullmatch(r"(\d{4})-W(\d{1,2})", value.strip())
    if not m:
        raise ValueError(f"週格式應為 YYYY-Www，例如 2025-W05 (收到 {value})")
    year = int(m.group(1))
    week = int(m.group(2))
    try:
        start = datetime.fromisocalendar(year, week, 1)
    except ValueError as exc:
        raise ValueError(f"週期參數無效: {value}") from exc
    start = start.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=7) - timedelta(seconds=1)
    return start, end


def resolve_time_range_args(
    month: Optional[str],
    week: Optional[str],
    start_raw: Optional[str],
    end_raw: Optional[str],
) -> Tuple[Optional[datetime], Optional[datetime]]:
    if month:
        return parse_month_range(month)
    if week:
        return parse_week_range(week)

    start_dt = parse_cli_time(start_raw)
    end_dt = parse_cli_time(end_raw)

    if start_dt and end_dt and end_dt < start_dt:
        raise ValueError("結束時間需晚於開始時間")
    return start_dt, end_dt


def collect_usercount_usage(row: Dict[str, str]) -> Dict[str, int]:
    usage: Dict[str, int] = {}
    for key, value in row.items():
        if not key or key == "用戶名稱":
            continue
        normalized = key.replace("：", ":")
        if "已使用" not in normalized:
            continue
        category = normalized.split("已使用", 1)[0].rstrip(":")
        if not category:
            continue
        usage[category] = usage.get(category, 0) + safe_int(value)
    return usage


def build_usercount_summary(
    file_path: Path,
    user_filter: Optional[str],
    show_zero: bool,
) -> List[Dict[str, Any]]:
    rows = read_csv_rows(file_path)
    keyword = user_filter.lower() if user_filter else None
    combined: Dict[str, Dict[str, Any]] = {}

    for row in rows:
        display_name = normalize_name(row.get("用戶名稱"), "N/A")
        key = display_name.lower()
        if keyword and keyword not in display_name.lower():
            continue

        usage = collect_usercount_usage(row)
        total = sum(usage.values())
        if total == 0 and not show_zero:
            continue

        entry = combined.setdefault(
            key,
            {"name": display_name, "total": 0, "usage": defaultdict(int)},
        )
        entry["total"] += total
        for label, pages in usage.items():
            entry["usage"][label] += pages

    summary: List[Dict[str, Any]] = []
    for record in combined.values():
        usage_dict = record["usage"]
        usage_list = [
            {"label": label, "pages": pages}
            for label, pages in sorted(usage_dict.items(), key=lambda item: item[1], reverse=True)
            if pages > 0
        ]
        summary.append({"name": record["name"], "total": record["total"], "usage": dict(usage_dict), "usage_list": usage_list})

    summary.sort(key=lambda item: item["total"], reverse=True)
    return summary


def load_usercount_summary(
    printer: str,
    user_filter: Optional[str],
    show_zero: bool,
) -> Optional[Dict[str, Any]]:
    file_path = latest_export_file("usercount", printer)
    if not file_path:
        return None
    items = build_usercount_summary(file_path, user_filter, show_zero)
    return {"printer": printer, "file_path": file_path, "items": items}


def summarize_usercount(
    printer: str,
    user_filter: Optional[str],
    limit: int,
    show_zero: bool,
) -> None:
    # Use DB
    summary, _ = fetch_latest_user_counts(printer, user_filter, show_zero)
    print(f"來源: MySQL DB (最新快照)")

    if not summary:
        print("找不到符合條件的用戶。")
        return

    max_rows = len(summary) if limit <= 0 else min(limit, len(summary))
    for item in summary[:max_rows]:
        name = item["name"]
        total = item["total"]
        usage_list = item["usage_list"]
        snippet = ", ".join(f"{part['label']}:{part['pages']}" for part in usage_list[:4]) or "0"
        print(f"- {name}: 總張數 {total} | {snippet}")


def cmd_counts(args: argparse.Namespace) -> None:
    printers = resolve_printers(args.printer)
    for base in printers:
        print(f"\n== {base} ==")
        summarize_usercount(base, args.user, args.limit, args.show_zero)



def _joblog_entries_from_csv_raw(path: Path) -> List[Dict[str, Any]]:
    entries: List[Dict[str, Any]] = []
    for row in read_csv_rows(path):
        user_display = normalize_name(row.get("用戶名稱"), "未知")
        login_display = normalize_name(row.get("登入名稱"), "N/A")
        entry = {
            "job_id": row.get("工作ID") or row.get("Job ID"),
            "account_job_id": row.get("帳戶工作ID") or row.get("Account Job ID"),
            "mode": row.get("工作模式") or row.get("Job Mode") or row.get("Mode"),
            "computer": row.get("電腦名稱") or row.get("Computer Name"),
            "user": row.get("用戶名稱") or row.get("User Name"),
            "login": row.get("登入名稱") or row.get("Login Name"),
            "start": parse_time_value(row.get("開始日期") or row.get("Start Date")),
            "end": parse_time_value(row.get("完成日期") or row.get("Completion Date")),
            "bw": safe_int(row.get("黑白總張數")),
            "color": safe_int(row.get("全彩總張數")),
            "file_name": row.get("檔案名稱"),
            "scan_type": row.get("傳送類型"),
            "destination": row.get("直接位址"),
        }
        entry["pages"] = entry["bw"] + entry["color"]
        entry["user_display"] = user_display
        entry["user_key"] = user_display.lower()
        entry["login_display"] = login_display
        entry["login_key"] = login_display.lower()
        entries.append(entry)
    return entries

def joblog_entries_from_csv(path: Path) -> List[Dict[str, Any]]:
    return _smart_load(path, _joblog_entries_from_csv_raw, "joblog_parsed")


def filter_joblog_entries(
    entries: List[Dict[str, Any]],
    user_kw: Optional[str],
    mode_kw: Optional[str],
    computer_kw: Optional[str],
    start_dt: Optional[datetime],
    end_dt: Optional[datetime],
) -> List[Dict[str, Any]]:
    filtered: List[Dict[str, Any]] = []
    user_kw_lower = user_kw.lower() if user_kw else None
    mode_kw_lower = mode_kw.lower() if mode_kw else None
    computer_kw_lower = computer_kw.lower() if computer_kw else None

    for entry in entries:
        user_name = (entry.get("user") or "").lower()
        login_name = (entry.get("login") or "").lower()
        mode_name = (entry.get("mode") or "").lower()
        computer_name = (entry.get("computer") or "").lower()
        start = entry.get("start")

        if user_kw_lower and user_kw_lower not in user_name and user_kw_lower not in login_name:
            continue
        if mode_kw_lower and mode_kw_lower not in mode_name:
            continue
        if computer_kw_lower and computer_kw_lower not in computer_name:
            continue
        if start_dt and (not start or start < start_dt):
            continue
        if end_dt and (not start or start > end_dt):
            continue

        filtered.append(entry)

    return filtered


def aggregate_joblog_reports(reports: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not reports:
        return None

    totals = {"jobs": 0, "bw": 0, "color": 0, "pages": 0}
    user_stats = defaultdict(lambda: {"jobs": 0, "bw": 0, "color": 0, "pages": 0, "user": "", "login": ""})

    for report in reports:
        rt = report.get("totals") or {}
        totals["jobs"] += rt.get("jobs", 0)
        totals["bw"] += rt.get("bw", 0)
        totals["color"] += rt.get("color", 0)
        totals["pages"] += rt.get("pages", 0)

        for entry in report.get("entries", []):
            key = (entry.get("user_key"), entry.get("login_key"))
            user_stats[key]["jobs"] += 1
            user_stats[key]["bw"] += entry.get("bw", 0)
            user_stats[key]["color"] += entry.get("color", 0)
            user_stats[key]["pages"] += entry.get("pages", 0)
            user_stats[key]["user"] = entry.get("user_display") or normalize_name(entry.get("user"))
            user_stats[key]["login"] = entry.get("login_display") or normalize_name(entry.get("login"), "N/A")

    users = [
        {
            "user": data["user"],
            "login": data["login"],
            "jobs": data["jobs"],
            "bw": data["bw"],
            "color": data["color"],
            "pages": data["pages"],
        }
        for data in sorted(user_stats.values(), key=lambda item: (item["pages"], item["jobs"]), reverse=True)
    ]

    return {"totals": totals, "users": users}


def _aggregate_entries_to_report(entries: List[Dict[str, Any]]) -> Dict[str, Any]:
    total_jobs = len(entries)
    total_bw = sum(entry["bw"] for entry in entries)
    total_color = sum(entry["color"] for entry in entries)
    total_pages = sum(entry["pages"] for entry in entries)

    pstats = defaultdict(lambda: {"jobs": 0, "bw": 0, "color": 0, "pages": 0, "user": "", "login": ""})
    for entry in entries:
        key = (entry.get("user_key"), entry.get("login_key"))
        pstats[key]["jobs"] += 1
        pstats[key]["bw"] += entry["bw"]
        pstats[key]["color"] += entry["color"]
        pstats[key]["pages"] += entry["pages"]
        pstats[key]["user"] = entry.get("user_display") or normalize_name(entry.get("user"))
        pstats[key]["login"] = entry.get("login_display") or normalize_name(entry.get("login"), "N/A")

    top_users = [
        {
            "user": data["user"],
            "login": data["login"],
            "jobs": data["jobs"],
            "bw": data["bw"],
            "color": data["color"],
            "pages": data["pages"],
        }
        for data in sorted(pstats.values(), key=lambda item: (item["pages"], item["jobs"]), reverse=True)
    ]

    recent = sorted(entries, key=lambda entry: entry.get("start") or datetime.min, reverse=True)

    return {
        "entries": entries,
        "top_users": top_users,
        "recent": recent,
        "totals": {
            "jobs": total_jobs,
            "bw": total_bw,
            "color": total_color,
            "pages": total_pages,
        },
    }

def build_joblog_report(
    file_path: Path,
    user_kw: Optional[str],
    mode_kw: Optional[str],
    computer_kw: Optional[str],
    start_dt: Optional[datetime],
    end_dt: Optional[datetime],
) -> Dict[str, Any]:
    # Backward compatibility for CLI or file-based usage
    entries = joblog_entries_from_csv(file_path)
    filtered = filter_joblog_entries(entries, user_kw, mode_kw, computer_kw, start_dt, end_dt)
    report = _aggregate_entries_to_report(filtered)
    report["file_path"] = file_path
    return report


def load_joblog_report(
    printer: str,
    user_kw: Optional[str],
    mode_kw: Optional[str],
    computer_kw: Optional[str],
    start_dt: Optional[datetime],
    end_dt: Optional[datetime],
) -> Optional[Dict[str, Any]]:
    # New DB-based implementation
    # 1. Fetch from DB
    entries = fetch_job_logs(printer, user_kw, mode_kw, computer_kw, start_dt, end_dt)
    
    # 2. Aggregate
    report = _aggregate_entries_to_report(entries)
    
    # 3. Add metadata
    report["printer"] = printer
    report["file_path"] = Path("MySQL_DB") 
    return report


def print_joblog_report(report: Dict[str, Any], limit: int, top: int) -> None:
    file_path = report["file_path"]
    print(f"使用檔案: {file_path.name}")
    if not report["entries"]:
        print("找不到符合條件的紀錄。")
        return

    totals = report["totals"]
    print(
        f"符合 {totals['jobs']} 筆，總張數 {totals['pages']} (黑白 {totals['bw']} / 彩色 {totals['color']})"
    )

    top_users = report["top_users"]
    top_n = len(top_users) if top <= 0 else min(top, len(top_users))
    print("主要使用者：")
    for data in top_users[:top_n]:
        user = data["user"]
        login = data["login"]
        login_part = f" / {login}" if login and login != user else ""
        print(
            f"- {user}{login_part}: {data['jobs']} 筆, {data['pages']} 張 (黑白 {data['bw']} / 彩色 {data['color']})"
        )

    recent = report["recent"]
    max_rows = len(recent) if limit <= 0 else min(limit, len(recent))
    print("最新紀錄：")
    for entry in recent[:max_rows]:
        job_id = entry.get("job_id") or entry.get("account_job_id") or "?"
        user = entry.get("user") or "未知"
        login = entry.get("login") or "N/A"
        mode = entry.get("mode") or "N/A"
        computer = entry.get("computer") or "N/A"
        start = format_dt(entry.get("start"))
        print(
            f"- #{job_id} {start} | {mode} | {user} / {login} | {entry['pages']} 張 (黑白 {entry['bw']}, 彩色 {entry['color']}) | 電腦: {computer}"
        )


def print_aggregated_summary(summary: Dict[str, Any], limit: int) -> None:
    totals = summary["totals"]
    print("\n== 匯總 (全部列印機) ==")
    print(f"總筆數 {totals['jobs']} ，張數 {totals['pages']} (黑白 {totals['bw']} / 彩色 {totals['color']})")

    users = summary["users"]
    slice_count = len(users) if limit <= 0 else min(limit, len(users))
    if not slice_count:
        print("目前沒有可用的彙總資料。")
        return

    print("跨機器主要使用者：")
    for item in users[:slice_count]:
        login_part = f" / {item['login']}" if item["login"] and item["login"] != item["user"] else ""
        print(
            f"- {item['user']}{login_part}: {item['jobs']} 筆, {item['pages']} 張 (黑白 {item['bw']} / 彩色 {item['color']})"
        )


USAGE_CATEGORY_CONFIG = {
    "printer_bw": {"label": "列印:黑白", "color": "bw", "mode": "print"},
    "printer_color": {"label": "列印:全彩", "color": "color", "mode": "print"},
    "copy_bw": {"label": "影印:黑白", "mode": "copy", "color": "bw"},
    "copy_color": {"label": "影印:全彩", "mode": "copy", "color": "color"},
}
DEFAULT_USAGE_CATEGORIES = list(USAGE_CATEGORY_CONFIG.keys())


def determine_mode_kind(value: Optional[str]) -> str:
    if not value:
        return "other"
    lowered = value.lower()
    if "列印" in value or "print" in lowered:
        return "print"
    if "影印" in value or "copy" in lowered:
        return "copy"
    return "other"


def aggregate_usage_by_categories(
    entries: List[Dict[str, Any]],
    categories: List[str],
    user_filter: Optional[str],
) -> Dict[str, Dict[str, Any]]:
    stats: Dict[str, Dict[str, Any]] = {}
    keyword = user_filter.lower() if user_filter else None
    active_categories = [c for c in categories if c in USAGE_CATEGORY_CONFIG] or DEFAULT_USAGE_CATEGORIES

    for entry in entries:
        display_name = entry.get("user_display") or normalize_name(entry.get("user"))
        if keyword and keyword not in display_name.lower():
            continue

        key = entry.get("user_key") or display_name.lower()
        record = stats.setdefault(
            key,
            {
                "name": display_name,
                "username": entry.get("user_key") or "",  # Add username field
                "totals": {cat: 0 for cat in active_categories},
                "total": 0,
            },
        )
        mode_kind = determine_mode_kind(entry.get("mode"))
        for cat in active_categories:
            conf = USAGE_CATEGORY_CONFIG.get(cat)
            if not conf or conf["mode"] != mode_kind:
                continue
            value = entry["bw"] if conf["color"] == "bw" else entry["color"]
            if value <= 0:
                continue
            record["totals"][cat] = record["totals"].get(cat, 0) + value
            record["total"] += value

    return stats


def log_update_event(source: str, status: str, message: str, log_id: int = 0) -> int:
    """
    Log an update event to DB.
    If log_id is 0, creates a new record (start).
    If log_id > 0, updates the existing record (end).
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            if log_id == 0:
                # Start new log
                sql = "INSERT INTO update_logs (trigger_source, status, message, start_time) VALUES (%s, %s, %s, NOW())"
                cursor.execute(sql, (source, status, message))
                lid = cursor.lastrowid
                print(f"DEBUG: Log Inserted ID={lid}")
                return lid
            else:
                # Update existing log
                print(f"DEBUG: Updating Log ID={log_id} to {status}")
                sql = "UPDATE update_logs SET status=%s, message=%s, end_time=NOW() WHERE id=%s"
                cursor.execute(sql, (status, message, log_id))
                return log_id
    except Exception as e:
        print(f"Logging failed: {e}")
        return 0
    finally:
        conn.close()


def cmd_download(args: argparse.Namespace) -> None:
    source = getattr(args, "source", "manual")
    log_id = log_update_event(source, "running", "開始下載更新...", 0)
    print(f"DEBUG: cmd_download started with log_id={log_id}")
    
    try:
        download_exports(resolve_printers(args.printer))
        print("DEBUG: download_exports finished, updating log...")
        log_update_event(source, "success", "更新成功完成", log_id)
        print("DEBUG: log updated to success")
    except Exception as e:
        print(f"DEBUG: download_exports failed: {e}")
        log_update_event(source, "error", f"更新失敗: {str(e)}", log_id)
        raise e


def cmd_counts(args: argparse.Namespace) -> None:
    printers = resolve_printers(args.printer)
    for base in printers:
        file_path = latest_export_file("usercount", base)
        print(f"\n== {base} ==")
        if not file_path:
            print("找不到對應的 usercount 檔案，請先執行 download。")
            continue
        summarize_usercount(base, file_path, args.user, args.limit, args.show_zero)


def cmd_jobs(args: argparse.Namespace) -> None:
    try:
        start_dt, end_dt = resolve_time_range_args(args.month, args.week, args.start, args.end)
    except ValueError as exc:
        print(exc)
        return

    printers = resolve_printers(args.printer)
    aggregated_sources: List[Dict[str, Any]] = []
    for base in printers:
        file_path = latest_export_file("joblog", base)
        print(f"\n== {base} ==")
        if not file_path:
            print("找不到對應的 joblog 檔案，請先執行 download。")
            continue

        report = build_joblog_report(file_path, args.user, args.mode, args.computer, start_dt, end_dt)
        report["printer"] = base
        aggregated_sources.append(report)
        print_joblog_report(report, args.limit, args.top)

    summary = aggregate_joblog_reports(aggregated_sources)
    if summary:
        print_aggregated_summary(summary, args.summary_limit)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Sharp MFP 匯出與查詢工具")
    sub = parser.add_subparsers(dest="command")

    download_parser = sub.add_parser("download", help="下載 usercount 與 joblog 匯出檔")
    download_parser.add_argument("--printer", help="指定列印機 IP (逗號分隔)，預設 all", default="all")
    download_parser.add_argument("--source", help="觸發來源 (manual/auto)", default="manual")
    download_parser.set_defaults(func=cmd_download)

    count_parser = sub.add_parser("counts", help="查詢用戶列印數量 (usercount)")
    count_parser.add_argument("--printer", default="all", help="指定列印機 IP 或 all")
    count_parser.add_argument("--user", help="用戶名稱關鍵字")
    count_parser.add_argument("--limit", type=int, default=10, help="顯示前幾名用戶 (<=0 表示全部)")
    count_parser.add_argument("--show-zero", action="store_true", help="同時列出 0 張的用戶")
    count_parser.set_defaults(func=cmd_counts)

    jobs_parser = sub.add_parser("jobs", help="查詢列印 / 影印紀錄 (joblog)")
    jobs_parser.add_argument("--printer", default="all", help="指定列印機 IP 或 all")
    jobs_parser.add_argument("--user", help="用戶或登入名稱關鍵字")
    jobs_parser.add_argument("--mode", help="工作模式關鍵字 (例如: 列印、影印)")
    jobs_parser.add_argument("--computer", help="電腦名稱關鍵字")
    jobs_parser.add_argument("--month", help="指定月份 (YYYY-MM)")
    jobs_parser.add_argument("--week", help="指定週 (YYYY-Www，例如 2026-W05)")
    jobs_parser.add_argument("--start", help="開始時間 (YYYY-MM-DD 或 YYYY-MM-DD HH:MM)")
    jobs_parser.add_argument("--end", help="結束時間 (YYYY-MM-DD 或 YYYY-MM-DD HH:MM)")
    jobs_parser.add_argument("--limit", type=int, default=10, help="顯示最新幾筆紀錄 (<=0 表示全部)")
    jobs_parser.add_argument("--top", type=int, default=5, help="摘要中顯示的使用者數 (<=0 表示全部)")
    jobs_parser.add_argument("--summary-limit", type=int, default=10, help="跨列印機彙總顯示的使用者數 (<=0 表示全部)")
    jobs_parser.set_defaults(func=cmd_jobs)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    if not getattr(args, "command", None):
        download_exports()
        return

    args.func(args)


if __name__ == "__main__":
    # Force UTF-8 output for Windows
    if sys.stdout.encoding.lower() != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except AttributeError:
            # Python < 3.7
            import codecs
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer)
            
    init_db()  # Ensure DB schema is up-to-date
    main()

PRINTER_ALIASES = {
    "http://10.64.48.120": "小學3樓大教員室列印機",
    "http://10.96.48.109": "小學5樓教員室列印機",
    "http://10.32.48.155": "中學1樓教務處列印機",
    "http://10.32.48.154": "中學105教員室列印機"
}
