from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse
import urllib.parse
from collections import defaultdict
import subprocess

from flask import Flask, render_template, request, send_file, Response, stream_with_context

try:
    from openpyxl import Workbook
except ImportError as exc:  # pragma: no cover - runtime guard
    raise RuntimeError("請先安裝 openpyxl 套件：pip install openpyxl") from exc

from sharp_mfp_export import (
    PRINTERS,
    DEFAULT_USAGE_CATEGORIES,
    USAGE_CATEGORY_CONFIG,
    aggregate_joblog_reports,
    aggregate_usage_by_categories,
    format_dt,
    load_joblog_report,
    parse_month_range,
    parse_time_value,
    parse_week_range,
    fetch_latest_user_counts,
    fetch_aggregated_users_paginated,
    fetch_job_logs_by_users,
    host_tag,
    normalize_name,
)

app = Flask(__name__)


try:
    from sharp_mfp_export import PRINTER_ALIASES
except ImportError:
    PRINTER_ALIASES = {}

def _printer_label(url: str) -> str:
    if url in PRINTER_ALIASES:
        return PRINTER_ALIASES[url]
    parsed = urlparse(url)
    return parsed.netloc or url


PRINTER_CHOICES = [{"value": printer, "label": _printer_label(printer)} for printer in PRINTERS]
CATEGORY_CHOICES = [
    {"value": key, "label": config["label"]} for key, config in USAGE_CATEGORY_CONFIG.items()
]


def _selected_printers(choice: Optional[str]) -> List[str]:
    if not choice or choice == "all":
        return PRINTERS
    return [choice]


def _to_int(value: Optional[str], default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _parse_query_datetime(raw: str, label: str, errors: List[str]) -> Optional[datetime]:
    if not raw:
        return None
    parsed = parse_time_value(raw)
    if not parsed:
        errors.append(f"{label} 格式無法解析：{raw}")
    return parsed



from threading import Lock
import threading

download_lock = Lock()

@app.route("/update_data")
def update_data():
    def generate():
        # Try to acquire lock
        if not download_lock.acquire(blocking=False):
            yield "data: {\"status\": \"error\", \"message\": \"已有更新正在進行中，請稍後再試。\"}\n\n"
            return

        try:
            process = subprocess.Popen(
                ["python", "-u",  "sharp_mfp_export.py", "download"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                encoding="utf-8",
                errors="replace"
            )
            
            yield "data: {\"status\": \"start\", \"message\": \"開始更新程序...\"}\n\n"

            if process.stdout:
                for line in process.stdout:
                    line = line.strip()
                    if not line:
                        continue
                    if line.startswith("== http"):
                        yield f"data: {{\"status\": \"progress\", \"message\": \"正在處理: {line[3:]}\"}}\n\n"
                    else:
                        yield f"data: {{\"status\": \"log\", \"message\": \"{line}\"}}\n\n"

            process.wait()
            if process.returncode == 0:
                yield "data: {\"status\": \"done\", \"message\": \"更新完成！\"}\n\n"
            else:
                yield "data: {\"status\": \"error\", \"message\": \"更新過程中發生錯誤\"}\n\n"
        
        except Exception as e:
            yield f"data: {{\"status\": \"error\", \"message\": \"系統錯誤: {str(e)}\"}}\n\n"
        
        finally:
            download_lock.release()

    return Response(stream_with_context(generate()), mimetype="text/event-stream")


@app.route("/.well-known/appspecific/com.chrome.devtools.json")
def chrome_devtools():
    return {}


def _resolve_time_range_from_query(
    mode: str,
    month: str,
    week: str,
    start_raw: str,
    end_raw: str,
    errors: List[str],
) -> Tuple[Optional[datetime], Optional[datetime]]:
    raw_mode = (mode or "all").lower()
    month = month.strip()
    week = week.strip()
    start_raw = start_raw.strip()
    end_raw = end_raw.strip()

    if raw_mode == "all":
        if month:
            mode = "month"
        elif week:
            mode = "week"
        elif start_raw or end_raw:
            mode = "custom"
        else:
            mode = raw_mode
    else:
        mode = raw_mode
    if mode == "month":
        if not month:
            errors.append("請輸入欲查詢的月份 (YYYY-MM)")
            return None, None
        try:
            return parse_month_range(month)
        except ValueError as exc:
            errors.append(str(exc))
            return None, None
    if mode == "week":
        if not week:
            errors.append("請輸入欲查詢的週期 (例如 2026-W05)")
            return None, None
        try:
            return parse_week_range(week)
        except ValueError as exc:
            errors.append(str(exc))
            return None, None
    if mode == "custom":
        start_dt = _parse_query_datetime(start_raw, "開始時間", errors)
        end_dt = _parse_query_datetime(end_raw, "結束時間", errors)
        if start_dt and end_dt and end_dt < start_dt:
            errors.append("結束時間需晚於開始時間")
            return None, None
        return start_dt, end_dt

    return None, None


def _collect_job_reports(
    printers: List[str],
    user_kw: Optional[str],
    mode_kw: Optional[str],
    computer_kw: Optional[str],
    start_dt: Optional[datetime],
    end_dt: Optional[datetime],
) -> Tuple[Dict[str, Dict[str, Any]], List[str]]:
    reports: Dict[str, Dict[str, Any]] = {}
    missing: List[str] = []
    for printer in printers:
        report = load_joblog_report(printer, user_kw, mode_kw, computer_kw, start_dt, end_dt)
        if not report:
            missing.append(printer)
            continue
        report["printer"] = printer
        reports[printer] = report
    return reports, missing


def _format_usage_entries(
    stats: Dict[str, Dict[str, Any]],
    categories: List[str],
    limit: int,
    show_zero: bool,
) -> List[Dict[str, Any]]:
    ordered = []
    for record in stats.values():
        cat_values = [
            {
                "key": key,
                "label": USAGE_CATEGORY_CONFIG[key]["label"],
                "pages": record["totals"].get(key, 0),
            }
            for key in categories
        ]
        total = sum(item["pages"] for item in cat_values)
        if total == 0 and not show_zero:
            continue
        category_map = {item["key"]: item["pages"] for item in cat_values}
        ordered.append({"name": record["name"], "total": total, "categories": cat_values, "category_map": category_map})

    ordered.sort(key=lambda item: item["total"], reverse=True)
    if limit > 0:
        ordered = ordered[:limit]
    return ordered


def _build_jobs_query() -> Dict[str, Any]:
    return {
        "printer": request.args.get("printer", "all"),
        "user": request.args.get("user", "").strip(),
        "mode": request.args.get("mode", "").strip(),
        "computer": request.args.get("computer", "").strip(),
        "time_mode": request.args.get("time_mode", "all"),
        "month": request.args.get("month", "").strip(),
        "week": request.args.get("week", "").strip(),
        "start": request.args.get("start", "").strip(),
        "end": request.args.get("end", "").strip(),
        "limit": request.args.get("limit", "5"),
        "page": request.args.get("page", "1"),
        "per_page": request.args.get("per_page", "2"),
    }


def _prepare_jobs_context(query_args: Dict[str, str]) -> Dict[str, Any]:
    # Parse pagination args
    try:
        page = int(query_args.get("page", 1))
        if page < 1:
            page = 1
    except ValueError:
        page = 1
    
    try:
        per_page = int(query_args.get("per_page", 2))
        if per_page not in [2, 5, 10, 20, 50, 100]:
            per_page = 2
    except ValueError:
        per_page = 2

    try:
        limit_val = int(query_args.get("limit", 5))
    except ValueError:
        limit_val = 5

    user_kw = query_args.get("user", "").strip()
    printer_pick = query_args.get("printer", "all")
    mode_pick = query_args.get("mode", "").strip()
    computer_pick = query_args.get("computer", "").strip()
    time_mode = query_args.get("time_mode", "all")

    # Time range
    start_dt, end_dt = None, None
    mode_display = ""

    if time_mode == "month":
        month_str = query_args.get("month", "")
        start_dt, end_dt = parse_month_range(month_str)
        mode_display = f"月份: {month_str}" if month_str else ""
    elif time_mode == "week":
        week_str = query_args.get("week", "")
        start_dt, end_dt = parse_week_range(week_str)
        mode_display = f"週次: {week_str}" if week_str else ""
    elif time_mode == "custom":
        start_str = query_args.get("start", "")
        end_str = query_args.get("end", "")
        start_dt = parse_time_value(start_str)
        end_dt = parse_time_value(end_str)
        mode_display = "自訂時間"

    # 1. Fetch paginated users (aggregated stats for sorting)
    users_list, total_users = fetch_aggregated_users_paginated(
        page, per_page, 
        printer_addr=printer_pick,
        user_kw=user_kw,
        mode_kw=mode_pick,
        computer_kw=computer_pick,
        start_dt=start_dt,
        end_dt=end_dt
    )

    if not users_list:
        errors = []
        if user_kw or mode_display:
            errors = ["查無符合條件的資料。"]
        return {
            "query": query_args,
            "errors": errors,
            "results": [],
            "query_string": urllib.parse.urlencode(query_args),
            "pagination": {
                "page": page,
                "per_page": per_page,
                "total_users": 0,
                "total_pages": 0
            }
        }

    # 2. Fetch detailed logs for these users
    detailed_entries = fetch_job_logs_by_users(
        users_list,
        printer_addr=printer_pick,
        mode_kw=mode_pick,
        computer_kw=computer_pick,
        start_dt=start_dt,
        end_dt=end_dt
    )
    
    # 3. Aggregate into report format
    user_map = {} # (user, login) -> list of entries
    for entry in detailed_entries:
        key = (entry["user"], entry["login"])
        if key not in user_map:
            user_map[key] = []
        user_map[key].append(entry)
        
    final_blocks = []
    # Preserve order from users_list (which is sorted by pages DESC)
    for u in users_list:
        key = (u["user"], u["login"])
        entries = user_map.get(key, [])
        if not entries:
            continue
            
        # Calculate totals
        total_jobs = len(entries)
        total_pages = sum(e["pages"] for e in entries)
        total_bw = sum(e["bw"] for e in entries)
        total_color = sum(e["color"] for e in entries)
        
        # Printer sub-totals
        p_stats = defaultdict(lambda: {"jobs": 0, "pages": 0})
        for e in entries:
            p_addr = e["printer"]
            p_stats[p_addr]["jobs"] += 1
            p_stats[p_addr]["pages"] += e["pages"]
            
        p_summaries = []
        for addr, st in p_stats.items():
            label = host_tag(addr)
            p_summaries.append({"label": label, "jobs": st["jobs"], "pages": st["pages"]})
        p_summaries.sort(key=lambda x: x["pages"], reverse=True)

        user_display = normalize_name(u["user"], "未知")
        login_display = normalize_name(u["login"], "N/A")

        # Slice entries for display if limit is set
        display_entries = entries
        if limit_val > 0:
            display_entries = entries[:limit_val]

        final_blocks.append({
            "name": user_display,
            "login": login_display,
            "totals": {
                "jobs": total_jobs,
                "pages": total_pages,
                "bw": total_bw,
                "color": total_color
            },
            "printer_totals": p_summaries,
            "entries": display_entries
        })

    # Calc total pages for pagination
    import math
    total_pages_count = math.ceil(total_users / per_page) if per_page > 0 else 0

    return {
        "query": query_args,
        "errors": [],
        "results": final_blocks,
        "query_string": urllib.parse.urlencode(query_args),
        "pagination": {
            "page": page,
            "per_page": per_page,
            "total_users": total_users,
            "total_pages": total_pages_count,
            "has_prev": page > 1,
            "has_next": page < total_pages_count,
            "prev_num": page - 1,
            "next_num": page + 1
        }
    }





def _build_counts_query() -> Dict[str, Any]:
    return {
        "printer": request.args.get("printer", "all"),
        "user": request.args.get("user", "").strip(),
        "time_mode": request.args.get("time_mode", "all"),
        "month": request.args.get("month", "").strip(),
        "week": request.args.get("week", "").strip(),
        "start": request.args.get("start", "").strip(),
        "end": request.args.get("end", "").strip(),
        "categories": request.args.getlist("category"),
        "limit": request.args.get("limit", "0"),
        "show_zero": request.args.get("show_zero") == "on",
    }



def _prepare_counts_context(query: Dict[str, Any], export_mode: bool = False) -> Dict[str, Any]:
    # Parse pagination
    if export_mode:
        page = 1
        per_page = 0
    else:
        try:
            page = int(request.args.get("page", 1))
            if page < 1: 
                page = 1
        except ValueError:
            page = 1

        try:
            per_page = int(request.args.get("per_page", 10))
            if per_page not in [2, 5, 10, 20, 50, 100]:
                per_page = 10
        except ValueError:
            per_page = 10
    
    # Update query dict so template receives the effective value (and dropdown selects correctly)
    query["per_page"] = str(per_page)

    limit = _to_int(query["limit"], 0) # Limit per user display is legacy, we rely on pagination now.
    # In original code, limit was used in _format_usage_entries to limit *number of users*?
    # If using pagination, we don't really need 'limit' for "Top N". Pagination handles it.
    # But maybe user wants "Top 5" instead of page 1?
    # Let's keep 'limit' as a separate secondary restriction if needed, or ignore it if per_page rules.
    # The user request said "customizable limit". per_page is that custom limit.
    # 'limit' in query dict comes from form input 'limit'.
    # In the HTML, 'limit' input is "顯示筆數". 
    # But wait, User Request 2 says "一頁最多只能顯示10條，并且可以自定義".
    # This refers to "entries per page" (per_page).
    # The old "limit" input was "Show Top N".
    # I should map the old 'limit' form field to 'per_page' or just use 'per_page'.
    # I will rely on 'per_page' and ignore 'limit' for determining valid users.

    errors: List[str] = []
    start_dt, end_dt = _resolve_time_range_from_query(
        query["time_mode"], query["month"], query["week"], query["start"], query["end"], errors
    )
    printer_pick = query["printer"]
    user_kw = query["user"] or None
    requested_categories = query["categories"] or DEFAULT_USAGE_CATEGORIES
    categories = [key for key in requested_categories if key in USAGE_CATEGORY_CONFIG] or DEFAULT_USAGE_CATEGORIES
    
    # 1. Fetch paginated users (Global Aggregation) for the Summary View and Pagination Control
    # Use "all" or specific printer logic for the global view.
    # Note: If printer_pick is "all", this is truly global.
    global_users, global_total = fetch_aggregated_users_paginated(
        page, per_page, 
        printer_addr=printer_pick,
        user_kw=user_kw,
        mode_kw=None,
        computer_kw=None,
        start_dt=start_dt,
        end_dt=end_dt
    )

    # Global/Aggregated Report
    aggregated = None
    if global_users:
        global_detailed = fetch_job_logs_by_users(
            global_users,
            printer_addr=printer_pick,
            mode_kw=None,
            computer_kw=None,
            start_dt=start_dt,
            end_dt=end_dt
        )
        agg_stats = aggregate_usage_by_categories(global_detailed, categories, user_kw)
        aggregated_entries = _format_usage_entries(agg_stats, categories, 0, False)
        aggregated = {"entries": aggregated_entries} if aggregated_entries else None

    # 2. Fetch Report Per Printer
    # We must paginate EACH printer independently so that "Page 1" shows the top users for THAT printer.
    category_columns = [
        {"key": key, "label": USAGE_CATEGORY_CONFIG[key]["label"]}
        for key in categories
    ]
    
    results = []
    printers = _selected_printers(printer_pick)
    
    for printer in printers:
        label = _printer_label(printer)
        
        # Paginated users for THIS printer
        p_users, p_total = fetch_aggregated_users_paginated(
            page, per_page, 
            printer_addr=printer, # Specific printer
            user_kw=user_kw,
            mode_kw=None,
            computer_kw=None,
            start_dt=start_dt,
            end_dt=end_dt
        )
        
        if not p_users:
             # Just show empty or "No data"
             results.append({
                "printer": printer,
                "label": label,
                "file_name": "DB_Query",
                "entries": [],
             })
             continue

        # Valid users found, fetch their logs
        p_detailed = fetch_job_logs_by_users(
            p_users,
            printer_addr=printer,
            mode_kw=None,
            computer_kw=None,
            start_dt=start_dt,
            end_dt=end_dt
        )
        
        stats = aggregate_usage_by_categories(p_detailed, categories, user_kw)
        formatted_entries = _format_usage_entries(stats, categories, 0, False)
        
        results.append({
            "printer": printer,
            "label": label,
            "file_name": "DB_Query",
            "entries": formatted_entries,
        })
        
    # Calculate total pages based on Global Total (Aggregation)
    # This ensures consistency for the bottom pagination control.
    # While individual printers might have fewer pages, the "Summary" usually covers the max.
    import math
    total_pages_count = math.ceil(global_total / per_page) if per_page > 0 else 0

    return {
        "results": results,
        "errors": errors,
        "category_columns": category_columns,
        "reports": [], # unused
        "aggregated": aggregated,
        "categories": categories,
        "pagination": {
            "page": page,
            "per_page": per_page,
            "total_users": global_total,
            "total_pages": total_pages_count,
            "has_prev": page > 1,
            "has_next": page < total_pages_count,
            "prev_num": page - 1,
            "next_num": page + 1
        }
    }


def _build_leaders_query() -> Dict[str, Any]:
    return {
        "printer": request.args.get("printer", "all"),
        "user": request.args.get("user", "").strip(),
        "mode": request.args.get("mode", "").strip(),
        "computer": request.args.get("computer", "").strip(),
        "time_mode": request.args.get("time_mode", "all"),
        "month": request.args.get("month", "").strip(),
        "week": request.args.get("week", "").strip(),
        "start": request.args.get("start", "").strip(),
        "end": request.args.get("end", "").strip(),
        "limit": request.args.get("limit", "10"),
        "summary_limit": request.args.get("summary_limit", "20"),
    }


def _prepare_leaders_context(query: Dict[str, Any]) -> Dict[str, Any]:
    limit = _to_int(query["limit"], 10)
    summary_limit = _to_int(query["summary_limit"], 20)
    errors: List[str] = []
    start_dt, end_dt = _resolve_time_range_from_query(
        query["time_mode"], query["month"], query["week"], query["start"], query["end"], errors
    )
    printers = _selected_printers(query["printer"])
    user_kw = query["user"] or None
    mode_kw = query["mode"] or None
    computer_kw = query["computer"] or None

    reports, _ = _collect_job_reports(printers, user_kw, mode_kw, computer_kw, start_dt, end_dt)

    results: List[Dict[str, Any]] = []
    for printer in printers:
        label = _printer_label(printer)
        report = reports.get(printer)
        if not report:
            results.append({"printer": printer, "label": label, "error": "找不到 joblog 匯出檔，請先下載。"})
            continue
        top_users = report["top_users"]
        if limit > 0:
            top_users = top_users[:limit]
        results.append(
            {
                "printer": printer,
                "label": label,
                "file_name": report["file_path"].name,
                "totals": report["totals"],
                "top_users": top_users,
            }
        )

    aggregated = None
    if reports:
        summary = aggregate_joblog_reports(list(reports.values()))
        if summary:
            users = summary["users"]
            if summary_limit > 0:
                users = users[:summary_limit]
            aggregated = {"totals": summary["totals"], "users": users}

    all_results = {"results": results, "errors": errors, "aggregated": aggregated, "reports": list(reports.values())}

    return all_results


def _workbook_response(wb: Workbook, filename: str):
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _build_jobs_workbook(reports: List[Dict[str, Any]]) -> Workbook:
    wb = Workbook()
    first_sheet = True
    for report in reports:
        printer = report["printer"]
        title = _printer_label(printer)[:31] or "Sheet"
        ws = wb.active if first_sheet else wb.create_sheet(title)
        ws.title = title
        first_sheet = False
        ws.append(["工作ID", "開始時間", "模式", "用戶", "登入名稱", "電腦名稱", "黑白張數", "彩色張數", "總張數"])
        for entry in report.get("entries", []):
            ws.append(
                [
                    entry.get("job_id") or entry.get("account_job_id") or "?",
                    format_dt(entry.get("start")),
                    entry.get("mode") or "N/A",
                    entry.get("user_display") or entry.get("user") or "未知",
                    entry.get("login_display") or entry.get("login") or "N/A",
                    entry.get("computer") or "N/A",
                    entry.get("bw", 0),
                    entry.get("color", 0),
                    entry.get("pages", 0),
                ]
            )
    return wb


def _build_counts_workbook(results: List[Dict[str, Any]], categories: List[str]) -> Workbook:
    wb = Workbook()
    first_sheet = True
    for block in results:
        title = _printer_label(block["printer"])[:31]
        ws = wb.active if first_sheet else wb.create_sheet(title)
        ws.title = title
        first_sheet = False
        headers = ["用戶"] + [USAGE_CATEGORY_CONFIG[key]["label"] for key in categories] + ["總張數"]
        ws.append(headers)
        for item in block.get("entries", []):
            category_map = item.get("category_map", {})
            row = [item["name"]]
            for key in categories:
                row.append(category_map.get(key, 0))
            row.append(item["total"])
            ws.append(row)
    return wb






def _build_combined_counts_workbook(entries: List[Dict[str, Any]], categories: List[str]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "跨機器彙總"
    headers = ["用戶"] + [USAGE_CATEGORY_CONFIG[key]["label"] for key in categories] + ["總張數"]
    ws.append(headers)
    for item in entries:
        category_map = item.get("category_map", {})
        row = [item.get("name", "未知")]
        for key in categories:
            row.append(category_map.get(key, 0))
        row.append(item.get("total", 0))
        ws.append(row)
    return wb


def _build_leaders_workbook(results: List[Dict[str, Any]], aggregated: Optional[Dict[str, Any]]) -> Workbook:
    wb = Workbook()
    first_sheet = True
    for block in results:
        title = _printer_label(block["printer"])[:31]
        ws = wb.active if first_sheet else wb.create_sheet(title)
        ws.title = title
        first_sheet = False
        ws.append(["用戶", "登入名稱", "筆數", "黑白張數", "彩色張數", "總張數"])
        for item in block.get("top_users", []):
            ws.append([
                item["user"],
                item["login"],
                item["jobs"],
                item["bw"],
                item["color"],
                item["pages"],
            ])

    if aggregated:
        ws = wb.create_sheet("跨機器彙總")
        ws.append(["用戶", "登入名稱", "筆數", "黑白張數", "彩色張數", "總張數"])
        for item in aggregated.get("users", []):
            ws.append([
                item["user"],
                item["login"],
                item["jobs"],
                item["bw"],
                item["color"],
                item["pages"],
            ])
    return wb


@app.route("/")
def index():
    return render_template("index.html", printer_choices=PRINTER_CHOICES)


@app.route("/counts")
def counts():
    query = _build_counts_query()
    context = _prepare_counts_context(query)
    query_string = request.query_string.decode() if request.query_string else ""
    return render_template(
        "counts.html",
        printer_choices=PRINTER_CHOICES,
        category_choices=CATEGORY_CHOICES,
        query=query,
        selected_categories=context["categories"],
        results=context["results"],
        errors=context["errors"],
        category_columns=context["category_columns"],
        aggregated=context["aggregated"],
        query_string=query_string,
        pagination=context.get("pagination"),
    )


@app.route("/jobs")
def jobs():
    query = _build_jobs_query()
    context = _prepare_jobs_context(query)
    query_string = request.query_string.decode() if request.query_string else ""
    return render_template(
        "jobs.html",
        printer_choices=PRINTER_CHOICES,
        query=context["query"],
        results=context["results"],
        errors=context.get("errors", []),
        query_string=query_string,
        pagination=context.get("pagination"),
    )


@app.route("/leaders")
def leaders():
    query = _build_leaders_query()
    context = _prepare_leaders_context(query)
    query_string = request.query_string.decode() if request.query_string else ""
    return render_template(
        "leaders.html",
        printer_choices=PRINTER_CHOICES,
        query=query,
        results=context["results"],
        errors=context["errors"],
        aggregated=context["aggregated"],
        query_string=query_string,
    )


@app.route("/export/jobs")
def export_jobs():
    query = _build_jobs_query()
    context = _prepare_jobs_context(query)
    wb = _build_jobs_workbook(context["reports"])
    return _workbook_response(wb, "jobs_export.xlsx")


@app.route("/export/stats")
def export_stats():
    query = _build_counts_query()
    # query["limit"] = "0" # No longer needed
    context = _prepare_counts_context(query, export_mode=True)
    categories = context["categories"]
    wb = _build_counts_workbook(context["results"], categories)
    return _workbook_response(wb, "stats_export.xlsx")


@app.route("/export/stats_combined")
def export_stats_combined():
    query = _build_counts_query()
    # query["limit"] = "0"
    context = _prepare_counts_context(query, export_mode=True)
    categories = context["categories"]
    aggregated = context.get("aggregated") or {"entries": []}
    entries = aggregated.get("entries", [])
    wb = _build_combined_counts_workbook(entries, categories)
    return _workbook_response(wb, "stats_combined.xlsx")


@app.route("/export/leaders")
def export_leaders():
    query = _build_leaders_query()
    query["limit"] = "0"
    query["summary_limit"] = "0"
    context = _prepare_leaders_context(query)
    wb = _build_leaders_workbook(context["results"], context["aggregated"])
    return _workbook_response(wb, "leaders_export.xlsx")


if __name__ == "__main__":
    try:
        from waitress import serve
        print("Starting Waitress production server on http://0.0.0.0:5000")
        serve(app, host="0.0.0.0", port=5000, threads=8)
    except ImportError:
        print("Waitress not installed, falling back to Flask development server")
        app.run(host="0.0.0.0", port=5000, debug=False)

