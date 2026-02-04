from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse
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
        "limit": request.args.get("limit", "10"),
    }


def _prepare_jobs_context(query: Dict[str, Any]) -> Dict[str, Any]:
    limit = _to_int(query["limit"], 20)
    errors: List[str] = []
    start_dt, end_dt = _resolve_time_range_from_query(
        query["time_mode"], query["month"], query["week"], query["start"], query["end"], errors
    )
    printers = _selected_printers(query["printer"])
    user_kw = query["user"] or None
    mode_kw = query["mode"] or None
    computer_kw = query["computer"] or None

    reports, missing = _collect_job_reports(printers, user_kw, mode_kw, computer_kw, start_dt, end_dt)
    for printer in missing:
        errors.append(f"{_printer_label(printer)} 找不到對應的 joblog 匯出檔，請先執行下載。")

    user_groups: Dict[str, Dict[str, Any]] = {}
    for printer in printers:
        report = reports.get(printer)
        if not report:
            continue
        label = _printer_label(printer)
        for entry in report.get("entries", []):
            user_key = entry.get("user_key") or (entry.get("user") or "未知").lower()
            display_name = entry.get("user_display") or entry.get("user") or "未知"
            login_name = entry.get("login_display") or entry.get("login") or "N/A"
            group = user_groups.setdefault(
                user_key,
                {
                    "name": display_name,
                    "login": login_name,
                    "totals": {"jobs": 0, "bw": 0, "color": 0, "pages": 0},
                    "printer_totals": {},
                    "entries": [],
                },
            )
            group["totals"]["jobs"] += 1
            group["totals"]["bw"] += entry.get("bw", 0)
            group["totals"]["color"] += entry.get("color", 0)
            group["totals"]["pages"] += entry.get("pages", 0)

            printer_totals = group["printer_totals"]
            printer_stat = printer_totals.setdefault(
                label,
                {"label": label, "jobs": 0, "bw": 0, "color": 0, "pages": 0},
            )
            printer_stat["jobs"] += 1
            printer_stat["bw"] += entry.get("bw", 0)
            printer_stat["color"] += entry.get("color", 0)
            printer_stat["pages"] += entry.get("pages", 0)

            start_dt_entry = entry.get("start")
            group["entries"].append(
                {
                    "printer": label,
                    "job_id": entry.get("job_id") or entry.get("account_job_id") or "?",
                    "mode": entry.get("mode") or "N/A",
                    "start": format_dt(start_dt_entry),
                    "start_sort": start_dt_entry or datetime.min,
                    "computer": entry.get("computer") or "N/A",
                    "pages": entry.get("pages", 0),
                    "bw": entry.get("bw", 0),
                    "color": entry.get("color", 0),
                }
            )

    user_results: List[Dict[str, Any]] = []
    for data in user_groups.values():
        entries = sorted(data["entries"], key=lambda item: item["start_sort"], reverse=True)
        if limit > 0:
            entries = entries[:limit]
        printer_totals = sorted(
            data["printer_totals"].values(),
            key=lambda item: (item["pages"], item["jobs"]),
            reverse=True,
        )
        user_results.append(
            {
                "name": data["name"],
                "login": data["login"],
                "totals": data["totals"],
                "printer_totals": printer_totals,
                "entries": entries,
            }
        )

    user_results.sort(key=lambda item: (item["totals"]["pages"], item["totals"]["jobs"]), reverse=True)

    return {
        "results": user_results,
        "errors": errors,
        "reports": list(reports.values()),
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


def _prepare_counts_context(query: Dict[str, Any]) -> Dict[str, Any]:
    limit = _to_int(query["limit"], 0)
    errors: List[str] = []
    start_dt, end_dt = _resolve_time_range_from_query(
        query["time_mode"], query["month"], query["week"], query["start"], query["end"], errors
    )
    printers = _selected_printers(query["printer"])
    user_kw = query["user"] or None
    requested_categories = query["categories"] or DEFAULT_USAGE_CATEGORIES
    categories = [key for key in requested_categories if key in USAGE_CATEGORY_CONFIG] or DEFAULT_USAGE_CATEGORIES
    show_zero = query["show_zero"]

    reports, _ = _collect_job_reports(printers, user_kw, None, None, start_dt, end_dt)

    category_columns = [
        {"key": key, "label": USAGE_CATEGORY_CONFIG[key]["label"]}
        for key in categories
    ]

    results: List[Dict[str, Any]] = []
    all_entries: List[Dict[str, Any]] = []
    for printer in printers:
        label = _printer_label(printer)
        report = reports.get(printer)
        if not report:
            results.append({"printer": printer, "label": label, "error": "找不到 joblog 匯出檔，請先下載。"})
            continue

        stats = aggregate_usage_by_categories(report["entries"], categories, user_kw)
        entries = _format_usage_entries(stats, categories, limit, show_zero)
        all_entries.extend(report["entries"])
        results.append(
            {
                "printer": printer,
                "label": label,
                "file_name": report["file_path"].name,
                "entries": entries,
            }
        )

    aggregated = None
    if all_entries:
        agg_stats = aggregate_usage_by_categories(all_entries, categories, user_kw)
        aggregated_entries = _format_usage_entries(agg_stats, categories, 0, show_zero)
        aggregated = {"entries": aggregated_entries}

    return {
        "results": results,
        "errors": errors,
        "category_columns": category_columns,
        "reports": list(reports.values()),
        "aggregated": aggregated,
        "categories": categories,
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
    )


@app.route("/jobs")
def jobs():
    query = _build_jobs_query()
    context = _prepare_jobs_context(query)
    query_string = request.query_string.decode() if request.query_string else ""
    return render_template(
        "jobs.html",
        printer_choices=PRINTER_CHOICES,
        query=query,
        results=context["results"],
        errors=context["errors"],
        query_string=query_string,
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
    query["limit"] = "0"
    context = _prepare_counts_context(query)
    categories = context["categories"]
    wb = _build_counts_workbook(context["results"], categories)
    return _workbook_response(wb, "stats_export.xlsx")


@app.route("/export/stats_combined")
def export_stats_combined():
    query = _build_counts_query()
    query["limit"] = "0"
    context = _prepare_counts_context(query)
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
    app.run(host="0.0.0.0", port=5000, debug=False)

