# -*- coding: utf-8 -*-
"""
考勤罚款统计脚本
================
功能：读取「员工刷卡记录表.xls」，自动判定每位员工每天的考勤状态，
      生成最终的罚款统计表（迟到 / 早退 / 未打卡 / 缺勤）。

用法（任选其一）：
  1) 把本脚本和「员工刷卡记录表.xls」放在同一个文件夹，双击运行（或命令行 python 考勤罚款统计.py）
  2) 命令行指定文件：python 考勤罚款统计.py "D:\\xx\\员工刷卡记录表.xls"
  3) 如需标注特殊情况（出差/送货/收货/发货/请假），再准备一个「特殊情况.xlsx」，
     列：日期 | 姓名 | 情况，放在同文件夹，脚本会自动读取并免除考核。

依赖：openpyxl、xlrd（首次运行请先：pip install openpyxl xlrd）
"""

import os
import re
import sys
import glob
from datetime import datetime, timedelta

try:
    import xlrd
except ImportError:
    print("缺少依赖 xlrd，请先运行： pip install xlrd")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("缺少依赖 openpyxl，请先运行： pip install openpyxl")
    sys.exit(1)


# ============================== 参数设置（可改） ==============================
MORNING_CUTOFF_MINUTES = 9 * 60              # 早上上班打卡截止：9:00
LATE_RATE_YUAN = 1                           # 迟到费率：每分钟 1 元
LATE_HALF_DAY_MINUTES = 30                   # 迟到 >= 30 分钟 = 半天缺勤
EVENING_CUTOFF = {"财务部": 18 * 60}          # 财务部下班截止：18:00
EVENING_CUTOFF_DEFAULT = 19 * 60             # 其他部门下班截止：19:00
NOON_MINUTES = 12 * 60                       # 上午/下午分界：12:00
LEAVE_WINDOW_END = 14 * 60                   # 上午请假判断：下午 14:00 前打卡
FINE_EARLY_LEAVE = False                     # 是否对「早退」按 1元/分钟 罚款（默认 False=只列示、人工判断）
EARLY_LEAVE_ALERT_MINUTES = 60               # 早退超过 N 分钟，提示"疑似下午请假/特殊情况"

# 外勤/送货人员名单：这些人的工作性质导致「下班不用打卡」，
# 因此只查他们是否迟到，缺下班卡不罚款。名字照此填写即可（可增删）。
NO_EVENING_PUNCH_STAFF = set()

# 特殊情况关键词：出现这些词 = 免考核（按全勤）
SPECIAL_REASONS = {"出差", "送货", "收货", "发货", "送货发货", "外勤"}


# ============================== 工具函数 ==============================

def parse_time_to_minutes(text):
    """把 '08:55' 这类文本转成"距当天 0 点的分钟数"。非时间返回 None。"""
    text = str(text).strip()
    match = re.match(r"^(\d{1,2}):(\d{2})$", text)
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    if 0 <= hour <= 23 and 0 <= minute <= 59:
        return hour * 60 + minute
    return None


def parse_cell_times(value):
    """把一格里的刷卡内容（可能用换行分隔多个时间）拆成时间列表（分钟）。"""
    if value is None:
        return []
    parts = re.split(r"[\n\r]+", str(value))
    result = []
    for part in parts:
        minutes = parse_time_to_minutes(part)
        if minutes is not None:
            result.append(minutes)
    return sorted(result)


def format_minutes(minutes):
    """把分钟数格式化成 'HH:MM'。"""
    minutes = int(minutes)
    return "%02d:%02d" % (minutes // 60, minutes % 60)


def normalize_date(text, base_date):
    """把特殊情况表里的日期（10.8 / 8 / 2026-10-08 / 10月8日 等）统一成 date 对象。"""
    if isinstance(text, datetime):
        return text.date()
    text = str(text).strip().replace("日", "").replace("号", "")
    match = re.search(r"(\d{1,4})[.\-/月](\d{1,2})", text)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        if 1 <= month <= 12 and 1 <= day <= 31:
            try:
                return base_date.replace(month=month, day=day)
            except ValueError:
                return None
    try:
        day = int(float(text))
        if 1 <= day <= 31:
            return base_date.replace(day=day)
    except (ValueError, TypeError):
        pass
    return None


# ============================== 读取刷卡记录 ==============================

def read_swipe_records(file_path):
    """读取刷卡记录表，返回 (员工列表, 起始日期)。"""
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    total_rows = sheet.nrows
    total_cols = sheet.ncols

    def cell_text(row, col):
        return str(sheet.cell_value(row, col)).strip()

    # 1) 从表头读取考勤起止日期，用于把「日号」映射成真实日期
    base_date = datetime(2026, 1, 1).date()
    for row_idx in range(min(total_rows, 6)):
        for col_idx in range(total_cols):
            text = cell_text(row_idx, col_idx)
            match = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})\s*[～~\-]\s*\d{4}-\d{1,2}-\d{1,2}", text)
            if match:
                base_date = datetime(int(match.group(1)), int(match.group(2)), int(match.group(3))).date()
                break

    employees = []
    row_idx = 0
    while row_idx < total_rows:
        # 找员工信息行：工号(4) / 姓名(11) / 部门(23)
        if (cell_text(row_idx, 4) == "工号：" and cell_text(row_idx, 10) == "姓名："
                and cell_text(row_idx, 22) == "部门："):
            employee_id = cell_text(row_idx, 5)
            name = cell_text(row_idx, 11)
            department = cell_text(row_idx, 23)

            # 下一行是 1~31 的日号行，用来建立「列 -> 日号」映射
            day_number_row = row_idx + 1
            column_to_day = {}
            if day_number_row < total_rows:
                for col_idx in range(total_cols):
                    value = sheet.cell_value(day_number_row, col_idx)
                    if isinstance(value, float) and value == int(value) and 1 <= value <= 31:
                        column_to_day[col_idx] = int(value)

            # 再往下若干行是刷卡时间
            swipe_rows = []
            cursor = row_idx + 2
            while cursor < total_rows and not (
                    cell_text(cursor, 4) == "工号：" and cell_text(cursor, 10) == "姓名："
                    and cell_text(cursor, 22) == "部门："):
                swipe_rows.append(cursor)
                cursor += 1

            employees.append({
                "id": employee_id,
                "name": name,
                "department": department,
                "column_to_day": column_to_day,
                "swipe_rows": swipe_rows,
            })
            row_idx = cursor
        else:
            row_idx += 1

    # 2) 汇总每位员工每天的刷卡时间（分钟列表）
    for employee in employees:
        day_times = {}
        for row in employee["swipe_rows"]:
            for col_idx, day in employee["column_to_day"].items():
                times = parse_cell_times(sheet.cell_value(row, col_idx))
                if times:
                    day_times.setdefault(day, []).extend(times)
        for day in day_times:
            day_times[day] = sorted(set(day_times[day]))
        employee["day_times"] = day_times

        # 3) 标记「外勤人员」：本月有上午打卡、但从未打过下班卡的人
        has_morning = False
        has_evening = False
        evening_cutoff = EVENING_CUTOFF.get(employee["department"], EVENING_CUTOFF_DEFAULT)
        for day, times in day_times.items():
            if any(t < NOON_MINUTES for t in times):
                has_morning = True
            if any(t >= evening_cutoff for t in times):
                has_evening = True
        employee["detected_external"] = has_morning and not has_evening

    return employees, base_date


# ============================== 读取特殊情况 ==============================

def read_special_cases(file_path, base_date):
    """读取特殊情况表，返回 {(姓名, 日期): 情况说明}。"""
    if not file_path or not os.path.exists(file_path):
        return {}
    special_map = {}
    if file_path.lower().endswith((".xlsx", ".xlsm")):
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active
        for row in sheet.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            if str(row[0]).strip() in ("日期", "date", "Date"):
                continue
            date_val = normalize_date(row[0], base_date)
            if date_val is None or len(row) < 2 or row[1] is None:
                continue
            name = str(row[1]).strip()
            reason = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            special_map[(name, date_val)] = reason
    else:
        with open(file_path, encoding="utf-8-sig") as f:
            for line in f:
                parts = [p.strip() for p in line.split(",")]
                if not parts or parts[0] in ("日期", "date"):
                    continue
                if len(parts) < 2:
                    continue
                date_val = normalize_date(parts[0], base_date)
                if date_val is None:
                    continue
                reason = parts[2] if len(parts) > 2 else ""
                special_map[(parts[1], date_val)] = reason
    return special_map


# ============================== 判定逻辑 ==============================

def judge_employee_day(employee, day, base_date, special_map):
    """判定某员工某天的考勤状态，返回 dict。"""
    name = employee["name"]
    department = employee["department"]
    times = employee["day_times"].get(day, [])
    date_obj = base_date + timedelta(days=day - 1)
    evening_cutoff = EVENING_CUTOFF.get(department, EVENING_CUTOFF_DEFAULT)
    is_external = employee["detected_external"] or (name in NO_EVENING_PUNCH_STAFF)

    morning_times = [t for t in times if t < NOON_MINUTES]
    afternoon_times = [t for t in times if t >= NOON_MINUTES]

    # 特殊情况（出差/送货/收货/发货/请假等）→ 免考核，按全勤
    special_reason = special_map.get((name, date_obj), "")
    is_special = any(keyword in special_reason for keyword in SPECIAL_REASONS)
    is_leave = ("请假" in special_reason) or ("调休" in special_reason)

    result = {
        "date": date_obj,
        "name": name,
        "department": department,
        "id": employee["id"],
        "morning": min(morning_times) if morning_times else None,
        "afternoon": afternoon_times if afternoon_times else [],
        "status": "",
        "detail": "",
        "fine_text": "",
        "fine_amount": 0.0,   # 迟到/早退 的元数
        "fine_half_day": 0,   # 半天 次数
        "fine_full_day": 0,   # 1天 次数
    }

    if is_special or is_leave:
        result["status"] = "特殊情况" if is_special else "请假"
        result["detail"] = special_reason if special_reason else result["status"]
        return result

    # 无任何刷卡
    if not times:
        result["status"] = "全天缺勤"
        result["detail"] = "上下班均未打卡"
        result["fine_full_day"] = 1
        result["fine_text"] = "1天"
        return result

    # 只有上午刷卡
    if morning_times and not afternoon_times:
        morning = min(morning_times)
        late_minutes = morning - MORNING_CUTOFF_MINUTES
        if late_minutes > 0:
            if late_minutes >= LATE_HALF_DAY_MINUTES:
                result["status"] = "迟到(半天)"
                result["detail"] = "迟到 %d 分钟" % late_minutes
                result["fine_half_day"] = 1
                result["fine_text"] = "半天"
            else:
                result["status"] = "迟到"
                result["detail"] = "迟到 %d 分钟" % late_minutes
                result["fine_amount"] = late_minutes * LATE_RATE_YUAN
                result["fine_text"] = "%d元" % result["fine_amount"]
            return result
        # 早上正常，看是否外勤
        if is_external:
            result["status"] = "正常(外勤)"
            result["detail"] = "外勤人员，不打下班卡"
            return result
        result["status"] = "下班未打卡"
        result["detail"] = "早上正常，下午未打卡"
        result["fine_half_day"] = 1
        result["fine_text"] = "半天"
        return result

    # 只有下午刷卡
    if not morning_times and afternoon_times:
        first_afternoon = min(afternoon_times)
        if first_afternoon <= LEAVE_WINDOW_END:
            # 下午 14:00 前打卡 = 上午请假
            result["status"] = "上午请假"
            result["detail"] = "上午未打卡，下午 %s 到岗" % format_minutes(first_afternoon)
            evening_punch = [t for t in afternoon_times if t >= evening_cutoff]
            if not evening_punch:
                if is_external:
                    return result
                result["status"] = "上午请假+下班未打卡"
                result["detail"] += "，下午未打下班卡"
                result["fine_half_day"] = 1
                result["fine_text"] = "半天"
            return result
        # 下午 14:00 后才来，且上午没卡 = 上班未打卡
        result["status"] = "上班未打卡"
        result["detail"] = "上午未打卡，下午 %s 才到岗" % format_minutes(first_afternoon)
        result["fine_half_day"] = 1
        result["fine_text"] = "半天"
        return result

    # 上下午都有刷卡
    morning = min(morning_times)
    last_afternoon = max(afternoon_times)
    late_minutes = morning - MORNING_CUTOFF_MINUTES

    if late_minutes > 0:
        if late_minutes >= LATE_HALF_DAY_MINUTES:
            result["status"] = "迟到(半天)"
            result["detail"] = "迟到 %d 分钟（>=30分钟）" % late_minutes
            result["fine_half_day"] = 1
            result["fine_text"] = "半天"
        else:
            result["status"] = "迟到"
            result["detail"] = "迟到 %d 分钟" % late_minutes
            result["fine_amount"] = late_minutes * LATE_RATE_YUAN
            result["fine_text"] = "%d元" % result["fine_amount"]
        return result

    early_minutes = evening_cutoff - last_afternoon
    if early_minutes > 0:
        result["status"] = "早退"
        result["detail"] = "早退 %d 分钟（%s 离开）" % (early_minutes, format_minutes(last_afternoon))
        if early_minutes > EARLY_LEAVE_ALERT_MINUTES:
            result["detail"] += "（超过%d分钟，疑似下午请假/特殊情况，请人工确认）" % EARLY_LEAVE_ALERT_MINUTES
        if FINE_EARLY_LEAVE:
            result["fine_amount"] = early_minutes * LATE_RATE_YUAN
            result["fine_text"] = "%d元" % result["fine_amount"]
        return result

    result["status"] = "正常"
    return result


# ============================== 生成 Excel ==============================

def write_output(output_path, employees, base_date, special_map, external_names):
    """生成罚款统计 xlsx。"""
    workbook = openpyxl.Workbook()

    header_fill = PatternFill("solid", fgColor="0070C0")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="微软雅黑", size=11, color="000000")
    red_font = Font(name="微软雅黑", size=11, color="C00000")
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    def style_header(ws, row, col_start, col_end):
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = thin_border

    def style_body(ws, row, col_start, col_end, emphasize=False):
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = red_font if emphasize else body_font
            cell.border = thin_border
            cell.alignment = center if col == 1 else left

    # ---- 汇总所有判定结果 ----
    all_records = []
    for employee in employees:
        for day in sorted(employee["day_times"].keys()):
            all_records.append(judge_employee_day(employee, day, base_date, special_map))

    late_records = [r for r in all_records if r["status"] in ("迟到", "迟到(半天)")]
    early_records = [r for r in all_records if r["status"] == "早退"]
    no_punch_records = [r for r in all_records if r["status"] in ("上班未打卡", "下班未打卡", "上午请假+下班未打卡")]
    absent_records = [r for r in all_records if r["status"] == "全天缺勤"]

    # ---- Sheet1: 罚款统计 ----
    ws1 = workbook.active
    ws1.title = "罚款统计"
    section_row = 1

    def write_section(title, records):
        nonlocal section_row
        ws1.cell(row=section_row, column=1, value=title).font = Font(name="微软雅黑", size=12, bold=True, color="0070C0")
        section_row += 1
        headers = ["日期", "姓名", "说明", "罚款"]
        for col, text in enumerate(headers, start=1):
            ws1.cell(row=section_row, column=col, value=text)
        style_header(ws1, section_row, 1, 4)
        section_row += 1
        for record in records:
            date_text = "%d.%d" % (record["date"].month, record["date"].day)
            ws1.cell(row=section_row, column=1, value=date_text)
            ws1.cell(row=section_row, column=2, value=record["name"])
            ws1.cell(row=section_row, column=3, value=record["detail"])
            ws1.cell(row=section_row, column=4, value=record["fine_text"] if record["fine_text"] else "待定")
            style_body(ws1, section_row, 1, 4, emphasize=(record["fine_text"] not in ("", "待定")))
            section_row += 1
        section_row += 1

    write_section("迟到", late_records)
    write_section("早退", early_records)
    write_section("未打卡", no_punch_records)
    write_section("缺勤", absent_records)

    for col, width in zip("ABCD", (12, 12, 30, 10)):
        ws1.column_dimensions[col].width = width
    ws1.freeze_panes = "A2"

    # ---- Sheet2: 每日明细 ----
    ws2 = workbook.create_sheet("每日明细")
    headers = ["工号", "姓名", "部门", "日期", "早上打卡", "下午打卡", "判定", "说明", "罚款"]
    for col, text in enumerate(headers, start=1):
        ws2.cell(row=1, column=col, value=text)
    style_header(ws2, 1, 1, len(headers))

    detail_row = 2
    for employee in sorted(employees, key=lambda e: (e["name"], int(e["id"] or 0))):
        for day in sorted(employee["day_times"].keys()):
            record = judge_employee_day(employee, day, base_date, special_map)
            date_obj = base_date + timedelta(days=day - 1)
            morning_text = format_minutes(record["morning"]) if record["morning"] else ""
            afternoon_text = "、".join(format_minutes(t) for t in record["afternoon"])
            ws2.cell(row=detail_row, column=1, value=employee["id"])
            ws2.cell(row=detail_row, column=2, value=employee["name"])
            ws2.cell(row=detail_row, column=3, value=employee["department"])
            ws2.cell(row=detail_row, column=4, value=date_obj.strftime("%m.%d"))
            ws2.cell(row=detail_row, column=5, value=morning_text)
            ws2.cell(row=detail_row, column=6, value=afternoon_text)
            ws2.cell(row=detail_row, column=7, value=record["status"])
            ws2.cell(row=detail_row, column=8, value=record["detail"])
            ws2.cell(row=detail_row, column=9, value=record["fine_text"])
            has_fine = record["fine_text"] not in ("", "待定", None) and record["fine_amount"] > 0
            style_body(ws2, detail_row, 1, len(headers), emphasize=has_fine)
            detail_row += 1

    for col, width in zip("ABCDEFGHI", (8, 12, 10, 10, 12, 18, 18, 32, 10)):
        ws2.column_dimensions[col].width = width
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = "A1:I%d" % (detail_row - 1)

    # ---- Sheet3: 个人汇总 ----
    ws3 = workbook.create_sheet("个人汇总")
    headers = ["姓名", "部门", "迟到次数", "迟到罚款(元)", "早退次数", "未打卡(半天)", "缺勤(天)", "合计"]
    for col, text in enumerate(headers, start=1):
        ws3.cell(row=1, column=col, value=text)
    style_header(ws3, 1, 1, len(headers))

    summary = {}
    for employee in employees:
        key = (employee["name"], employee["department"])
        if key not in summary:
            summary[key] = {"late_count": 0, "late_fine": 0.0, "early_count": 0,
                            "half_day": 0, "full_day": 0}
        for day in employee["day_times"].keys():
            record = judge_employee_day(employee, day, base_date, special_map)
            if record["status"] in ("迟到", "迟到(半天)"):
                summary[key]["late_count"] += 1
                summary[key]["late_fine"] += record["fine_amount"]
            if record["status"] == "早退":
                summary[key]["early_count"] += 1
            summary[key]["half_day"] += record["fine_half_day"]
            summary[key]["full_day"] += record["fine_full_day"]

    row = 2
    for (name, department), data in summary.items():
        ws3.cell(row=row, column=1, value=name)
        ws3.cell(row=row, column=2, value=department)
        ws3.cell(row=row, column=3, value=data["late_count"])
        ws3.cell(row=row, column=4, value=data["late_fine"])
        ws3.cell(row=row, column=5, value=data["early_count"])
        ws3.cell(row=row, column=6, value=data["half_day"])
        ws3.cell(row=row, column=7, value=data["full_day"])
        total_text = "%.0f元" % data["late_fine"]
        if data["half_day"]:
            total_text += " + %d个半天" % data["half_day"]
        if data["full_day"]:
            total_text += " + %d天" % data["full_day"]
        if data["early_count"]:
            total_text += "（另有%d次早退待定）" % data["early_count"]
        ws3.cell(row=row, column=8, value=total_text)
        style_body(ws3, row, 1, len(headers))
        row += 1

    for col, width in zip("ABCDEFGH", (12, 10, 10, 14, 10, 12, 10, 26)):
        ws3.column_dimensions[col].width = width
    ws3.freeze_panes = "A2"

    # ---- Sheet4: 说明 ----
    ws4 = workbook.create_sheet("说明")
    notes = [
        "考勤罚款统计规则说明",
        "",
        "1. 打卡时间段：早上 9:00 前上班打卡；财务部 18:00 后下班打卡，其他部门 19:00 后下班打卡。",
        "2. 早上取最早那次打卡时间，早于 9:00 即正常。",
        "3. 最早打卡晚于 9:00 = 迟到，晚几分钟扣几元（1元/分钟）。",
        "4. 迟到 30 分钟及以上 = 半天缺勤。",
        "5. 提前说明特殊情况（出差/送货/收货/发货等）= 免考核，按全勤。",
        "6. 上午请假：下午 14:00 前打卡即判定为上午请假（不计罚款），但下午仍需打下班卡。",
        "7. 下午只要打一次下班卡即可，多打无效。",
        "8. 上下班都未打卡：优先看是否标注特殊情况；无特殊情况 = 全天缺勤（1天）。",
        "9. 早退（下班早于截止时间）默认只列示、不自动罚款，请人工判断是否按 1元/分钟 扣。",
    ]
    if external_names:
        notes.append("")
        notes.append("以下人员本月从未打下班卡，已自动按「外勤」处理（只查迟到、不罚未打卡）：")
        notes.append("  " + "、".join(sorted(external_names)))
        notes.append("如名单不对，请编辑脚本顶部 NO_EVENING_PUNCH_STAFF 增删姓名。")
    notes += [
        "",
        "特殊情况表格式（可选的「特殊情况.xlsx」，放在同文件夹）：",
        "  三列：日期 | 姓名 | 情况",
        "  日期可写 10.8 / 8 / 2026-10-08；情况写 出差/送货/收货/发货/请假 等。",
    ]
    for idx, text in enumerate(notes, start=1):
        cell = ws4.cell(row=idx, column=1, value=text)
        if idx == 1:
            cell.font = Font(name="微软雅黑", size=13, bold=True, color="0070C0")
        else:
            cell.font = Font(name="微软雅黑", size=11)
    ws4.column_dimensions["A"].width = 92

    workbook.save(output_path)
    return len(all_records), len(late_records), len(early_records), len(no_punch_records), len(absent_records)


# ============================== 主流程 ==============================

def main():
    # 1) 定位刷卡记录文件
    swipe_file = None
    if len(sys.argv) > 1:
        swipe_file = sys.argv[1]
    else:
        candidates = glob.glob("员工刷卡记录表*.xls") + glob.glob("*.xls")
        if candidates:
            swipe_file = candidates[0]

    if not swipe_file or not os.path.exists(swipe_file):
        print("未找到刷卡记录文件。请把「员工刷卡记录表.xls」放到本脚本同目录，")
        print("或运行： python 考勤罚款统计.py \"文件路径\"")
        sys.exit(1)

    # 2) 定位特殊情况文件（可选）
    special_file = None
    if len(sys.argv) > 2:
        special_file = sys.argv[2]
    else:
        candidates = glob.glob("特殊情况*.xlsx")
        if candidates:
            special_file = candidates[0]

    print("正在读取刷卡记录：", swipe_file)
    employees, base_date = read_swipe_records(swipe_file)
    print("共读取员工 %d 人，考勤起始日期 %s" % (len(employees), base_date))

    # 3) 外勤人员名单（手动配置 + 自动识别）
    external_names = set(NO_EVENING_PUNCH_STAFF)
    for employee in employees:
        if employee["detected_external"]:
            external_names.add(employee["name"])
    if external_names:
        print("识别到外勤人员（不打下班卡）：%s" % "、".join(sorted(external_names)))

    special_map = read_special_cases(special_file, base_date)
    if special_map:
        print("已读取特殊情况 %d 条" % len(special_map))
    else:
        print("未提供特殊情况表（如需免除出差/送货/请假等，请准备「特殊情况.xlsx」）")

    # 4) 生成结果
    output_dir = os.path.dirname(os.path.abspath(swipe_file))
    output_path = os.path.join(output_dir, "考勤罚款统计_%s.xlsx" % base_date.strftime("%Y-%m"))
    counts = write_output(output_path, employees, base_date, special_map, external_names)

    print("生成完成：", output_path)
    print("共判定记录 %d 条；迟到 %d 条、早退 %d 条、未打卡 %d 条、缺勤 %d 条"
          % (counts[0], counts[1], counts[2], counts[3], counts[4]))


if __name__ == "__main__":
    main()
