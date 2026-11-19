#!/usr/bin/env python3
"""Monthly employee work-day plan builder.

Reads the master orders workbook (one tab per order) and produces a new Excel
file with a table of "מבצע" (employee) x month -> number of work days.

Run:  python work_plan.py
      python work_plan.py --input master.xlsx --start 2026-01 --end 2026-12 --output plan.xlsx
"""
import argparse
import datetime as dt
import os
import re
import sys
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is missing. Install it with:  pip install openpyxl")

# Tabs to ignore at the start of the workbook (orders / in-progress / closed).
SKIP_TABS = {"הזמנות", "בתהליך", "סגורות"}

# Phase 1: customers (cell K2 or J2 of each order tab) to include.
# Spelling variants found in the file are listed.
CUSTOMERS = {"TEAM NETCOMM", "TEAM NETCOM", "HP", "HPE"}

HEADER_ROW = 5
FIRST_DATA_ROW = 7

COL_NAME = "מבצע"
COL_PRODUCT = "תוצר"
COL_QTY = "כמות תוצרים"
COL_ACTIVATION = "מועד הפעלת התוצר"
COL_APPROVAL = "מועד אישור תוצר"

NO_NAME = "(ללא שם מבצע)"


def clean(value):
    """Collapse whitespace so 'ציון אלון ' and 'ציון אלון' are one person."""
    return re.sub(r"\s+", " ", str(value)).strip() if value is not None else ""


def parse_month(text):
    """Accepts YYYY-MM, MM/YYYY, MM-YYYY."""
    text = text.strip()
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})", text)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
    else:
        m = re.fullmatch(r"(\d{1,2})[-/.](\d{4})", text)
        if not m:
            raise ValueError(f"Bad month '{text}'. Use YYYY-MM, for example 2026-03")
        month, year = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise ValueError(f"Bad month '{text}'")
    return dt.date(year, month, 1)


def month_range(start, end):
    months, cur = [], start
    while cur <= end:
        months.append(cur)
        cur = dt.date(cur.year + (cur.month == 12), cur.month % 12 + 1, 1)
    return months


def day_factor(product):
    """'חצי יום' = half day, 'יום עבודה' = full day, anything else = not a work-day product."""
    if "חצי יום" in product:
        return 0.5
    if "יום עבודה" in product:
        return 1.0
    return None


def valid_date(value):
    return isinstance(value, (dt.datetime, dt.date)) and value.year >= 2000


def read_orders(path, basis):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    records, skipped = [], []
    tabs_used = 0
    for ws in wb.worksheets[0:]:
        if ws.title in SKIP_TABS or ws.max_row < HEADER_ROW:
            continue
        header = {clean(c.value): c.column for c in ws[HEADER_ROW] if c.value}
        if COL_NAME not in header or COL_PRODUCT not in header:
            continue  # price lists / summary tabs
        customer = clean(ws["K2"].value or ws["J2"].value).upper()
        if customer not in CUSTOMERS:
            continue
        tabs_used += 1
        for r in range(FIRST_DATA_ROW, ws.max_row + 1):
            product = clean(ws.cell(r, header[COL_PRODUCT]).value)
            if 'סה"כ' in product:
                break  # totals block, end of this order's lines
            factor = day_factor(product)
            if factor is None:
                continue
            name = clean(ws.cell(r, header[COL_NAME]).value)
            qty = ws.cell(r, header[COL_QTY]).value if COL_QTY in header else None
            act = ws.cell(r, header[COL_ACTIVATION]).value
            appr = ws.cell(r, header[COL_APPROVAL]).value
            row_info = (ws.title, r, name, product)
            if not isinstance(qty, (int, float)) or qty <= 0:
                skipped.append(row_info + ("אין כמות תוצרים",))
                continue
            date = act if basis == "activation" else appr
            if not valid_date(date):
                skipped.append(row_info + ("אין תאריך תקין",))
                continue
            records.append({
                "tab": ws.title, "row": r, "customer": customer,
                "name": name or NO_NAME, "product": product,
                "month": dt.date(date.year, date.month, 1),
                "days": qty * factor,
                "activation": act if valid_date(act) else None,
                "approval": appr if valid_date(appr) else None,
            })
    return records, skipped, tabs_used


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_output(path, records, skipped, months, start, end, basis, source):
    in_range = [r for r in records if start <= r["month"] <= end]
    table = defaultdict(lambda: defaultdict(float))
    for r in in_range:
        table[r["name"]][r["month"]] += r["days"]
    names = sorted(table, key=lambda n: (n == NO_NAME, n))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "תכנון חודשי"
    ws.sheet_view.rightToLeft = True
    headers = [COL_NAME] + [m.strftime("%m/%Y") for m in months] + ['סה"כ']
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, name in enumerate(names, start=2):
        ws.append([name] + [table[name].get(m, 0) or None for m in months] + [None])
        last = get_column_letter(len(months) + 1)
        ws.cell(i, len(headers), f"=SUM(B{i}:{last}{i})")
    total_row = len(names) + 2
    ws.cell(total_row, 1, 'סה"כ')
    for c in range(2, len(headers) + 1):
        col = get_column_letter(c)
        ws.cell(total_row, c, f"=SUM({col}2:{col}{total_row - 1})")
    for c in range(1, len(headers) + 1):
        ws.cell(total_row, c).font = Font(bold=True)
    ws.column_dimensions["A"].width = 26
    for c in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.freeze_panes = "B2"

    det = wb.create_sheet("פירוט")
    det.sheet_view.rightToLeft = True
    det.append(["חודש", COL_NAME, "לקוח", "טאב הזמנה", "שורה", COL_PRODUCT,
                "ימי עבודה", COL_ACTIVATION, COL_APPROVAL])
    style_header(det, 1, 9)
    for r in sorted(in_range, key=lambda x: (x["month"], x["name"], x["tab"], x["row"])):
        det.append([r["month"].strftime("%m/%Y"), r["name"], r["customer"], r["tab"], r["row"],
                    r["product"], r["days"], r["activation"], r["approval"]])
        for c in (8, 9):
            det.cell(det.max_row, c).number_format = "DD/MM/YYYY"
    for col, w in zip("ABCDEFGHI", (10, 22, 14, 16, 7, 38, 11, 16, 16)):
        det.column_dimensions[col].width = w
    det.freeze_panes = "A2"
    det.auto_filter.ref = det.dimensions

    ex = wb.create_sheet("שורות שדולגו")
    ex.sheet_view.rightToLeft = True
    ex.append(["טאב הזמנה", "שורה", COL_NAME, COL_PRODUCT, "סיבה"])
    style_header(ex, 1, 5)
    for row in skipped:
        ex.append(list(row))
    for col, w in zip("ABCDE", (16, 7, 22, 38, 22)):
        ex.column_dimensions[col].width = w

    info = wb.create_sheet("הגדרות")
    info.sheet_view.rightToLeft = True
    basis_he = COL_ACTIVATION if basis == "activation" else COL_APPROVAL
    for line in (
        ("קובץ מקור", os.path.basename(source)),
        ("טווח חודשים", f"{start:%m/%Y} - {end:%m/%Y}"),
        ("חודש נקבע לפי", basis_he),
        ("לקוחות כלולים", ", ".join(sorted(CUSTOMERS))),
        ("חישוב", "כמות תוצרים x 1 ליום עבודה, x 0.5 לחצי יום"),
        ("נוצר ב", dt.datetime.now().strftime("%d/%m/%Y %H:%M")),
    ):
        info.append(line)
    info.column_dimensions["A"].width = 18
    info.column_dimensions["B"].width = 60

    wb.save(path)
    return len(names), sum(r["days"] for r in in_range), len(in_range)


def ask(prompt, default=None):
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip().strip('"')
    return value or default


def main():
    ap = argparse.ArgumentParser(description="Monthly employee work-day plan from the master orders Excel file")
    ap.add_argument("--input", help="path to the master Excel file")
    ap.add_argument("--start", help="first month, YYYY-MM")
    ap.add_argument("--end", help="last month, YYYY-MM")
    ap.add_argument("--output", help="path of the new Excel file")
    ap.add_argument("--basis", choices=("activation", "approval"), default="activation",
                    help="which date decides the month (default: activation = מועד הפעלת התוצר)")
    a = ap.parse_args()

    inp = a.input or ask("Path to the master Excel file")
    start = parse_month(a.start or ask("Start month (YYYY-MM)"))
    end = parse_month(a.end or ask("End month (YYYY-MM)"))
    if end < start:
        sys.exit("End month is before start month.")
    if not inp or not os.path.isfile(inp):
        sys.exit(f"File not found: {inp}")
    default_out = os.path.join(os.path.dirname(os.path.abspath(inp)),
                               f"work_plan_{start:%Y%m}_{end:%Y%m}.xlsx")
    out = a.output or ask("Where to save the new file", default_out)

    print("Reading file (can take up to a minute)...")
    records, skipped, tabs = read_orders(inp, a.basis)
    n_emp, total, n_rows = write_output(out, records, skipped, month_range(start, end),
                                        start, end, a.basis, inp)
    print(f"Order tabs used: {tabs} | employees: {n_emp} | rows: {n_rows} | work days: {total:g}")
    print(f"Rows skipped (no quantity / no valid date): {len(skipped)}")
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
