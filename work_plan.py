#!/usr/bin/env python3
"""Monthly employee work-day plan builder.

Reads the master orders workbook (one tab per order) and produces a new Excel
file with a table of "מבצע" (employee) x month -> number of work days.

Run:  python work_plan.py
      python work_plan.py --input master.xlsx --start 2026-01 --end 2026-12 --output plan.xlsx
"""
import argparse
import datetime as dt
import logging
import os
import re
import sys
import traceback
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
CUSTOMERS = {"TEAM NETCOMM", "TEAM NETCOM", "HP", "HPE", "IBM"}

# IBM tabs: only these "תוצר NNN" products count, each unit = this many work days.
IBM_CUSTOMERS = {"IBM"}
IBM_PRODUCT_DAYS = {"068": 0.5, "069": 0.5, "001": 22.0, "123": 18.0}

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(APP_DIR, "work_plan.log")
CONFIG_NAME = "working_days.xlsx"  # config: month | working days (edit it freely)
CONFIG_PATH = os.path.join(APP_DIR, CONFIG_NAME)
OUTPUT_NAME = "work_plan.xlsx"  # always the same name; overwritten on every run
log = logging.getLogger("work_plan")

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


def ibm_day_factor(product):
    """IBM: look up the 'תוצר NNN' number, e.g. 'תוצר 068' -> 0.5 day per unit."""
    m = re.search(r"תוצר\s*(\d{3})(?!\d)", product)
    return IBM_PRODUCT_DAYS.get(m.group(1)) if m else None


# --- Israeli work calendar (used only to FILL the config file the first time) ---------------
WORK_WEEK = (6, 0, 1, 2, 3)   # Python weekday numbers: Sun, Mon, Tue, Wed, Thu
HOLIDAY_EVE_VALUE = 0.5       # an eve of a holiday (ערב חג) counts as this many working days
EVE_HOLIDAYS = ("ראש השנה", "יום כיפור", "סוכות", "פסח", "שבועות")  # eves that get the value above
# Chol HaMoed and Hoshana Rabbah are treated as regular working days.


def _israel_holidays(years):
    try:
        import holidays
    except ImportError:
        log.warning("Package 'holidays' is missing (pip install holidays): "
                    "working days will ignore Israeli holidays")
        return {}
    return dict(holidays.country_holidays("IL", years=years))


def month_calendar(month, hol):
    """Returns (working days, notes) for one month: Sun-Thu, minus holidays, eves count as half."""
    n, notes_h, notes_e = 0.0, [], []
    d = month
    while d.month == month.month:
        if d.weekday() in WORK_WEEK:
            name = hol.get(d)
            tomorrow = hol.get(d + dt.timedelta(days=1))
            eve_name = tomorrow if tomorrow in EVE_HOLIDAYS else None
            if name:
                notes_h.append(f"{d:%d/%m} {name}")
            elif eve_name:
                n += HOLIDAY_EVE_VALUE
                notes_e.append(f"{d:%d/%m} ערב {eve_name}")
            else:
                n += 1
        d += dt.timedelta(days=1)
    notes = "; ".join(filter(None, ["חג: " + ", ".join(notes_h) if notes_h else "",
                                    "ערב חג (חצי יום): " + ", ".join(notes_e) if notes_e else ""]))
    return n, notes


def ensure_config(path, first=dt.date(2024, 1, 1), last=dt.date(2028, 12, 1)):
    """Create the config workbook (month | working days) if it does not exist yet."""
    if os.path.isfile(path):
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ימי עבודה"
    ws.sheet_view.rightToLeft = True
    ws.append(["חודש", "ימי עבודה", "הערות (חגים וערבי חג, לא משפיע על החישוב)"])
    style_header(ws, 1, 3)
    hol = _israel_holidays(range(first.year, last.year + 1))
    for m in month_range(first, last):
        days, notes = month_calendar(m, hol)
        ws.append([m, days, notes])
        ws.cell(ws.max_row, 1).number_format = "MM/YYYY"
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 70
    ws.freeze_panes = "A2"
    wb.save(path)
    log.info("Created config file with the Israeli work calendar (Sun-Thu, holidays, eves): %s", path)


def read_config(path):
    """Returns {first-of-month date: working days}. Month cell may be a date or text like 03/2026."""
    cfg = {}
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    for month, days in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        if month is None or not isinstance(days, (int, float)):
            continue
        try:
            m = dt.date(month.year, month.month, 1) if isinstance(month, (dt.date, dt.datetime)) \
                else parse_month(str(month))
        except ValueError:
            log.warning("Config: ignoring unreadable month '%s'", month)
            continue
        cfg[m] = days
    return cfg


RED = PatternFill("solid", fgColor="FF9999")     # more than the month's working days
YELLOW = PatternFill("solid", fgColor="FFEB84")  # less
GREEN = PatternFill("solid", fgColor="A9D08E")   # exactly equal


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
            factor = ibm_day_factor(product) if customer in IBM_CUSTOMERS else day_factor(product)
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


def write_output(path, records, skipped, months, start, end, basis, source, config=None):
    config = config or {}
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

    # Working days per month from the config file + colouring of each employee cell.
    cfg_row = total_row + 1
    ws.cell(cfg_row, 1, "ימי עבודה בחודש (קובץ הגדרות)")
    ws.cell(cfg_row, 1).font = Font(italic=True)
    for j, m in enumerate(months, start=2):
        ws.cell(cfg_row, j, config.get(m))
        ws.cell(cfg_row, j).font = Font(italic=True)
    missing = [m.strftime("%m/%Y") for m in months if m not in config]
    if missing:
        log.warning("Months missing in %s (not coloured): %s", CONFIG_NAME, ", ".join(missing))
    for i, name in enumerate(names, start=2):
        for j, m in enumerate(months, start=2):
            val, limit = table[name].get(m, 0), config.get(m)
            if not val or limit is None:
                continue  # empty cells and months without a config value stay uncoloured
            ws.cell(i, j).fill = RED if val > limit + 1e-9 else GREEN if abs(val - limit) <= 1e-9 else YELLOW
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
        ("חישוב IBM", "תוצר 068 ו-069 = 0.5 יום, תוצר 001 = 22 ימים, תוצר 123 = 18 ימים (לכל יחידה)"),
        ("צבעים", "אדום = יותר מימי העבודה בחודש, צהוב = פחות, ירוק = בדיוק (לפי working_days.xlsx)"),
        ("נוצר ב", dt.datetime.now().strftime("%d/%m/%Y %H:%M")),
    ):
        info.append(line)
    info.column_dimensions["A"].width = 18
    info.column_dimensions["B"].width = 60

    wb.save(path)
    return len(names), sum(r["days"] for r in in_range), len(in_range)


def find_input():
    """The master workbook is assumed to sit in the folder of this script (or the current folder).
    The newest .xlsx wins; the output file and Excel temp files (~$...) are ignored."""
    for folder in dict.fromkeys((APP_DIR, os.getcwd())):
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
                 and f.lower() not in (OUTPUT_NAME, CONFIG_NAME)]
        if files:
            files.sort(key=os.path.getmtime, reverse=True)
            if len(files) > 1:
                log.info("Several Excel files found, using the newest: %s", os.path.basename(files[0]))
            return files[0]
    raise SystemExit(f"No Excel file found in {APP_DIR}. Put the master file in this folder.")


def ask(prompt, default=None):
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip().strip('"')
    return value or default


def run():
    ap = argparse.ArgumentParser(description="Monthly employee work-day plan from the master orders Excel file")
    ap.add_argument("--input", help="master Excel file (default: the newest .xlsx in this folder)")
    ap.add_argument("--start", help="first month, YYYY-MM")
    ap.add_argument("--end", help="last month, YYYY-MM")
    ap.add_argument("--output", help=f"output file (default: {OUTPUT_NAME} in this folder)")
    ap.add_argument("--basis", choices=("activation", "approval"), default="activation",
                    help="which date decides the month (default: activation = מועד הפעלת התוצר)")
    a = ap.parse_args()

    inp = a.input or find_input()
    log.info("Using source file: %s", inp)
    start = parse_month(a.start or ask("Start month (YYYY-MM)"))
    end = parse_month(a.end or ask("End month (YYYY-MM)"))
    if end < start:
        sys.exit("End month is before start month.")
    if not inp or not os.path.isfile(inp):
        sys.exit(f"File not found: {inp}")
    # The output file and the log are saved in the same folder as this script.
    out = a.output or os.path.join(APP_DIR, OUTPUT_NAME)
    log.info("input=%s start=%s end=%s output=%s basis=%s", inp, start, end, out, a.basis)

    ensure_config(CONFIG_PATH)
    config = read_config(CONFIG_PATH)

    log.info("Reading file (can take up to a minute)...")
    records, skipped, tabs = read_orders(inp, a.basis)
    log.info("Read %d usable rows, %d skipped, %d order tabs", len(records), len(skipped), tabs)
    try:
        n_emp, total, n_rows = write_output(out, records, skipped, month_range(start, end),
                                            start, end, a.basis, inp, config)
    except PermissionError:
        raise PermissionError(f"Cannot write {out}. Close the file if it is open in Excel, "
                              "or choose another output path.")
    log.info("Order tabs used: %d | employees: %d | rows: %d | work days: %g",
             tabs, n_emp, n_rows, total)
    log.info("Rows skipped (no quantity / no valid date): %d", len(skipped))
    log.info("Saved: %s", out)


def main():
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(LOG_PATH, mode="a", encoding="utf-8"),
                  logging.StreamHandler(sys.stdout)])
    log.info("==== start ==== python %s | %s", sys.version.split()[0], sys.platform)
    try:
        run()
    except SystemExit as e:
        if e.code not in (None, 0):
            log.error("Stopped: %s", e.code)
        raise
    except BaseException:
        log.error("FAILED:\n%s", traceback.format_exc())
        print(f"\nAn error occurred. Details were saved to: {LOG_PATH}")
        sys.exit(1)


if __name__ == "__main__":
    main()
