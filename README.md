# Employee Work Plan

Local tool that reads the master orders Excel file (קובץ אב תוצרים) and writes a new
Excel file: **employee (מבצע) x month -> work days**.

## Run
```
pip install -r requirements.txt
python work_plan.py            # asks for: file path, start month, end month, output path
```
Or without prompts:
```
python work_plan.py --input master.xlsx --start 2026-01 --end 2026-12 --output plan.xlsx
```
On Windows you can double-click `run.bat`.

## Rules (phase 1)
- Ignores tabs `הזמנות`, `בתהליך`, `סגורות` and non-order tabs (price lists, summaries).
- Only order tabs whose customer (cell K2/J2) is TEAM NETCOMM or HP/HPE.
- Only rows whose `תוצר` is `יום עבודה` (x1) or `חצי יום` (x0.5), multiplied by `כמות תוצרים`.
- Month = month of `מועד הפעלת התוצר`. Use `--basis approval` to use `מועד אישור תוצר` instead.
- Rows with no quantity or no valid date are listed in the `שורות שדולגו` sheet.
- Output sheets: `תכנון חודשי` (the table), `פירוט` (every row behind the numbers), `שורות שדולגו`, `הגדרות`.

The source workbook contains prices and personal data, so `*.xlsx` is git-ignored. Do not commit it.
