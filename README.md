# Employee Work Plan

Local tool that reads the master orders Excel file (קובץ אב תוצרים) and writes a new
Excel file: **employee (מבצע) x month -> work days**.

## Run
```
pip install -r requirements.txt
python work_plan.py            # asks only for start month and end month
```
Put the master Excel file in the same folder as `work_plan.py` (the newest `.xlsx` is used).
The result is always saved as `work_plan.xlsx` in that folder (overwritten each run); errors go to `work_plan.log`.

Or without prompts:
```
python work_plan.py --input master.xlsx --start 2026-01 --end 2026-12 --output plan.xlsx
```
On Windows you can double-click `run.bat`.

## Rules (phase 1)
- Ignores tabs `הזמנות`, `בתהליך`, `סגורות` and non-order tabs (price lists, summaries).
- Only order tabs whose customer (cell K2/J2) is TEAM NETCOMM, HP/HPE or IBM.
- Only rows whose `תוצר` is `יום עבודה` (x1) or `חצי יום` (x0.5), multiplied by `כמות תוצרים`.
- IBM tabs (customer `IBM`): only `תוצר 068` and `תוצר 069` (0.5 day each), `תוצר 001` (22 days), `תוצר 123` (18 days), multiplied by `כמות תוצרים`.
  Edit `IBM_PRODUCT_DAYS` in `work_plan.py` to change these.
- Month = month of `מועד הפעלת התוצר`. Use `--basis approval` to use `מועד אישור תוצר` instead.
- Rows with no quantity or no valid date are listed in the `שורות שדולגו` sheet.
- Output sheets: `תכנון חודשי` (the table), `פירוט` (every row behind the numbers), `שורות שדולגו`, `הגדרות`.

The source workbook contains prices and personal data, so `*.xlsx` is git-ignored. Do not commit it.

## Working-days config
`working_days.xlsx` (columns: month | working days) is created next to the script on first run from the Israeli work calendar: Sunday-Thursday, minus holidays, with holiday eves (ערב חג) as half a day. Chol HaMoed counts as regular days. Delete the file to regenerate it; edit any number by hand to override. Settings: `WORK_WEEK`, `HOLIDAY_EVE_VALUE`, `EVE_HOLIDAYS` in `work_plan.py`.
In the plan, each employee/month cell is **red** if the days are more than the month's working days, **yellow** if less, **green** if exactly equal. Empty cells and months missing from the config are not coloured.
