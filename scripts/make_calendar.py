# 建立/替換設定檔「行事曆」:平日全上課,再依放假日清單標記
# 用法: python3 make_calendar.py --config <設定檔.xlsx> --start YYYY-MM-DD --end YYYY-MM-DD \
#         --holidays "YYYY-MM-DD=名稱,..." [--workdays "YYYY-MM-DD=補假上課名稱,..."]
# 放假日來源:人事總處辦公日曆表(要查最新公告)+ 學校行事曆(開學日/休業式/補假上課日)
import argparse, datetime
import openpyxl

ap = argparse.ArgumentParser()
ap.add_argument("--config", required=True)
ap.add_argument("--start", required=True, help="學期第一天(含寒假前的最後一天,建議到休業後一週)")
ap.add_argument("--end", required=True)
ap.add_argument("--holidays", default="", help='逗號分隔 "YYYY-MM-DD=名稱"')
ap.add_argument("--workdays", default="", help='逗號分隔 "YYYY-MM-DD=名稱"(週六補課/補假上班日)')
a = ap.parse_args()
hol = {}
for it in a.holidays.split(","):
    if it.strip():
        d, _, n = it.partition("=")
        hol[datetime.date.fromisoformat(d.strip())] = n.strip() or "放假"
wk = {}
for it in a.workdays.split(","):
    if it.strip():
        d, _, n = it.partition("=")
        wk[datetime.date.fromisoformat(d.strip())] = n.strip() or "補假上課"

d0, d1 = datetime.date.fromisoformat(a.start), datetime.date.fromisoformat(a.end)
wbn = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]
rows = []
d = d0
while d <= d1:
    if d.weekday() < 5 or d in wk:
        rows.append([d, wbn[d.weekday()],
                     "放假" if d in hol else "上課",
                     hol.get(d) or wk.get(d) or ""])
    d += datetime.timedelta(days=1)

wb = openpyxl.load_workbook(a.config)
if "行事曆" in wb.sheetnames: del wb["行事曆"]
ws = wb.create_sheet("行事曆")
ws.append(["日期", "星期", "上課/放假", "備註"])
for r in rows: ws.append(r)
wb.save(a.config)
n_on = sum(1 for r in rows if r[2] == "上課")
print(f"行事曆 {d0} ~ {d1}:共 {len(rows)} 列,上課 {n_on} 天,放假 {len(rows)-n_on} 天(請核對人事總處公告與學校行事曆)")
