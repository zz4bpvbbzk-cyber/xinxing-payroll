# 建立新學期的空白設定檔(之後填人員/課表/排除;行事曆用 make_calendar.py)
# 用法: python3 new_config.py --config config/config-116-2.xlsx --year 116 --sem 2 \
#         --months 2,3,4,5,6 --start 2027-02-22 --end 2027-06-30 [--price 405] [--made YYYY-MM-DD]
import argparse, datetime, os
import openpyxl

ap = argparse.ArgumentParser()
ap.add_argument("--config", required=True)
ap.add_argument("--year", type=int, required=True, help="民國學年度,如 116")
ap.add_argument("--sem", type=int, required=True, help="1=上學期 2=下學期")
ap.add_argument("--months", required=True, help="要生成清冊的月分,如 2,3,4,5,6")
ap.add_argument("--start", required=True, help="學期開始日 YYYY-MM-DD")
ap.add_argument("--end", required=True, help="學期最後上課日 YYYY-MM-DD(休業式當天算上課)")
ap.add_argument("--price", type=int, default=405)
ap.add_argument("--made", default=datetime.date.today().isoformat())
a = ap.parse_args()
os.makedirs(os.path.dirname(os.path.abspath(a.config)), exist_ok=True)

wb = openpyxl.Workbook()
p = wb.active; p.title = "參數"
p.append(["項目", "值", "說明"])
for row in [["學年度", a.year, "清冊標題用"], ["學期", a.sem, "1=上 2=下"],
            ["單價", a.price, "每節鐘點費,改這裡全域生效"],
            ["月份", a.months, "要生成的月分(學期間;1月=次年1月)"],
            ["學期開始", a.start, ""], ["學期結束", a.end, ""], ["製表日", a.made, ""],
            ["十二月計費上限", 20, "會計關帳:12月清冊只算到這天(含),12月下旬併入隔年度清冊"]]:
    p.append(row)
for name, head in [
    ("人員", ["代碼", "姓名", "職稱", "身分別", "資格", "局號帳號", "備註"]),
    ("固定節數人員", ["姓名", "每週堂數", "備註"]),
    ("週次制人員", ["姓名", "職稱", "類別", "每週節數", "協同日", "備註"]),
    ("越南語人員", ["姓名", "上課星期(逗號分隔)", "單價", "備註"]),
    ("節次排除", ["姓名", "星期", "節", "班級", "科目", "支應單位", "不採計清冊", "生效月份"]),
    ("固定節數覆寫", ["姓名", "月分", "本月總堂數"]),
    ("超鐘點扣節", ["姓名", "月分", "本月扣節數(I欄)"]),
    ("行事曆", ["日期", "星期", "上課/放假", "備註"]),
]:
    ws = wb.create_sheet(name); ws.append(head)
wb.save(a.config)
print("saved:", a.config)
print("下一步:填四張人員表 → make_calendar.py 建行事曆 → import_json.py 匯課表 → 節次排除 → gen.py")
