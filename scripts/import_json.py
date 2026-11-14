# 匯入排課系統 JSON → 設定檔「節次課表」,並驗證「節次排除」列可命中
# 用法: python3 import_json.py --config <設定檔.xlsx> --json <排課.json> [<個人.json> ...] [--target 姓名,姓名]
import argparse, json, re, sys
import openpyxl

CN = ["星期一", "星期二", "星期三", "星期四", "星期五"]

def norm_class(s):
    s = str(s).replace(" ", "")
    m = re.match(r"^(\d)0?(\d)$", s)          # 501/51 → 5-1班
    if m: return f"{m.group(1)}-{m.group(2)}班"
    m = re.match(r"^(\d)年(\d+)班?$", s)      # 5年1班 → 5-1班
    if m: return f"{m.group(1)}-{m.group(2)}班"
    return s

def norm_subject(s):
    return str(s).replace("合", "")            # 綜合2 → 綜2

ap = argparse.ArgumentParser(description="排課 JSON 匯入設定檔")
ap.add_argument("--config", required=True)
ap.add_argument("--json", nargs="+", required=True, help="排課系統匯出 JSON(可多檔,個人檔會併入)")
ap.add_argument("--target", default="", help="只匯入這些人(逗號分隔);空白=全部教師")
a = ap.parse_args()
targets = [t.strip() for t in a.target.split(",") if t.strip()] if a.target.strip() else None

slotmap, rows = {}, []
for f in a.json:
    d = json.load(open(f, encoding="utf-8"))
    snap = d["schedule"]["snapshot"]
    lessons = snap["preScheduledLessons"] + snap["generatedLessons"]
    for L in sorted(lessons, key=lambda x: (x["teacherName"], x["day"], x["period"])):
        t = L["teacherName"]
        if targets and t not in targets: continue
        day, per, cls, sub = L["day"], L["period"], str(L["className"]), str(L["subjectName"])
        key = (t, day, per)
        if key in slotmap:
            if (slotmap[key][0], slotmap[key][1]) != (cls, sub):
                print(f"⚠ 衝突:{t} {CN[day]}第{per+1}節 {slotmap[key]} vs {(cls, sub)}(保留先者)")
            continue
        slotmap[key] = (cls, sub)
        rows.append([t, CN[day], per + 1, cls, sub])

wb = openpyxl.load_workbook(a.config)
if "節次課表" in wb.sheetnames: del wb["節次課表"]
ws = wb.create_sheet("節次課表")
ws.append(["姓名", "星期", "節", "班級", "科目"])
for r in sorted(rows): ws.append(r)
wb.save(a.config)
import collections
cnt = collections.Counter(r[0] for r in rows)
print(f"節次課表寫入 {len(rows)} 節:", dict(cnt))

# 驗證「節次排除」列是否都能命中實際節次
wb = openpyxl.load_workbook(a.config, data_only=True)
ex = wb["節次排除"]
slots = {}
for r in rows: slots.setdefault(r[0], []).append((CN.index(r[1]), r[2] - 1, r[3], r[4]))
miss = 0
for row in ex.iter_rows(min_row=2, values_only=True):
    if not row or not row[0]: continue
    name, wd, per, cls, sub = row[0], str(row[1]), int(row[2]), row[3], row[4]
    day = "一二三四五".index(wd.replace("星期", "").replace("週", ""))
    hit = any(s[0] == day and s[1] == per - 1 and norm_class(s[2]) == norm_class(str(cls))
              and (norm_subject(s[3]) == norm_subject(str(sub)) or norm_subject(str(sub)).startswith(norm_subject(s[3])) or norm_subject(s[3]).startswith(norm_subject(str(sub))))
              for s in slots.get(name, []))
    if not hit:
        if name not in slots:
            print(f"– 略過:排除列 {name} 的課表不在本次匯入名單,未驗證(--target 或 --json 未含)")
            continue
        miss += 1
        print(f"✗ 未命中:{name} {wd}第{per}節 {cls} {sub}(請核對拼字/班級名/節次,或該員課表已變)")
if miss:
    print(f"RESULT: {miss} 列排除未命中 → 請更新「節次排除」")
    sys.exit(2)
print("RESULT: 節次排除全數命中 ✓")
