# 新興國小 薪資清冊生成器(原檔填值版)— 技能版
# 用法: python3 gen.py --config <設定檔.xlsx> [--out <輸出根目錄>] [--legacy]
#   --legacy: 舊均攤公式引擎(不套節次排除),供迴歸驗證
# 輸入: 設定檔 config-<學年度>-<學期>.xlsx + templates/ 五個原始範本
# 輸出: <out>/<學年度>-<學期>/<月>月/*.xlsx + 驗算報告-<學年度>-<學期>.xlsx
# 驗證: 版型守衛(僅資料格可動)+ fixtures.json(與設定檔同層,若存在)+ 越南語插列結構檢查
import openpyxl, os, re, sys, datetime, calendar, json
from openpyxl.cell.cell import MergedCell
from copy import copy

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.abspath(os.path.join(_SCRIPT_DIR, "..", "templates"))

# ---------- 設定檔 ----------
class Config:
    def __init__(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        self.params = {r[0]: r[1] for r in wb["參數"].iter_rows(min_row=2, values_only=True) if r[0]}
        def tbl(name, cols):
            ws = wb[name]
            head = [c.value for c in ws[1]]
            idx = [head.index(c) for c in cols]
            out = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[idx[0]] is None: continue
                out.append({cols[i]: row[idx[i]] for i in range(len(cols))})
            return out
        self.people   = tbl("人員", ["姓名","職稱","身分別","資格","局號帳號","備註"])
        self.fixed    = tbl("固定節數人員", ["姓名","每週堂數","備註"])
        self.weekly   = tbl("週次制人員", ["姓名","職稱","類別","每週節數","協同日","備註"])
        self.viet     = tbl("越南語人員", ["姓名","上課星期(逗號分隔)","單價","備註"])
        self.excl     = tbl("節次排除", ["姓名","星期","節","班級","科目","支應單位","不採計清冊","生效月份"])
        self.fix_ovr  = tbl("固定節數覆寫", ["姓名","月分","本月總堂數"]) if "固定節數覆寫" in wb.sheetnames else []
        self.super_d  = tbl("超鐘點扣節", ["姓名","月分","本月扣節數(I欄)"]) if "超鐘點扣節" in wb.sheetnames else []
        self.grad    = tbl("畢業後排除", ["類別","姓名","班級","星期","生效起日"]) if "畢業後排除" in wb.sheetnames else []
        self.sections = {}
        if "節次課表" in wb.sheetnames:
            for row in tbl("節次課表", ["姓名","星期","節","班級","科目"]):
                wd = str(row["星期"]).replace("星期","").replace("週","")
                self.sections.setdefault(row["姓名"], []).append(
                    ("一二三四五".index(wd), int(row["節"]) - 1, str(row["班級"]), str(row["科目"])))
        self._mcnt_cache = {}
        self.cal = []
        for row in wb["行事曆"].iter_rows(min_row=2, values_only=True):
            if row[0] is None: continue
            d = row[0]
            d = d.date() if isinstance(d, datetime.datetime) else d
            self.cal.append((d, row[2]))
        self.legacy = False
        self.price = int(self.params["單價"])
        self.sa    = int(self.params["學年度"])          # 民國學年度
        self.sem   = int(self.params["學期"])
        self.months = [int(m) for m in str(self.params["月份"]).split(",")]
        self.made  = datetime.date.fromisoformat(str(self.params["製表日"]))
        self.sem_start = datetime.date.fromisoformat(str(self.params["學期開始"]))
        self.sem_end = datetime.date.fromisoformat(str(self.params["學期結束"]))
        self.dec_cap = int(self.params.get("十二月計費上限") or 31)

    def is_on(self, d):          # 當天是否上課
        for cd, st in self.cal:
            if cd == d: return st == "上課"
        return False

    def mcnt(self, y, m):
        """該月週一~五各至上課日次數(快取)"""
        key = (y, m)
        if key not in self._mcnt_cache:
            c = [0]*5
            for dt in self.calc_days(y, m):   # 走 calc_days:12月關帳上限+1月併入
                c[dt.weekday()] += 1
            self._mcnt_cache[key] = c
        return self._mcnt_cache[key]

    def excluded_slot(self, name, slot, month):
        """節次級排除:slot=(day,period,班級,科目);班級/科目做正規化比對"""
        if self.legacy: return False
        day, per, cls, sub = slot
        for e in self.excl:
            if e["姓名"] != name or "固定節數" not in str(e["不採計清冊"]): continue
            eff = str(e["生效月份"] or "").strip()
            if eff and str(month) not in [x.strip() for x in eff.split(",")]: continue
            wd = str(e["星期"]).replace("星期","").replace("週","")
            if "一二三四五".index(wd) != day: continue
            if int(str(e["節"])) != per + 1: continue
            if norm_class(e["班級"]) != norm_class(cls): continue
            a, b = norm_subject(e["科目"]), norm_subject(sub)
            if not (a == b or a.startswith(b) or b.startswith(a)): continue
            return True
        return False

    def exclusion_slot_check(self, name):
        """設定檔排除列是否都能命中該員實際節次(防打錯字)"""
        secs = self.sections.get(name, [])
        hits = []
        for i, e in enumerate(self.excl):
            if e["姓名"] != name or "固定節數" not in str(e["不採計清冊"]): continue
            ok = any(self.excluded_slot(name, s, 9999) for s in secs)  # 9999=繞過生效月份
            if not ok: hits.append((e["星期"], e["節"], e["班級"], e["科目"]))
        return hits

    def on_days(self, y, m):
        # 該月所有上課日;12月受會計關帳上限:只算到「十二月計費上限」日(含)
        nd = calendar.monthrange(y, m)[1]
        cap = self.dec_cap if m == 12 else 31
        return [datetime.date(y, m, d) for d in range(1, min(nd, cap)+1) if self.is_on(datetime.date(y, m, d))]

    def carryin_days(self, y, m):
        # 12月關帳的後續:該月若為1月,併入上一年12月「上限日次日~12/31」的上課日
        if m == 1 and self.dec_cap < 31:
            start = datetime.date(y - 1, 12, self.dec_cap + 1)
            return sorted(d for d, st in self.cal if start <= d <= datetime.date(y-1, 12, 31) and st == "上課")
        return []

    def calc_days(self, y, m):
        """清冊實際計費日:本月上課日 + (1月併入12月下旬)"""
        return sorted(self.on_days(y, m) + self.carryin_days(y, m))

    def excluded(self, name, register, month):
        """該員在 register 清冊、month 月被排除的節數(節次排除清單)"""
        if self.legacy: return 0
        n = 0
        for e in self.excl:
            if e["姓名"] != name or register not in str(e["不採計清冊"]): continue
            eff = str(e["生效月份"] or "").strip()
            if eff and str(month) not in [x.strip() for x in eff.split(",")]: continue
            n += 1
        return n

    @staticmethod
    def _pdate(v):
        if isinstance(v, datetime.datetime): return v.date()
        if isinstance(v, datetime.date): return v
        return datetime.date.fromisoformat(str(v)[:10])

    def grad_days(self, y, m):
        """畢業後排除(固定節數):{班級: {weekday: 生效日起的上課日數}}"""
        out = {}
        for g in self.grad:
            if str(g["類別"]).strip() != "固定節數": continue
            st = self._pdate(g["生效起日"]); cnt = {}
            for d in self.calc_days(y, m):
                if d >= st: cnt[d.weekday()] = cnt.get(d.weekday(), 0) + 1
            out[norm_class(g["班級"])] = cnt
        return out

    def grad_cut(self, y, m, slots):
        """本月因畢業停課扣掉的堂數(slots: (星期,節,班級,科目))"""
        gd = self.grad_days(y, m); cut = 0
        for s in slots:
            cc = gd.get(norm_class(s[2]))
            if cc: cut += cc.get(s[0], 0)
        return cut

    def grad_stop(self, name, kind):
        """畢業後排除(直播共學/越南語):該員生效起日"""
        for g in self.grad:
            if str(g["類別"]).strip() == kind and str(g["姓名"] or "").strip() == name:
                return self._pdate(g["生效起日"])
        return None

    def super_weeks(self, y, m):
        """超鐘點週數:週一~週日為一週,該週含計費日(含1月併入的12月下旬)即計 1 週"""
        weeks = set()
        for d in self.calc_days(y, m):
            weeks.add(d - datetime.timedelta(days=d.weekday()))
        return sorted(weeks)

    def span(self, y, m):
        """超鐘點/直播標題的授課區間:第一個計費週的週一 ~ 最後一個計費週內最後的上課日"""
        wks = self.super_weeks(y, m)
        start = wks[0]
        lastwk = wks[-1]; lastsun = lastwk + datetime.timedelta(days=6)
        end = max(d for d, st in self.cal if lastwk <= d <= lastsun and st == "上課")
        return start, end

    def live_weeks(self, wd_cn, y, m, stop=None):
        """直播共學(純整月制,115.11.11 修訂):該清冊計費月內,協同星期幾為上課日的次數。
        國定假日/寒假自動扣;12 月關帳由 on_days 收斂、12 月下旬併入 1 月由 calc_days 處理。
        沒有開學月延伸(舊版 9 月曾得陳碧雪4+周憶伶5;校方確認改純整月=3+4)。"""
        cn = ["星期一","星期二","星期三","星期四","星期五","星期六","星期日"]
        k = cn.index(wd_cn)
        return sum(1 for d in self.calc_days(y, m) if d.weekday() == k and (stop is None or d < stop))

    def roc(self, d):  # 民國年
        return d.year - 1911

def xls_round(x):  # Excel ROUND(正值,0) = 四捨五入
    import math
    return int(math.floor(x + 0.5))

def norm_class(s):
    s = str(s).replace(" ", "")
    m = re.match(r"^(\d)0?(\d)$", s)          # 501 / 51 → 5-1班
    if m: return f"{m.group(1)}-{m.group(2)}班"
    m = re.match(r"^(\d)年(\d+)班?$", s)      # 5年1班 → 5-1班
    if m: return f"{m.group(1)}-{m.group(2)}班"
    return s

def norm_subject(s):
    return str(s).replace("合", "")            # 綜合2 → 綜2

# ---------- 通用工具 ----------
def clear_row(ws, r, cmax):
    for c in range(1, cmax+1):
        if not isinstance(ws.cell(r, c), MergedCell): ws.cell(r, c).value = None

def copy_row_style(ws, src, dst, cmax):
    for c in range(1, cmax+1):
        a, b = ws.cell(src, c), ws.cell(dst, c)
        b.font = copy(a.font); b.border = copy(a.border); b.fill = copy(a.fill)
        b.alignment = copy(a.alignment); b.number_format = a.number_format
    if ws.row_dimensions[src].height:
        ws.row_dimensions[dst].height = ws.row_dimensions[src].height

def shift_merged(ws, first_row, k):
    """insert_rows 後把合併範圍下移 k 列(openpyxl 不自動搬)"""
    rngs = [str(r) for r in ws.merged_cells.ranges]
    for r in rngs: ws.unmerge_cells(r)
    for r in rngs:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", r)
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        if r1 >= first_row: r1 += k
        if r2 >= first_row: r2 += k
        ws.merge_cells(f"{c1}{r1}:{c2}{r2}")

def retitle(ws, coord, month_roc, span_s=None, sem_s=None):
    s = ws[coord].value
    if sem_s: s = re.sub(r"\d{3}學年度第\d學期", sem_s, s)
    s = re.sub(r"\d{3}-\d{1,2}月份", f"{month_roc[0]}-{month_roc[1]}月份", s)
    if span_s: s = re.sub(r"\(\d{3}\.\d{2}\.\d{2}-\d{3}\.\d{2}\.\d{2}\)", f"({span_s[0]}-{span_s[1]})", s)
    ws[coord] = s

def fmt_span(d): return f"{d.year-1911}.{d.month:02d}.{d.day:02d}"

# ---------- 各清冊生成 ----------
def gen_fixed(cfg, y, m, tpl):
    wb = openpyxl.load_workbook(tpl); ws = wb.worksheets[0]
    days = len(cfg.calc_days(y, m))
    ovr = {o["姓名"]: int(o["本月總堂數"]) for o in cfg.fix_ovr if int(o["月分"]) == m}
    rep, rows = [], {}
    r = 6
    while ws.cell(r, 2).value and "合" not in str(ws.cell(r, 2).value):
        name = str(ws.cell(r, 2).value).strip()
        person = [p for p in cfg.fixed if p["姓名"] == name]
        assert person, f"固定節數:範本有 {name} 但設定檔沒有"
        secs = cfg.sections.get(name)
        if secs and not cfg.legacy:
            # 節次層精算:逐節次乘該月該星期之上課日數(排除列另扣)
            for bad in cfg.exclusion_slot_check(name):
                print(f"!! 警告:{name} 的排除列 {bad} 未命中任何實際節次(請核對設定檔)")
            kept = [s for s in secs if not cfg.excluded_slot(name, s, m)]
            wk = len(kept)
            cut = cfg.grad_cut(y, m, kept)
            e = ovr.get(name, sum(cfg.mcnt(y, m)[s[0]] for s in kept) - cut)
            if cut: rep.append((name, f"畢業後(6/16起)六年級停課,本月扣 {cut} 堂"))
        else:
            wk = int(person[0]["每週堂數"]) - cfg.excluded(name, "固定節數", m)
            e = ovr.get(name, xls_round(wk / 5 * days))
        assert wk > 0, f"{name} 採計週堂數為 0,請從範本移除該列"
        ws.cell(r, 3).value = cfg.price
        ws.cell(r, 4).value = wk
        ws.cell(r, 5).value = e
        rows[name] = (wk, e, e * cfg.price)
        r += 1
    total = sum(v[2] for v in rows.values())
    cs = cfg.calc_days(y, m)
    ws["A2"] = f"  {cfg.sa}學年度第{cfg.sem}學期{y-1911}年{m}月份固定授課節數兼任教師授課鐘點費({fmt_span(cs[0])}-{fmt_span(cs[-1])})"
    ws["M3"] = f" {y-1911}年{cfg.made.month:02d}月{cfg.made.day:02d}日製表"
    ws.title = f"{y-1911}-{m:02d}固定授課"
    for p_ in cfg.fixed:
        ex = cfg.excluded(p_["姓名"], "固定節數", m)
        if ex: rep.append((p_["姓名"], f"{p_['每週堂數']}−{ex}(輔導處兼輔)=採計{p_['每週堂數']-ex}節/週"))
    return wb, {"total": total, "rows": rows, "days": days, "rep": rep}

def gen_super(cfg, y, m, tpl):
    wb = openpyxl.load_workbook(tpl); ws = wb.worksheets[0]
    wks = cfg.super_weeks(y, m); n_weeks = len(wks)
    if len(wks) == 1 and (wks[0] + datetime.timedelta(days=2)).month > m:
        # 開學月僅一個計費週、且該週週三落在次月 → 依週次歸屬次月,本月不另出冊(避免雙重支付)
        return None, {"total": 0, "rows": {}, "weeks": 0, "removed": [], "span": None}
    start, end = cfg.span(y, m)
    r = 6; num = 0; rows = {}; removed = []
    while ws.cell(r, 3).value and "合" not in str(ws.cell(r, 2).value or ""):
        name = str(ws.cell(r, 3).value).strip()
        person = [p for p in cfg.weekly if p["姓名"] == name and p["類別"] == "超鐘點"]
        assert person, f"超鐘點:範本有 {name} 但設定檔沒有"
        wk_h = int(person[0]["每週節數"]) - cfg.excluded(name, "超鐘點", m)
        if wk_h <= 0:
            removed.append((name, person[0]["職稱"]))
            clear_row(ws, r, 9); ws.row_dimensions[r].hidden = True
            r += 1; continue
        num += 1
        ws.cell(r, 1).value = num
        ws.cell(r, 2).value = person[0]["職稱"]
        ws.cell(r, 4).value = wk_h
        ws.cell(r, 5).value = n_weeks
        ws.cell(r, 7).value = cfg.price
        ded = sum(int(d["本月扣節數(I欄)"]) for d in cfg.super_d if d["姓名"] == name and int(d["月分"]) == m)
        if ded: ws.cell(r, 9).value = ded
        rows[name] = (wk_h, n_weeks, wk_h * n_weeks * cfg.price - ded * cfg.price)
        r += 1
    total = sum(v[2] for v in rows.values())
    retitle(ws, "A1", (y-1911, f"{m:02d}"), (fmt_span(start), fmt_span(end)), f"{cfg.sa}學年度第{cfg.sem}學期")
    ws.title = f"{y-1911}{m:02d}"
    return wb, {"total": total, "rows": rows, "weeks": n_weeks, "removed": removed,
                "span": (fmt_span(start), fmt_span(end))}

def gen_live(cfg, y, m, tpl):
    wb = openpyxl.load_workbook(tpl); ws = wb.worksheets[0]
    r = 6; num = 0; rows = {}; removed_live = []
    while ws.cell(r, 3).value and "合" not in str(ws.cell(r, 2).value or ""):
        name = str(ws.cell(r, 3).value).strip()
        person = [p for p in cfg.weekly if p["姓名"] == name and p["類別"] == "直播共學"]
        assert person, f"直播共學:範本有 {name} 但設定檔沒有"
        wd = person[0]["協同日"]
        wks = cfg.live_weeks(wd, y, m, cfg.grad_stop(name, "直播共學"))
        ex = cfg.excluded(name, "直播共學", m)
        wk_h = int(person[0]["每週節數"]) - ex
        if wks - ex <= 0:
            removed_live.append(name)
            clear_row(ws, r, 9); ws.cell(r, 3).value = name  # 隱藏列保留姓名,金額清空
            ws.row_dimensions[r].hidden = True
            r += 1; continue
        num += 1
        ws.cell(r, 1).value = num
        ws.cell(r, 2).value = person[0]["職稱"]
        ws.cell(r, 4).value = wk_h
        ws.cell(r, 5).value = wks
        ws.cell(r, 6).value = wk_h * wks  # 總堂數:範本此格可能是殘留舊值,一律由程式覆寫
        ws.cell(r, 7).value = cfg.price
        ws.cell(r, 9).value = f"{wd}協同"
        rows[name] = (wd, wks, wk_h * wks * cfg.price)
        r += 1
    total = sum(v[2] for v in rows.values())
    _cs = cfg.calc_days(y, m)
    retitle(ws, "B1", (y-1911, f"{m}"), (fmt_span(_cs[0]), fmt_span(_cs[-1])), f"{cfg.sa}學年度第{cfg.sem}學期")
    ws.title = f"{y-1911}{m:02d}協同"
    return wb, {"total": total, "rows": rows, "removed": removed_live}

def gen_viet(cfg, y, m, tpl):
    wb = openpyxl.load_workbook(tpl); ws = wb.worksheets[0]
    person = cfg.viet[0]
    cn = {"星期一":0,"星期二":1,"星期三":2,"星期四":3,"星期五":4}
    wds = [cn[w.strip()] for w in str(person["上課星期(逗號分隔)"]).split(",")]
    vstop = {}
    for g in cfg.grad:
        if str(g["類別"]).strip() == "越南語" and str(g["姓名"] or "").strip() == str(person["姓名"]).strip():
            w_ = str(g["星期"] or "").strip()
            if w_ in cn: vstop[cn[w_]] = cfg._pdate(g["生效起日"])
    dates = [d for d in cfg.calc_days(y, m) if d.weekday() in wds and d < vstop.get(d.weekday(), datetime.date.max)]
    if not dates:
        return None, {"total": 0, "n": 0, "dates": []}
    price = int(person["單價"])
    TOTAL_ROW = 24            # 範本:資料 7-23(17 列),總計在第 24 列
    MAXR = 23
    for r in range(7, TOTAL_ROW): clear_row(ws, r, 6)
    k = len(dates) - (MAXR - 7 + 1)
    if k > 0:
        pre = [str(r) for r in ws.merged_cells.ranges]
        ws.insert_rows(TOTAL_ROW, amount=k)
        post = [str(r) for r in ws.merged_cells.ranges]
        if post != pre:   # 自動搬移與預期不同→手動重建
            for r in post:
                try: ws.unmerge_cells(r)
                except KeyError: pass
            for r in pre:
                mm = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", r)
                c1, r1, c2, r2 = mm.group(1), int(mm.group(2)), mm.group(3), int(mm.group(4))
                if r1 >= TOTAL_ROW: r1 += k
                if r2 >= TOTAL_ROW: r2 += k
                ws.merge_cells(f"{c1}{r1}:{c2}{r2}")
        for i in range(k): copy_row_style(ws, MAXR, MAXR + 1 + i, 6)
    trow = TOTAL_ROW + max(k, 0)
    for i, d in enumerate(dates):
        r = 7 + i
        wd = "星期一星期二星期三星期四星期五星期六星期日"[d.weekday()*3:d.weekday()*3+3]
        ws.cell(r, 1).value = i + 1
        ws.cell(r, 2).value = f"{d.month}月{d.day}日{wd}"
        ws.cell(r, 3).value = 1
        ws.cell(r, 4).value = price
        ws.cell(r, 5).value = f"=C{r}*D{r}"
    ws.cell(trow, 1).value = "總計"
    ws.cell(trow, 3).value = f"   {len(dates)} 節"
    ws.cell(trow, 5).value = f" {len(dates)*price} 元"
    cs = cfg.calc_days(y, m)
    ws["A2"] = f"上課月分：{y-1911}年 {m} 月（{fmt_span(cs[0])}-{fmt_span(cs[-1])}）"
    ws["D5"] = f"教支老師姓名：{person['姓名']}"
    ws.title = f"表5-{m}月印領清冊(教支老師)"
    return wb, {"total": len(dates) * price, "n": len(dates), "dates": [str(d) for d in dates]}

CN = {8:"八",9:"九",10:"十",11:"十一",12:"十二",1:"一",2:"二",3:"三",4:"四",5:"五",6:"六",7:"七"}
def gen_short(cfg, y, m, tpl):
    wb = openpyxl.load_workbook(tpl); ws = wb.worksheets[0]
    for r in range(5, 18): clear_row(ws, r, 15)   # 清掉範本附帶的示範/前月資料列
    cs = cfg.calc_days(y, m)
    ws["A1"] = f"高雄市新興區新興國民小學{cfg.sa}學年度第{cfg.sem}學期{y-1911}年{CN[m]}月份代課鐘點費清冊({fmt_span(cs[0])}-{fmt_span(cs[-1])})"
    ws["O2"] = f"{y-1911}/{cfg.made.month}/{cfg.made.day}製表"
    return wb, {"total": None, "note": "空白版型(資料由表單紀錄人工/後續自動填入)"}

# ---------- 版型守衛:輸出 vs 範本,只允許資料格不同 ----------
def fmt_diff(out_wb, tpl_path, allowed, maxc=16):
    tpl = openpyxl.load_workbook(tpl_path); tws = tpl.worksheets[0]; ows = out_wb.worksheets[0]
    bad = []
    for r in range(1, max(tws.max_row, ows.max_row)+1):
        for c in range(1, maxc+1):
            a, b = tws.cell(r, c).value, ows.cell(r, c).value
            if a != b and openpyxl.utils.get_column_letter(c)+str(r) not in allowed:
                bad.append(f"{ows.title}!{openpyxl.utils.get_column_letter(c)}{r}: {a!r}→{b!r}")
    return bad

ALLOWED = {
    "fixed":  {"A2","M3"} | {f"{col}{r}" for r in range(6,12) for col in "ACDE"},
    "super":  {"A1"} | {f"{col}{r}" for r in range(6,15) for col in "ABCDEFGHI"},
    "live":   {"B1"} | {f"{col}{r}" for r in range(6,8) for col in "ABDEFGHI"},
    "viet":   {"A2","D5","A24","C24","E24"} | {f"{col}{r}" for r in range(7,24) for col in "ABCDEF"},
    "short":  {"A1","O2"} | {f"{col}{r}" for r in range(5,18) for col in "ABCDEFGHIJKLMNO"},
}

def build_report(cfg, report, outroot):
    from openpyxl import Workbook
    wb = Workbook(); wb.remove(wb.active)
    def tbl(name, header, rows):
        ws = wb.create_sheet(name); ws.append(header)
        for r in rows: ws.append(r)
    tbl("總覽", ["月分","固定節數","超鐘點","直播共學","越南語","短代"],
        [[f"{m}月", s["fixed"]["total"], s["super"]["total"], s["live"]["total"], s["viet"]["total"], "空白版型(短代資料人工/表單匯入)"]
         for m, s in sorted(report["months"].items())])
    tbl("節次排除", ["姓名","星期","節","班級","科目","支應單位","不採計清冊","生效月份"],
        [[e[c] for c in ("姓名","星期","節","班級","科目","支應單位","不採計清冊","生效月份")] for e in cfg.excl])
    for m, s in sorted(report["months"].items()):
        tbl(f"固定節數-{m}月", ["姓名","採計週堂","本月堂數","金額"],
            [[k, v[0], v[1], v[2]] for k, v in s["fixed"]["rows"].items()])
        tbl(f"超鐘點-{m}月", ["姓名","每週節數","週數","金額"],
            [[k, v[0], v[1], v[2]] for k, v in s["super"]["rows"].items()] +
            [["(已移除)"+n, "—", "—", "—"] for n, _ in s["super"]["removed"]])
        tbl(f"直播共學-{m}月", ["姓名","協同日","本月次數","金額"],
            [[k, v[0], v[1], v[2]] for k, v in s["live"]["rows"].items()])
        tbl(f"越南語-{m}月", ["編號","上課日"],
            [[i+1, d] for i, d in enumerate(s["viet"]["dates"])])
    p = os.path.join(outroot, f"{cfg.sa}-{cfg.sem}", f"驗算報告-{cfg.sa}-{cfg.sem}.xlsx")
    wb.save(p); print("report:", p)

# ---------- 主流程 ----------
def parse_args():
    import argparse
    ap = argparse.ArgumentParser(description="新興國小薪資清冊生成器(原檔填值版)")
    ap.add_argument("--config", required=True, help="設定檔 xlsx 路徑")
    ap.add_argument("--out", default=None, help="輸出根目錄(預設:<設定檔上層>/out)")
    ap.add_argument("--legacy", action="store_true", help="舊均攤公式引擎,不套節次排除(迴歸驗證)")
    return ap.parse_args()

def main():
    args = parse_args()
    legacy = args.legacy
    cfg = Config(args.config)
    cfg.legacy = legacy
    outroot = os.path.abspath(args.out) if args.out else \
        os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(args.config)), "..", "out"))
    fxp = os.path.join(os.path.dirname(os.path.abspath(args.config)), "fixtures.json")
    fixtures = json.load(open(fxp, encoding="utf-8")) if os.path.exists(fxp) else None
    tplmap = {"fixed":"固定節數.xlsx","super":"超鐘點.xlsx","live":"直播共學.xlsx","viet":"越南語.xlsx","short":"短代.xlsx"}
    genmap = {"fixed":gen_fixed,"super":gen_super,"live":gen_live,"viet":gen_viet,"short":gen_short}
    report = {"months":{}, "excl":None, "diff":[]}
    ok = True
    for m in cfg.months:
        y = cfg.sem_start.year if m >= 8 else cfg.sem_end.year
        odir = os.path.join(outroot, f"{cfg.sa}-{cfg.sem}" + ("(legacy)" if legacy else ""), f"{m}月")
        os.makedirs(odir, exist_ok=True)
        sums = {}
        for key, tname in tplmap.items():
            tpl = os.path.join(TPL, tname)
            wb, info = genmap[key](cfg, y, m, tpl)
            if wb is None:
                print(f"  [{m}月] {tname}:本月無計費內容,不出冊({info.get('note','0節/0週')})" if info.get("note") else f"  [{m}月] {tname}:本月無計費內容,不出冊")
            else:
                if not legacy:
                    bad = fmt_diff(wb, tpl, ALLOWED[key])
                    if key == "viet" and info.get("n", 0) > 17: bad = []  # 插列月份(12月)版型另以結構檢查
                    if bad: ok = False; report["diff"] += [f"{m}月 {tname}: {b}" for b in bad]
                fname = f"115年{m}月-" + {"fixed":"固定節數清冊","super":"超鐘點清冊","live":"直播共學清冊","viet":"越南語清冊","short":"短代清冊(空白)"}[key] + ".xlsx"
                wb.save(os.path.join(odir, fname))
            sums[key] = info
        if fixtures:
            exp = (fixtures.get("legacy_totals" if legacy else "totals") or {}).get(str(m), {})
            for k, v in exp.items():
                got = sums[k]["total"]
                flag = "OK " if got == v else f"NG(={got})"
                print(f"[fixtures] {m:>2}月 {k:<6} 預期 {v:>7}  {flag}")
                if got != v: ok = False
        pr = (fixtures or {}).get("人別", {}).get(str(m)) if (fixtures and not legacy) else None
        if pr:
            for k, v in pr.items():
                row = sums["fixed"]["rows"].get(k)
                assert row and row[0] == v[0] and row[1] == v[1], f"人別地面真值不符 {m}月 {k}: 期望{v} 得到{row}"
            print(f"[地面真值] {m}月固定節數人別與使用者原始檔一致 ✓")
        report["months"][m] = sums
        print(f"[{m}月] 固定{sums['fixed']['total']} 超鐘點{sums['super']['total']} 直播{sums['live']['total']} 越南語{sums['viet']['total']} 短代=空白")

    # 越南語插列結構檢查(逐月:總計列位置+合併儲存格完整)
    if not legacy:
        wst = openpyxl.load_workbook(os.path.join(TPL, "越南語.xlsx")).worksheets[0]
        for m, s_ in report["months"].items():
            n = s_["viet"]["n"]; trow = 24 + max(0, n - 17)
            fp = os.path.join(outroot, f"{cfg.sa}-{cfg.sem}", f"{m}月", f"115年{m}月-越南語清冊.xlsx")
            if not os.path.exists(fp): continue  # 本月 0 節未出冊
            wsg = openpyxl.load_workbook(fp).worksheets[0]
            assert wsg[f"A{trow}"].value == "總計" and str(wsg[f"C{trow}"].value).strip() == f"{n} 節", f"越南語{m}月總計位置/值錯誤"
            assert len(wsg.merged_cells.ranges) == len(wst.merged_cells.ranges), f"越南語{m}月合併儲存格數量改變"
        print("[結構檢查] 越南語逐月插列版型 OK(總計位置+合併儲存格完整)")
    if report["diff"]:
        print("!! 版型差異(不允許的格):"); [print("   ", d) for d in report["diff"]]
    if not legacy and ok: build_report(cfg, report, outroot)
    print("RESULT:", "ALL PASS" if ok else "FAIL")
    return cfg, report, ok

if __name__ == "__main__":
    main()
