# -*- coding: utf-8 -*-
"""周报 ↔ 并购之友日报 双向核对引擎
用法:
  python3 crosscheck.py --daily mgy_20261112.json --period-dir "<周报目录>" --out report.json
输出四类结论:
  A 日报有·周报缺   → 疑似遗漏（最要紧）
  B 周报有·日报无   → 口径差异或周报独有（需人工判断，不必然是错）
  C 双方都有但字段不一致 → 金额/比例/进展 冲突
  D 双方都有且一致 → 通过
"""
import json, re, argparse, glob, os, datetime
from openpyxl import load_workbook

CODE_RE = re.compile(r"(\d{6})")

def norm_code(v):
    if v is None: return None
    m = CODE_RE.search(str(v))
    return m.group(1) if m else None

def load_workbook_rows(path):
    """把工作簿所有 sheet 读成 (sheet, row_idx, {header: value}) 列表"""
    wb = load_workbook(path, data_only=True, read_only=True)
    out = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        if not rows: continue
        # 找表头行：前 3 行里非空单元格最多且含"代码"或"简称"
        hdr_i, hdr = 0, None
        for i in range(min(3, len(rows))):
            r = rows[i]
            vals = [str(x) for x in r if x is not None]
            if any(("代码" in v) or ("简称" in v) or ("股票" in v) for v in vals):
                hdr_i, hdr = i, r; break
        if hdr is None:
            hdr_i, hdr = 0, rows[0]
        for j in range(hdr_i + 1, len(rows)):
            r = rows[j]
            if not any(x is not None for x in r): continue
            d = {}
            for k, h in enumerate(hdr):
                if h is None: continue
                d[str(h).replace("\n", "").strip()] = r[k] if k < len(r) else None
            out.append({"sheet": sn, "row": j + 1, "data": d})
    wb.close()
    return out

def index_by_code(rows):
    idx = {}
    for r in rows:
        for k, v in r["data"].items():
            if "代码" in k:
                c = norm_code(v)
                if c:
                    idx.setdefault(c, []).append(r)
                    break
    return idx

# 周报里与日报可比的表
SHEET_MAP = [
    ("沪深北交易所首次披露", "首次披露"),
    ("现金重大购买", "现金重大购买"),
    ("现金重大出售", "现金重大出售"),
    ("深交所在审", "在审"), ("上交所在审", "在审"), ("北交所在审", "在审"),
    ("重组委通过", "过会"), ("重组委暂缓审议", "暂缓"), ("重组委否决", "否决"),
    ("证监会注册", "注册"),
    ("撤回终止", "终止"),
]
CTRL_SHEETS = [("2026年首次披露", "控制权变更-首次披露"), ("2026年已完成", "控制权变更-已完成")]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--daily", required=True)
    ap.add_argument("--period-dir", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    daily = json.load(open(a.daily, encoding="utf-8"))
    pdates = _period_dates(a.period_dir)
    items = daily["daily"]["items"]

    # 载入周报
    files = sorted(glob.glob(os.path.join(a.period_dir, "*.xlsx")))
    w0 = [f for f in files if os.path.basename(f).startswith("0、")]
    w7 = [f for f in files if os.path.basename(f).startswith("7、")]
    rows0 = load_workbook_rows(w0[0]) if w0 else []
    rows7 = load_workbook_rows(w7[0]) if w7 else []

    idx0 = index_by_code(rows0)
    idx7 = index_by_code(rows7)

    results = []
    for it in items:
        code = it["code"]
        found = []
        for r in idx0.get(code, []):
            if r["sheet"] in dict(SHEET_MAP) or r["sheet"] in [s for s, _ in CTRL_SHEETS]:
                found.append(("0号", r["sheet"], r["row"]))
        for r in idx7.get(code, []):
            if r["sheet"] in [s for s, _ in CTRL_SHEETS]:
                found.append(("7号", r["sheet"], r["row"]))

        entry = {
            "code": code, "name": it["name"], "title": it["title"],
            "category": it["category"],
            "is_control_change": it["is_control_change"],
            "amounts_wan": it["amounts_wan"],
            "found_in_weekly": found,
            "verdict": None, "note": None,
        }
        entry["major_reorg"] = it.get("major_reorg")
        entry["first_disclose"] = it.get("first_disclose")
        entry["is_acquisition"] = it.get("is_acquisition")
        if not found:
            # 判定优先级（用户 2026-11-13 口径）：
            #   1) 非重大纯现金收购 → E 口径外（用户明确不纳入统计）
            #   2) 首次披露不在本期窗口 → F 跨期（既有项目进展，不算遗漏）
            #   3) 其余 → A 疑似遗漏
            fd = it.get("first_disclose")
            in_window = None
            if fd and pdates and not fd.startswith("????"):
                try:
                    d0 = datetime.date(*map(int, fd.split("-")))
                    in_window = pdates[0] <= d0 <= pdates[1]
                except Exception:
                    in_window = None
            if it.get("is_acquisition") and it.get("major_reorg") is False:
                entry["verdict"] = "E_口径外_非重大现金收购"
                entry["note"] = "公告/日报明示不构成重大资产重组，按口径不纳入 0号现金重大购买；无需补录"
            elif in_window is False:
                entry["verdict"] = "F_跨期_首次披露不在本期"
                entry["note"] = "首次披露日 %s 不在本期窗口 %s~%s，属既有项目进展，不算遗漏" % (fd, pdates[0], pdates[1])
                if it.get("is_acquisition") and it.get("major_reorg") is False:
                    entry["note"] += "；且公告明示不构成重大资产重组"
            else:
                entry["verdict"] = "A_日报有_周报缺"
                entry["note"] = "疑似遗漏：周报 0号/7号 均未检索到该代码，需回公告核实"
        else:
            entry["verdict"] = "D_双方都有"
            entry["note"] = "已命中 %d 处，需逐字段比对金额/比例/进展" % len(found)
        results.append(entry)

    # B 类：周报本期新增（首次披露日期在本期）但日报未提
    period_dates = _period_dates(a.period_dir)
    daily_codes = {i["code"] for i in items if i["code"]}
    weekly_new = []
    for r in rows0:
        if r["sheet"] not in dict(SHEET_MAP): continue
        d = r["data"]
        c = None
        for k, v in d.items():
            if "代码" in k:
                c = norm_code(v); break
        if not c: continue
        # 取首次披露日期
        dt = None
        for k, v in d.items():
            if "首次披露" in k and isinstance(v, (datetime.datetime, datetime.date)):
                dt = v.date() if hasattr(v, "date") else v; break
        if dt and period_dates and period_dates[0] <= dt <= period_dates[1]:
            if c not in daily_codes:
                weekly_new.append({"code": c, "sheet": r["sheet"], "row": r["row"], "date": str(dt),
                                   "name": _pick(d, ["股票简称", "公司名称", "简称"])})
    for r in rows7:
        if r["sheet"] not in [s for s, _ in CTRL_SHEETS]: continue
        d = r["data"]
        c = None
        for k, v in d.items():
            if "代码" in k: c = norm_code(v); break
        if not c: continue
        dt = None
        for k, v in d.items():
            if "提示性公告披露日期" in k and isinstance(v, (datetime.datetime, datetime.date)):
                dt = v.date() if hasattr(v, "date") else v; break
        if dt and period_dates and period_dates[0] <= dt <= period_dates[1] and c not in daily_codes:
            weekly_new.append({"code": c, "sheet": "7号" + r["sheet"], "row": r["row"], "date": str(dt),
                               "name": _pick(d, ["公司简称", "简称"])})

    summary = {
        "daily_title": daily["article"]["title"],
        "daily_publish": daily["article"]["publish_time"],
        "daily_window": (daily["daily"].get("window_start"), daily["daily"].get("window_end")),
        "daily_declared": daily["daily"].get("declared_count"),
        "daily_parsed": daily["daily"]["parsed_count"],
        "count_A_missing": sum(1 for r in results if r["verdict"].startswith("A")),
        "count_D_found": sum(1 for r in results if r["verdict"].startswith("D")),
        "count_E_scope_out": sum(1 for r in results if r["verdict"].startswith("E")),
        "count_F_cross_period": sum(1 for r in results if r["verdict"].startswith("F")),
        "count_B_weekly_only": len(weekly_new),
    }
    out = {"summary": summary, "A_and_D": results, "B_weekly_only": weekly_new}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

    print(json.dumps(summary, ensure_ascii=False, indent=1))
    print("\n--- A 类（日报有·周报缺，疑似遗漏）---")
    for r in results:
        if r["verdict"].startswith("A"):
            print("  %s %s | %s" % (r["code"], r["name"], r["title"][:50]))
    print("\n--- E 类（口径外·非重大纯现金收购，不纳入统计）---")
    for r in results:
        if r["verdict"].startswith("E"):
            print("  %s %s | %s" % (r["code"], r["name"], r["title"][:52]))
    print("\n--- F 类（跨期·首次披露不在本期窗口）---")
    for r in results:
        if r["verdict"].startswith("F"):
            print("  %s %s | %s" % (r["code"], r["name"], r["note"]))
    print("\n--- D 类（双方都有）---")
    for r in results:
        if r["verdict"].startswith("D"):
            print("  %s %s → %s" % (r["code"], r["name"], r["found_in_weekly"]))
    print("\n--- B 类（周报本期新增·日报未提，前 15）---")
    for r in weekly_new[:15]:
        print("  %s %s %s r%s" % (r["code"], r["name"], r["sheet"], r["row"]))

def _pick(d, keys):
    for k in keys:
        for kk, vv in d.items():
            if k in kk and vv: return vv
    return None

def _period_dates(d):
    m = re.search(r"(\d{8})至(\d{8})", d)
    if not m: return None
    f = lambda s: datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    return (f(m.group(1)), f(m.group(2)))

if __name__ == "__main__":
    main()
