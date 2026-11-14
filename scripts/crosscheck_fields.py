# -*- coding: utf-8 -*-
"""C 类：双方都有时，金额/比例/进展 逐字段比对"""
import json, re, glob, os, datetime, argparse
from openpyxl import load_workbook

def norm(s):
    if s is None: return ""
    return re.sub(r"[\s\u3000,，。；;：:（）()【】\[\]\-—_/]", "", str(s))

def get_row(ws, r):
    d = {}
    hdr = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    for c, h in enumerate(hdr, 1):
        if h: d[str(h).replace("\n", "").strip()] = ws.cell(r, c).value
    return d

def find(ws, code, col_kw="代码"):
    for c in range(1, min(ws.max_column, 8) + 1):
        if col_kw in str(ws.cell(1, c).value or ""):
            for r in range(2, ws.max_row + 1):
                v = ws.cell(r, c).value
                if v and str(code) in str(v):
                    return r
    return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--daily"); ap.add_argument("--period-dir"); ap.add_argument("--out")
    a = ap.parse_args()
    daily = json.load(open(a.daily, encoding="utf-8"))["daily"]["items"]
    files = sorted(glob.glob(os.path.join(a.period_dir, "*.xlsx")))
    f0 = [f for f in files if os.path.basename(f).startswith("0、")][0]
    f7 = [f for f in files if os.path.basename(f).startswith("7、")][0]
    wb0, wb7 = load_workbook(f0, data_only=True), load_workbook(f7, data_only=True)

    out = []
    for it in daily:
        code = it["code"]
        rec = {"code": code, "name": it["name"], "checks": []}
        # 控制权类 → 7号
        if it["is_control_change"] and "2026年首次披露" in wb7.sheetnames:
            ws = wb7["2026年首次披露"]
            r = find(ws, code)
            if r:
                d = get_row(ws, r)
                amt_w = None
                for k, v in d.items():
                    if "交易金额" in k: amt_w = v; break
                amt_d = it["amounts_wan"][0] if it["amounts_wan"] else None
                rec["checks"].append(_cmp_amt("7号交易金额", amt_w, amt_d, it["amounts_wan"]))
                rec["checks"].append(_cmp_progress("7号进展", d.get("进展"), it["progress"]))
                rec["checks"].append(_cmp_pct("7号新实控人股比", d.get("新实控人控制股比"), it["percents"]))
        # 收购类 → 0号现金重大购买/首次披露
        else:
            for sn in ["现金重大购买", "沪深北交易所首次披露"]:
                if sn not in wb0.sheetnames: continue
                ws = wb0[sn]
                r = find(ws, code)
                if r:
                    d = get_row(ws, r)
                    amt_w = None
                    for k, v in d.items():
                        if "交易总价值" in k or "交易金额" in k: amt_w = v; break
                    amt_d = it["amounts_wan"][0] if it["amounts_wan"] else None
                    rec["checks"].append(_cmp_amt("%s交易金额" % sn, amt_w, amt_d, it["amounts_wan"]))
                    break
        rec["has_conflict"] = any(c["status"] == "CONFLICT" for c in rec["checks"])
        rec["has_scope"] = any(c["status"] == "SCOPE" for c in rec["checks"])
        out.append(rec)
    wb0.close(); wb7.close()

    conf = [r for r in out if r["has_conflict"]]
    scope = [r for r in out if r.get("has_scope")]
    print("比对 %d 条：真冲突 %d 条，口径差异 %d 条\n" % (len(out), len(conf), len(scope)))
    for r in out:
        print("【%s %s】" % (r["code"], r["name"]))
        for c in r["checks"]:
            print("   %-16s %-9s 周报=%r 日报=%r" % (c["field"], c["status"], c["weekly"], c["daily"]))
        if not r["checks"]:
            print("   （周报未命中，无字段可比）")
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

def _cmp_amt(field, w, d, all_amts=None):
    """金额比对。日报常把"协议转让+要约"合并披露，周报只记协议转让，
    因此差额若能在日报其他金额里找到组成项，判 SCOPE（口径差异），不判 CONFLICT。"""
    st = "OK"
    if isinstance(w, (int, float)) and d:
        wv, dv = float(w), float(d)
        if abs(wv - dv) / max(abs(dv), 1) > 0.005:
            parts = [a for a in (all_amts or []) if a and abs(a - dv) > 0.005]
            # 周报金额 + 某个日报分项 ≈ 日报合计 → 口径差异
            if any(abs(wv + p - dv) / max(abs(dv), 1) <= 0.01 for p in parts):
                st = "SCOPE"
            elif any(abs(wv - p) / max(abs(p), 1) <= 0.005 for p in parts):
                st = "SCOPE"   # 周报取了其中一个分项
            else:
                st = "CONFLICT"
    elif w is None and d:
        st = "MISSING_WEEKLY"
    return {"field": field, "status": st, "weekly": w, "daily": d}

def _cmp_pct(field, w, pcts):
    """股比比对。周报记"控制权变更完成后"比例，日报正文会同时出现
    当前要约比例与完成后比例，任一命中即判 OK；均不命中再判 CONFLICT。"""
    if w is None or not pcts: return {"field": field, "status": "SKIP", "weekly": w, "daily": pcts}
    wv = float(w) * 100 if float(w) <= 1 else float(w)
    hit = any(abs(wv - p) < 0.02 for p in pcts)
    if hit:
        return {"field": field, "status": "OK", "weekly": round(wv, 4), "daily": pcts[:6]}
    # 差额可由"要约比例+增持比例"合成 → 口径差异
    for i in range(len(pcts)):
        for j in range(len(pcts)):
            if i != j and abs(pcts[i] + pcts[j] - wv) < 0.02:
                return {"field": field, "status": "SCOPE", "weekly": round(wv, 4), "daily": pcts[:6]}
    return {"field": field, "status": "CONFLICT", "weekly": round(wv, 4), "daily": pcts[:6]}

def _cmp_progress(field, w, dtext):
    if not w or not dtext: return {"field": field, "status": "SKIP", "weekly": w, "daily": None}
    kw = ["已完成", "进行中", "终止", "过会", "注册"]
    dw = next((k for k in kw if k in str(w)), None)
    dd = next((k for k in kw if k in str(dtext)), None)
    st = "OK" if dw == dd else ("CONFLICT" if (dw and dd) else "SKIP")
    return {"field": field, "status": st, "weekly": w, "daily": dd}

if __name__ == "__main__":
    main()
