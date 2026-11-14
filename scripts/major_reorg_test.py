#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""重大资产重组认定计算器（《上市公司重大资产重组管理办法》第十四条）

用法:
  python3 major_reorg_test.py --json case.json
  python3 major_reorg_test.py --listco-total 164447 --listco-equity 119523 --listco-rev 42256 \
      --target-total 17002.11 --target-net 7570.28 --target-rev 10190.72 --deal 56160

口径（购买股权）:
  资产总额 = max(标的资产总额, 交易金额) / 上市公司合并资产总计
  资产净额 = max(标的资产净额, 交易金额) / 上市公司归母权益
  营业收入 = 标的营业收入              / 上市公司营业收入   ← 不与交易金额取孰高
  任一 ≥ 50% → 构成重大资产重组
注意:
  - 分子用标的【最近一个会计年度经审计】数，不用评估基准日半年数
  - 交易金额用对价全额，不按收购比例折算
  - 分母用上市公司【同期经审计】数；资产净额分母是【归母】权益
"""
import argparse, json, sys

THRESHOLD = 0.50

def assess(listco, target, deal, direction="purchase"):
    """listco/target 均为万元 dict: total / net / rev"""
    rows = []
    if direction == "purchase":
        rows.append(("资产总额", max(target["total"], deal), listco["total"]))
        rows.append(("资产净额", max(target["net"], deal), listco["equity"]))
        rows.append(("营业收入", target["rev"], listco["rev"]))
    else:  # sale：出售资产，分子不取交易金额孰高，按标的账面数
        rows.append(("资产总额", target["total"], listco["total"]))
        rows.append(("资产净额", target["net"], listco["equity"]))
        rows.append(("营业收入", target["rev"], listco["rev"]))

    out, mx = [], 0.0
    for name, num, den in rows:
        r = num / den if den else None
        out.append({"indicator": name, "numerator_wan": round(num, 2),
                    "denominator_wan": round(den, 2),
                    "ratio": round(r, 6) if r is not None else None,
                    "pass": (r is not None and r >= THRESHOLD)})
        if r: mx = max(mx, r)
    return {
        "direction": direction,
        "threshold": THRESHOLD,
        "rows": out,
        "max_ratio": round(mx, 6),
        "is_major": mx >= THRESHOLD,
        "conclusion": "构成重大资产重组，纳入 0号现金重大购买/出售" if mx >= THRESHOLD
                      else "不构成重大资产重组，不纳入统计",
    }

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json")
    ap.add_argument("--direction", default="purchase", choices=["purchase", "sale"])
    ap.add_argument("--listco-total", type=float); ap.add_argument("--listco-equity", type=float)
    ap.add_argument("--listco-rev", type=float)
    ap.add_argument("--target-total", type=float); ap.add_argument("--target-net", type=float)
    ap.add_argument("--target-rev", type=float)
    ap.add_argument("--deal", type=float)
    ap.add_argument("--out")
    a = ap.parse_args()

    if a.json:
        c = json.load(open(a.json, encoding="utf-8"))
        listco, target, deal = c["listco"], c["target"], c["deal"]
        direction = c.get("direction", "purchase")
    else:
        need = ["listco_total", "listco_equity", "listco_rev",
                "target_total", "target_net", "target_rev", "deal"]
        miss = [n for n in need if getattr(a, n) is None]
        if miss:
            ap.error("缺少参数: " + ", ".join("--" + m.replace("_", "-") for m in miss))
        listco = {"total": a.listco_total, "equity": a.listco_equity, "rev": a.listco_rev}
        target = {"total": a.target_total, "net": a.target_net, "rev": a.target_rev}
        deal = a.deal
        direction = a.direction

    r = assess(listco, target, deal, direction)
    print("%-10s %16s %16s %10s %6s" % ("指标", "分子(万元)", "分母(万元)", "占比", "≥50%"))
    print("-" * 66)
    for row in r["rows"]:
        print("%-10s %16.2f %16.2f %9.2f%% %6s" % (
            row["indicator"], row["numerator_wan"], row["denominator_wan"],
            (row["ratio"] or 0) * 100, "是" if row["pass"] else "否"))
    print("-" * 66)
    print("最高占比 %.2f%% → %s" % (r["max_ratio"] * 100, r["conclusion"]))
    if a.out:
        json.dump(r, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main()
