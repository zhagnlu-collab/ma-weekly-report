# -*- coding: utf-8 -*-
"""一期周报 ↔ 全部日报 批量核对总入口"""
import json, subprocess, sys, os, glob, argparse, datetime, re

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period-dir", required=True)
    ap.add_argument("--dailies", default=".")          # mgy_YYYYMMDD.json 所在目录
    ap.add_argument("--out", default="crosscheck_period.json")
    a = ap.parse_args()

    files = sorted(glob.glob(os.path.join(a.dailies, "mgy_2026*.json")))
    files = [f for f in files if re.search(r"mgy_\d{8}\.json$", f)]
    print("载入日报 %d 期\n" % len(files))

    allA, allD, allE, allF = [], [], [], []
    per_day = []
    for f in files:
        day = os.path.basename(f)[4:12]
        tag = os.path.join(a.dailies, "xc_%s.json" % day)
        r = subprocess.run([sys.executable, "crosscheck.py", "--daily", f,
                            "--period-dir", a.period_dir, "--out", tag],
                           capture_output=True, text=True)
        if r.returncode != 0:
            print("  跳过 %s: %s" % (day, r.stderr.strip()[-160:])); continue
        d = json.load(open(tag, encoding="utf-8"))
        s = d["summary"]
        per_day.append({"day": day, "title": s["daily_title"],
                        "window": s["daily_window"], "declared": s["daily_declared"],
                        "parsed": s["daily_parsed"],
                        "A": s["count_A_missing"], "D": s["count_D_found"],
                        "B": s["count_B_weekly_only"]})
        for it in d["A_and_D"]:
            it["day"] = day
            v = it["verdict"]
            if v.startswith("A"):
                allA.append(it)
            elif v.startswith("E"):
                allE.append(it)
            elif v.startswith("F"):
                allF.append(it)
            else:
                allD.append(it)

    # 去重（同一公司可能跨多日出现）
    def _dedup(lst):
        seen = {}
        for it in lst:
            k = it["code"] or it["name"]
            if k not in seen:
                seen[k] = it; seen[k]["days"] = [it["day"]]
            else:
                seen[k]["days"].append(it["day"])
        return seen

    seenA, seenD, seenE, seenF = _dedup(allA), _dedup(allD), _dedup(allE), _dedup(allF)

    out = {"period_dir": a.period_dir,
           "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "per_day": per_day,
           "A_missing": list(seenA.values()),
           "D_found": list(seenD.values()),
           "E_scope_out": list(seenE.values()),
           "F_cross_period": list(seenF.values()),
           "summary": {"days": len(per_day),
                       "A_unique": len(seenA), "D_unique": len(seenD), "E_unique": len(seenE), "F_unique": len(seenF)}}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

    print("=" * 74)
    print("%-10s %-22s %5s %5s %4s %4s" % ("日期", "日报窗口", "声明", "解析", "缺", "有"))
    for p in per_day:
        print("%-10s %-22s %5s %5s %4d %4d" % (
            p["day"], (p["window"][0] or "?")[:10] + "→" + (p["window"][1] or "?")[:10],
            p["declared"], p["parsed"], p["A"], p["D"]))
    print("=" * 74)
    print("\n【A 类】日报有·周报缺（去重后 %d 家）" % len(seenA))
    for it in seenA.values():
        print("  %-7s %-8s %s" % (it["code"], it["name"], it["title"][:46]))
        print("          出现于: %s" % ",".join(it["days"]))
    print("\n【E 类】口径外·不构成重大资产重组的纯现金收购（%d 家，不纳入统计）" % len(seenE))
    for it in seenE.values():
        print("  %-7s %-8s %s" % (it["code"], it["name"], it["title"][:50]))
    print("\n【F 类】跨期·首次披露不在本期窗口（%d 家，不算遗漏）" % len(seenF))
    for it in seenF.values():
        print("  %-7s %-8s %s" % (it["code"], it["name"], it["note"][:66]))
    print("\n【D 类】双方都有（去重后 %d 家）" % len(seenD))
    for it in seenD.values():
        locs = "; ".join("%s/%s r%s" % (x[0], x[1], x[2]) for x in it["found_in_weekly"])
        print("  %-7s %-8s %s" % (it["code"], it["name"], locs[:80]))

if __name__ == "__main__":
    main()
