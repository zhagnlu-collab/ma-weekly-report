# -*- coding: utf-8 -*-
"""并购之友 公众号日报抓取器
用法:
  python3 fetch_mgy.py --url https://mp.weixin.qq.com/s/xxxx          # 单篇
  python3 fetch_mgy.py --account-articles --gh gh_xxxx [--pages N]    # 历史列表(TikHub)
环境变量: TIKHUB_API_KEY（仅历史列表需要）
"""
import re, html, json, os, sys, argparse, datetime, urllib.request, urllib.parse

UA_MOBILE = ("Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
             "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1")
UA_DESKTOP = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36")

def http_get(url, ua=UA_MOBILE, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": ua})
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")

def http_post_json(url, payload, headers=None, timeout=40):
    data = json.dumps(payload).encode()
    h = {"Content-Type": "application/json", "User-Agent": UA_DESKTOP}
    if headers: h.update(headers)
    req = urllib.request.Request(url, data=data, headers=h, method="POST")
    return json.loads(urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore"))

# ---------------- 单篇文章 ----------------
def fetch_article(url):
    h = http_get(url)
    if "该内容已被发布者删除" in h or "参数错误" in h:
        raise RuntimeError("文章不可访问（已删除或参数错误）")
    title_m = re.search(r"var msg_title = '([^']+)'", h) or re.search(r'property="og:title" content="([^"]+)"', h)
    title = html.unescape(title_m.group(1)).strip() if title_m else None
    ct_m = re.search(r'var ct = "(\d+)"', h) or re.search(r'create_time"?\s*[:=]\s*"?(\d+)', h)
    ct = int(ct_m.group(1)) if ct_m else None
    gh_m = re.search(r'user_name\s*=\s*["\'](gh_[0-9a-zA-Z]+)', h)
    gh = gh_m.group(1) if gh_m else None
    nick_m = re.search(r'nickname\s*[:=]\s*["\']([^"\']+)', h)
    nickname = html.unescape(nick_m.group(1)) if nick_m else None
    biz_m = re.search(r'biz\s*[:=]\s*["\']([A-Za-z0-9=+/]{10,})', h)
    biz = biz_m.group(1) if biz_m else None
    body = _extract_body(h)
    return {"url": url, "title": title, "gh": gh, "nickname": nickname, "biz": biz,
            "publish_ts": ct,
            "publish_time": datetime.datetime.fromtimestamp(ct).strftime("%Y-%m-%d %H:%M:%S") if ct else None,
            "text": body}

def _extract_body(h):
    m = re.search(r'id="js_content"[^>]*>(.*?)</div>\s*<script', h, re.S) or re.search(r'id="js_content"[^>]*>(.*)', h, re.S)
    if not m: return ""
    t = m.group(1)
    t = re.sub(r"<script.*?</script>", "", t, flags=re.S)
    t = re.sub(r"<style.*?</style>", "", t, flags=re.S)
    t = re.sub(r"<br\s*/?>", "\n", t)
    t = re.sub(r"</p>|</section>|</h\d>|</li>|</div>|</tr>", "\n", t)
    t = re.sub(r"<td[^>]*>", "\t", t)
    t = re.sub(r"<[^>]+>", "", t)
    t = html.unescape(t)
    return "\n".join(l.strip() for l in t.split("\n") if l.strip())

# ---------------- 结构化解析 ----------------
CAT_RE  = re.compile(r"^([一二三四五六七八九十]+)、\s*(.+?)\s*(?:[（(](\d+)\s*起[）)])?\s*$")
# 两种格式：
#   A) "1．【粤宏远A 000573】55,000.00万元现金收购…"   （带序号）
#   B) "【华新建材 600801】约90.34亿巴西雷亚尔收购…"   （不带序号）
ITEM_RE = re.compile(r"^(?:(\d+)[．.、]\s*)?【\s*([^】]+?)\s*】\s*(.*)$")
CODE_RE = re.compile(r"^([A-Za-z0-9\*]+)\s+(\d{6})$")



def _first_disclose_date(text):
    """从『信息来源』里抽取本次披露日期（YYYY-MM-DD），用于判断是否落在本期窗口。
    取文本中最靠前的『2026年X月X日』；若无年份则记为 ????-MM-DD。"""
    if not text: return None
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    if m:
        return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    m2 = re.search(r"(\d{1,2})月(\d{1,2})日", text)
    if m2:
        return "????-%02d-%02d" % (int(m2.group(1)), int(m2.group(2)))
    return None


def _major_status(body):
    """判定是否构成重大资产重组。返回 True=构成 / False=不构成 / None=正文未提及。
    用户 2026-11-13 明确：不构成重大资产重组的纯现金收购或出售，不纳入周报统计；
    且公告已明确结论的直接采信，不再复算三项指标。
    兼容表述：「不构成」「亦不构成」「不构成……中规定的重大资产重组」
              「不属于」「并非」「不涉及」；以及「构成」「属于」。"""
    if not body: return None
    # 逐个"重大资产重组"出现位置，看其前方 30 字内是否有否定词
    for m in re.finditer(r"重大资产重组", body):
        pre = body[max(0, m.start() - 30):m.start()]
        # 截断到最近的句读，避免跨句误判
        pre = re.split(r"[。；\n]", pre)[-1]
        if re.search(r"(不构成|亦不构成|不属于|亦不属于|并非|不涉及|未构成)", pre):
            return False
    # 明确"构成……重大资产重组"（排除前面是否定词的情形）
    for m in re.finditer(r"(构成|属于)[^。；\n]{0,14}重大资产重组", body):
        pre = body[max(0, m.start() - 6):m.start()]
        if re.search(r"[不未非]", pre): continue
        return True
    return None

def parse_daily(text):
    """把日报正文解析为 {meta, items[]}"""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    meta = {}
    # 首段统计句
    for l in lines[:4]:
        # 支持 "2026年11月11日6:00至11月12日6:00"（第二段省略年份）与全写两种
        # 兼容三种：6:00 / 6时 / 6点，第二段可省略年份
        m = re.search(
            r"(\d{4})年(\d{1,2})月(\d{1,2})日\s*(\d{1,2})(?:[:：](\d{2})|时|点)?\s*至\s*"
            r"(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日\s*(\d{1,2})(?:[:：](\d{2})|时|点)?", l)
        if m:
            y1 = int(m.group(1))
            y2 = int(m.group(6)) if m.group(6) else y1
            h1 = int(m.group(4)); mi1 = int(m.group(5)) if m.group(5) else 0
            h2 = int(m.group(9)); mi2 = int(m.group(10)) if m.group(10) else 0
            meta["window_start"] = "%d-%02d-%02d %02d:%02d" % (y1, int(m.group(2)), int(m.group(3)), h1, mi1)
            meta["window_end"]   = "%d-%02d-%02d %02d:%02d" % (y2, int(m.group(7)), int(m.group(8)), h2, mi2)
        m2 = re.search(r"(?:共有\s*(\d+)\s*家公司披露|梳理出\s*(\d+)\s*起)", l)
        if m2: meta["declared_count"] = int(m2.group(1) or m2.group(2))
        m3 = re.search(r"对外收购取得控制权\s*(\d+)\s*起", l)
        if m3: meta["declared_acq"] = int(m3.group(1))
        m4 = re.search(r"控制权变动\s*(\d+)\s*起", l)
        if m4: meta["declared_ctrl"] = int(m4.group(1))

    items, cur_cat, cur_cat_n = [], None, None
    for l in lines:
        mc = CAT_RE.match(l)
        if mc and not ITEM_RE.match(l):
            cur_cat = mc.group(2).strip()
            cur_cat_n = int(mc.group(3)) if mc.group(3) else None
            continue
        mi = ITEM_RE.match(l)
        if mi:
            head, rest = mi.group(2), mi.group(3)
            cm = CODE_RE.match(head)
            if cm:
                nm, cd = cm.group(1).strip(), cm.group(2)
            else:
                # 【粤宏远A 000573】无空格 / 【上海亚虹 603159】等
                cm2 = re.match(r"^(.+?)\s*(\d{6})$", head.strip())
                if cm2:
                    nm, cd = cm2.group(1).strip(), cm2.group(2)
                else:
                    nm, cd = head.strip(), None
            items.append({
                "category": cur_cat,
                "seq_in_cat": int(mi.group(1)) if mi.group(1) else None,
                "name": nm,
                "code": cd,
                "title": rest.strip(),
                "body_lines": [],
            })
        elif items:
            items[-1]["body_lines"].append(l)

    for it in items:
        body = "\n".join(it["body_lines"])
        it["body"] = body
        # 金额：优先标题，其次"交易作价/总价/转让总价"附近
        it["amounts_wan"] = _amounts(it["title"]) or _amounts(body)
        it["percents"] = _percents(it["title"]) or _percents(body)
        it["info_src"] = _section(body, "信息来源")
        it["plan"]      = _section(body, "交易方案")
        it["progress"]  = _section(body, "最新进展")
        it["comment"]   = _section(body, "并购之友解读")
        it["is_control_change"] = bool(re.search(r"(自身控制权|控制权发生转让|控制权变更)", it["category"] or ""))
        it["is_acquisition"]    = bool(re.search(r"(对外收购|取得标的公司控制权)", it["category"] or ""))
        it["major_reorg"]       = _major_status(it["body"])
        it["first_disclose"]    = _first_disclose_date(it["info_src"] or it["body"])
    meta["items"] = items
    meta["parsed_count"] = len(items)
    meta["by_cat"] = {}
    for it in items:
        meta["by_cat"][it["category"]] = meta["by_cat"].get(it["category"], 0) + 1
    return meta

def _amounts(s):
    return [float(x.replace(",", "")) for x in re.findall(r"([\d,]+\.\d{2})\s*万元", s or "")][:8]

def _percents(s):
    return [float(x) for x in re.findall(r"(\d+\.\d{2})\s*%", s or "")][:10]

SECTION_NAMES = ["信息来源", "交易方案", "最新进展", "并购之友解读"]

def _section(body, name):
    # 形态一：① 信息来源 / ② 交易方案 …
    m = re.search(r"[①②③④⑤⑥⑦⑧⑨⑩]\s*" + re.escape(name) + r"\s*\n?(.*?)(?=[①②③④⑤⑥⑦⑧⑨⑩]\s*\S|\Z)", body, re.S)
    if m: return m.group(1).strip()
    # 形态二：信息来源：xxx（无圈号，靠下一个节名或行尾切分）
    others = "|".join(re.escape(x) for x in SECTION_NAMES if x != name)
    m2 = re.search(re.escape(name) + r"\s*[:：]\s*(.*?)(?=(?:" + others + r")\s*[:：]|\Z)", body, re.S)
    return m2.group(1).strip() if m2 else None

# ---------------- TikHub 历史列表 ----------------
TIKHUB_BASE = "https://api.tikhub.io/api/v1/wechat_mp/v2"

def tikhub_account_articles(gh, pages=3, page_size=20):
    key = os.environ.get("TIKHUB_API_KEY")
    if not key:
        raise RuntimeError("缺少环境变量 TIKHUB_API_KEY")
    out, offset = [], ""
    for _ in range(pages):
        payload = {"username": gh, "raw": False}
        if offset: payload["offset"] = offset
        r = http_post_json(TIKHUB_BASE + "/fetch_account_articles", payload,
                           headers={"Authorization": "Bearer " + key})
        d = r.get("data") or {}
        arts = d.get("articles") or []
        for a in arts:
            ct = a.get("create_time") or a.get("publish_time")
            try: ct = int(ct) if ct else None
            except: ct = None
            out.append({
                "title": a.get("title"),
                "url": a.get("url") or a.get("link") or a.get("source_url"),
                "app_msg_id": a.get("app_msg_id"),
                "digest": a.get("digest"),
                "publish_ts": ct,
                "publish_time": datetime.datetime.fromtimestamp(ct).strftime("%Y-%m-%d %H:%M:%S") if ct else None,
            })
        if d.get("is_end"): break
        offset = d.get("next_offset")
        if not offset: break
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url")
    ap.add_argument("--account-articles", action="store_true")
    ap.add_argument("--gh")
    ap.add_argument("--pages", type=int, default=3)
    ap.add_argument("--out")
    a = ap.parse_args()

    if a.account_articles:
        arts = tikhub_account_articles(a.gh, a.pages)
        res = {"gh": a.gh, "count": len(arts), "articles": arts}
    elif a.url:
        art = fetch_article(a.url)
        meta = parse_daily(art["text"])
        res = {"article": {k: v for k, v in art.items() if k != "text"}, "daily": meta}
        if a.out:
            base = a.out.rsplit(".", 1)[0]
            open(base + "_raw.txt", "w", encoding="utf-8").write(art["text"])
    else:
        ap.error("需要 --url 或 --account-articles")

    s = json.dumps(res, ensure_ascii=False, indent=1)
    if a.out: open(a.out, "w", encoding="utf-8").write(s)
    print(s[:3000])

if __name__ == "__main__":
    main()
