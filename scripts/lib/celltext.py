"""单元格展示文本安全层（v4.3.8）。

背景：第 39→40 期 Excel 盲测发现 14 个单元格被塞入整段公告正文
（最长 1200 字，以「证券代码：/公告编号：/公司及董事会全体成员保证/
重要内容提示」开头）。根因不是采集或解析，而是**证据层与展示层未分离**：

    extract_filings.pdf_facts() 的 out['summary'] 是「前两页 PDF 原文」，
    供 classify_title / 业务口径正则使用（见 from_cninfo 多处 re.search），
    属于内部证据；build_package._row_values() 却直接把它当「交易概要」
    写进单元格。分类规则依赖原文，不能删 summary，必须在写单元格的边界
    做净化，并对成品加硬门禁。

本模块提供三件事：
  is_raw_filing_body()  —— 判定单元格是否为整段公告正文（误报可控）
  filing_brief()        —— 从公告原文抽取有界、无版式噪声的展示摘要
  safe_cell_value()     —— 写单元格前的最后一道净化
  normalize_name()      —— 归并 CJK 之间被 PDF 抽取插入的空格（英 力 特）
"""
from __future__ import annotations

import html
import re

# 公告抬头/免责声明特征。单个特征会误伤现场人工撰写的长摘要
# （例：8 号历史库 I135「本次南京三超…（证券代码：300554…）」），
# 因此判定要求「前 80 字内命中 >= 2 个不同特征」且长度 > 200。
BOILERPLATE_MARKERS = tuple(re.compile(p) for p in (
    r'证券代码\s*[:：]',
    r'证券简称\s*[:：]',
    r'公告编号\s*[:：]',
    r'公司及董事会全体成员保证',
    r'本公司董事会及全体董事',
    r'误导性陈述或者重大遗漏',
    r'没有虚假记载',
    r'重要内容提示',
    r'—\s*1\s*—',
    r'1\s*/\s*8',
))

# 正文里常见的免责句式，抽取摘要时整句丢弃。
DISCLAIMER_RE = re.compile(
    r'虚假记载|误导性陈述|重大遗漏|保证本公告|内容的真实|承担法律责任|'
    r'敬请投资者注意|风险提示|不存在任何虚假')
BULLET_RE = re.compile(r'^[\uf06c\uf0b7\u25cf\u2022\u25a0\u25c6\u25aa\xb7\u00b7\ufffd'
                       r'|\u30fb\u2026l\-—\d\.、\s]+')
HEAD_LINE_RE = re.compile(
    r'—\s*\d+\s*—|^\s*\d+\s*/\s*\d+\s*$|^证券代码|^证券简称|^公告编号|'
    r'^债券代码|^债券简称|公司及董事会全体成员保证|本公司董事会及全体董事|'
    r'误导性陈述或者重大遗漏|重要内容提示|没有虚假记载')
VERB_RE = re.compile(r'拟|签署|转让|收购|增资|认购|投资|购买|出售|持有|出资|设立|受让|变更|合作|协议')
LABEL_RE = re.compile(
    r'^(转让价格及比例|控制权变更|征集期限及方式|交易概况|交易概述|关联交易概述|'
    r'特别提示|重要提示|重要内容提示|一[、.]|二[、.]|三[、.]|本次交易概况)')
COMPANY_RE = re.compile(r'[\u4e00-\u9fff（）()]{2,40}?(?:股份有限公司|有限责任公司|有限公司)')

MAX_BRIEF = 260
# 现场人工稿里存在 702 / 1661 字的 legitimate 长摘要，不能用小上限硬截断。
# 净化层与硬门禁共用同一个 is_raw_filing_body 判据：
#   门禁禁止的 -> 净化层一定移除；门禁允许的 -> 净化层一字不改。
# 这样不会出现「门禁判合法、净化层却悄悄改写」造成的无法解释的差异。
MAX_BRIEF_FALLBACK = 400


def is_raw_filing_body(value) -> bool:
    """单元格是否含整段公告正文（>=2 个抬头特征 + 长度 > 200）。

    现场人工撰写的长摘要里偶尔会引用「证券代码：」，但只命中 1 个特征，
    不会被误判；该判据在第 40 期现场稿上实测 0 命中、生成稿 14 命中。
    """
    if not isinstance(value, str) or len(value) <= 200:
        return False
    head = value[:80]
    return sum(bool(m.search(head)) for m in BOILERPLATE_MARKERS) >= 2


def _strip_boilerplate_head(text: str) -> str:
    """切掉公告抬头区，保留「重要内容提示」之后的实质内容。"""
    t = html.unescape(text or '')
    t = (t.replace('\uf06c', '').replace('\uf0b7', '').replace('\u25cf', '')
          .replace('\ufffd', '').replace('\u25a0', ''))
    m = re.search(r'(?:重要内容提示|重要提示|特别提示)\s*[:：]?', t)
    if m:
        return t[m.end():]
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
    return '\n'.join(ln for ln in lines if not HEAD_LINE_RE.search(ln))


def _flatten(text: str) -> str:
    """PDF 抽取会在 CJK 中间插换行；归并空白但保留西文/数字词间空格。"""
    t = re.sub(r'(?<=[A-Za-z0-9])\s+(?=[A-Za-z0-9])', ' ', text)
    return re.sub(r'\s+', '', t)


def filing_brief(raw: str, limit: int = MAX_BRIEF) -> str:
    """从公告原文抽取有界、无版式噪声的展示摘要。

    不追求复现现场人工撰写的措辞（属内容层，仍未通过盲测），只保证：
      1. 不含证券代码/公告编号/董事会保证等抬头与免责句式；
      2. 长度有界（<= limit），不会把整份公告塞进一个单元格；
      3. 至少含一个交易动作，避免把小标题当摘要。
    """
    if not raw:
        return ''
    t = _flatten(_strip_boilerplate_head(raw))
    t = BULLET_RE.sub('', t)
    # 「转让价格及比例国能英力特…」这类小标题前缀，从首个公司全称处切开
    if LABEL_RE.match(t) or (len(t) > 30 and not t.startswith(('本次', '近日', '公司'))):
        m = COMPANY_RE.search(t[:60])
        if m and m.start() > 0:
            t = t[m.start():]
    sents = re.split(r'(?<=[。！？；])', t)
    out = ''
    for s in sents:
        s = BULLET_RE.sub('', s).strip()
        s = LABEL_RE.sub('', s)
        if len(s) < 25 or DISCLAIMER_RE.search(s) or not VERB_RE.search(s[:70]):
            continue
        if out and len(out) + len(s) > limit:
            break
        out += s
        if len(out) >= 110:
            break
    return out[:limit]


def safe_cell_value(value, fallback: str = ''):
    """写单元格前的最后一道净化：整段公告正文 -> 有界摘要 -> 公告标题。

    判据与 scan_raw_filing_cells / quality_gate 完全一致（is_raw_filing_body）。
    非字符串（金额、日期、比率）和合法文本一律原样返回，不截断、不改写：
    现场人工稿存在 702 / 1661 字的长摘要，硬截断会凭空制造字段差异。
    """
    if not isinstance(value, str) or not is_raw_filing_body(value):
        return value
    brief = filing_brief(value)
    if len(brief) >= 25:
        return brief
    # 兜底用公告标题：已是「巨潮公告｜日期｜」前缀剥离后的官方文件名，
    # 可核对、有界，比留空或塞正文都稳妥。
    return clean_locator(fallback)


def clean_locator(locator: str) -> str:
    """去掉来源定位串的「巨潮公告｜YYYY-MM-DD｜」前缀，只留公告标题。"""
    return re.sub(r'^巨潮公告｜[^｜]+｜', '', str(locator or '')).strip()[:MAX_BRIEF_FALLBACK]


def normalize_name(value: str) -> str:
    """归并 CJK 字符之间被 PDF 抽取插入的空格：英 力 特 -> 英力特。"""
    s = str(value or '').strip()
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r'(?<=[\u4e00-\u9fff])[ \t\u3000]+(?=[\u4e00-\u9fff])', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def scan_raw_filing_cells(paths) -> list[dict]:
    """硬门禁用：扫描成品 xlsx，返回所有「整段公告正文入格」的单元格。"""
    import openpyxl
    hits: list[dict] = []
    for p in paths:
        if p.suffix.lower() not in ('.xlsx', '.xlsm') or p.name.startswith('~$'):
            continue
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if is_raw_filing_body(c.value):
                            hits.append({'file': p.name, 'sheet': ws.title,
                                         'cell': c.coordinate, 'length': len(c.value),
                                         'head': c.value[:60]})
        finally:
            wb.close()
    return hits
