#!/usr/bin/env python3
"""公告 -> 结构化事件台账（Deal / Event / Evidence），并输出待核清单。

核心原则：
- 官方公告清单证明公告存在；公告标题只证明其明确表达的事件，不用标题猜结论；
- 同一公司同一交易的多份同日/连续公告合并为一个 deal_id，不按公告逐份新建交易；
- 首次披露、首次草案、受理、问询、上会、注册、完成、终止是不同里程碑；
- 结论性字段必须回到 PDF 全文，逐页扫描并保留页码/原文摘录；抽不到就留空。
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import logging
import re
import sqlite3
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

BOARDS = {'沪主板', '深主板', '科创板', '创业板', '北交所'}
KINDS = {'ma_audit', 'ma_cash_purchase', 'ma_cash_sale', 'control', 'partial_transfer',
         'second_shareholder', 'ipo', 'strategic_investment', 'professional_cooperation'}
AUDIT_EVENTS = {'accepted', 'inquiry', 'replied', 'suspended', 'resumed',
                'committee_scheduled', 'committee_passed', 'committee_deferred',
                'committee_rejected', 'simplified_passed', 'registration_submitted',
                'registered', 'terminated', 'completed'}
SZSE_STATE_MAP = {
    '已受理': 'accepted', '已问询': 'inquiry', '已回复': 'replied',
    '并购重组委会议通过': 'committee_passed', '并购重组委会议暂缓': 'committee_deferred',
    '并购重组委会议否决': 'committee_rejected', '提交注册': 'registration_submitted',
    '注册生效': 'registered', '注册结果': 'registered', '终止(撤回)': 'terminated',
    '中止': 'suspended', '中止审查': 'suspended',
}
BSE_STATE_MAP = {
    '已受理': 'accepted', '已问询': 'inquiry', '通过': 'committee_passed',
    '未通过': 'committee_rejected', '暂缓审议': 'committee_deferred',
    '提交注册': 'registration_submitted', '注册': 'registered',
    '不予注册': 'committee_rejected', '中止': 'suspended', '终止': 'terminated',
}
SSE_STATE_MAP = {
    '已受理': 'accepted', '已问询': 'inquiry', '已回复': 'replied',
    '重组委会议结果': 'committee_passed', '提交注册': 'registration_submitted',
    '注册结果': 'registered', '注册生效': 'registered', '终止': 'terminated',
    '中止(财报更新)': 'suspended', '中止及财报更新': 'suspended',
    '补充审核已问询': 'inquiry', '补充审核已回复': 'replied',
}

# 交易所导出状态值的括号/空白全半角不统一：上交所 2026-09-30 批量「中止（财报更新）」
# 用全角括号，而映射表写的是半角，曾导致 10 条窗口内状态全部落入待核、不生成事件，
# 进而使在审既有行的审核状态/更新日期不回写、汇总源数据少 1 家。
_FULLWIDTH_TRANS = str.maketrans('（）［］，；：　', '()[],;: ')


def normalize_state(value: str) -> str:
    """仅做形状归一化，不改变语义；用于状态映射查表，不用于展示或证据留痕。"""
    s = clean_text(value or '').translate(_FULLWIDTH_TRANS)
    return re.sub(r'\s+', '', s)


def lookup_state(value: str, table: dict[str, str]) -> str | None:
    """先精确、后形状归一化查表。不做模糊包含匹配，避免「终止」被误判成「中止」。"""
    raw = clean_text(value or '')
    if raw in table:
        return table[raw]
    key = normalize_state(raw)
    for candidate, event in table.items():
        if normalize_state(candidate) == key:
            return event
    return None


AMOUNT_PATTERNS = [
    r'(?:交易作价|交易金额|交易对价|标的资产作价|标的资产交易价格|收购价款)(?:合计|共计|总额)?[^。\n]{0,40}?([\d,]+(?:\.\d+)?)\s*(亿元|万元|元)(?!/股|每股|每张)',
    r'(?:转让价款)(?:合计|共计|总额)[^。\n]{0,24}?(?:为|人民币)?\s*([\d,]+(?:\.\d+)?)\s*(亿元|万元|元)(?!/股|每股|每张)',
]
SUPPORT_PATTERNS = [
    r'(?:募集配套资金|配套融资)(?:总额|金额)?(?:不超过|为|合计)?[^\d。]{0,16}([\d,]+(?:\.\d+)?)\s*(亿元|万元|元)',
    r'(?:配套募集资金)(?:总额|金额)?(?:不超过|为|合计)?[^\d。]{0,16}([\d,]+(?:\.\d+)?)\s*(亿元|万元|元)',
]
ADVISOR_NAMES = (
    '中信证券', '中信建投', '华泰联合', '国泰海通', '国泰君安', '海通证券',
    '中金公司', '招商证券', '广发证券', '申万宏源', '国金证券', '国投证券',
    '浙商证券', '东吴证券', '西部证券', '国元证券', '国信证券', '天风证券',
    '长江证券', '东方证券', '中银国际', '中泰证券', '财通证券', '南京证券',
    '银河证券', '中国银河', '华林证券', '首创证券', '第一创业', '联储证券',
    '民生证券', '国联民生', '兴业证券', '光大证券', '方正证券', '华安证券',
    '华创证券', '国海证券', '山西证券', '东兴证券', '华西证券', '平安证券',
)
TAG_RE = re.compile(r'<[^>]+>')
SPACE_RE = re.compile(r'[ \t\u3000]+')
DATE_RE = re.compile(r'(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日')


def clean_text(value: str) -> str:
    return re.sub(r'\s+', ' ', TAG_RE.sub('', html.unescape(value or ''))).strip()


def board_of(code: str) -> str:
    c = clean_text(code).upper()
    digits = re.sub(r'\D', '', c)[:6]
    if digits.startswith(('688', '689')):
        return '科创板'
    if digits.startswith(('60', '605')):
        return '沪主板'
    if digits.startswith(('300', '301')):
        return '创业板'
    if digits.startswith('00'):
        return '深主板'
    if digits.startswith(('8', '4', '92')):
        return '北交所'
    return '沪主板'


def exchange_code(code: str) -> str:
    c = clean_text(code)
    digits = re.sub(r'\D', '', c)[:6]
    if not digits:
        return c
    if digits.startswith(('60', '68', '90')):
        return f'{digits}.SH'
    if digits.startswith(('8', '4', '92')):
        return f'{digits}.BJ'
    return f'{digits}.SZ'


def _wan(num: str, unit: str) -> float | None:
    try:
        v = float(num.replace(',', ''))
    except ValueError:
        return None
    return v * 10000 if unit == '亿元' else (v if unit == '万元' else v / 10000)


def _sentence(text: str, start: int, end: int, radius: int = 120) -> str:
    left = max(text.rfind('。', 0, start), text.rfind('\n', 0, start))
    right_candidates = [x for x in (text.find('。', end), text.find('\n', end)) if x >= 0]
    right = min(right_candidates) + 1 if right_candidates else min(len(text), end + radius)
    return SPACE_RE.sub(' ', text[max(0, left + 1):right]).strip()[:500]


def _first_match(pages: list[str], patterns: list[str], field: str) -> tuple[tuple[str, str] | None, dict]:
    for page_no, text in enumerate(pages, start=1):
        for pat in patterns:
            m = re.search(pat, text)
            if m:
                return (m.group(1), m.group(2)), {
                    'field': field, 'page': page_no,
                    'excerpt': _sentence(text, m.start(), m.end())}
    return None, {}


def _clean_title(title: str) -> str:
    return re.sub(r'\s+', ' ', TAG_RE.sub('', html.unescape(title or ''))).strip()


def classify_title(title: str) -> list[tuple[str, str, str]]:
    """Return milestone/kind/confidence for only explicit title signals.

    A title can carry more than one milestone (e.g. committee meeting result and
    registration submission). Ambiguous "权益变动/收购报告书/进展" titles yield no
    event and are sent to pending review instead of becoming confirmed facts.
    """
    t = _clean_title(title)
    hits: list[tuple[str, str, str]] = []

    def add(ev: str, kind: str, conf: str = 'HIGH'):
        if (ev, kind, conf) not in hits:
            hits.append((ev, kind, conf))

    # Exclude service/verification/self-inspection documents from first disclosure.
    service = bool(re.search(r'持续督导|督导报告|法律意见|核查意见|专项核查|会计师事务所|律师事务所|评估报告|自查报告|问询函回复|问询函.{0,12}回复|回复.{0,60}问询|关于回复.{0,80}问询|财务顾问报告|董事会关于|相关主体.{0,10}不存在|符合.{0,20}规定|停牌前.{0,20}股东', t))
    ma_context = bool(re.search(r'购买资产|重大资产重组|重大资产购买|重大资产出售|资产置换|吸收合并|重组上市', t))
    if re.search(r'终止.*(?:首次公开发行|发行上市|上市审核|IPO)|撤回.*(?:发行上市|上市申请)|终止审查', t):
        add('terminated', 'ipo')
    if ma_context and not re.search(r'终止筹划', t) and re.search(r'终止.{0,10}(?:重大资产重组|发行股份|重组事项)|重大资产重组.{0,10}终止|撤回.{0,10}(?:重大资产重组|发行股份)', t):
        add('terminated', 'ma_audit')
    if not service and re.search(r'终止筹划股权收购', t):
        add('terminated', 'ma_cash_purchase', 'HIGH')
    if not service and re.search(r'终止.{0,12}协议转让|协议转让.{0,12}终止', t):
        add('terminated', 'partial_transfer')
    issuer_share_swap = ('换股吸收合并' not in t or '关于公司换股吸收合并' in t)
    if not service and ma_context and issuer_share_swap and re.search(r'(?:同意|获).{0,100}注册|注册生效|注册批复|同意注册', t):
        add('registered', 'ma_audit')
    if not service and ma_context and re.search(r'提交注册|报送注册', t):
        add('registration_submitted', 'ma_audit')
    if not service and ma_context and re.search(r'并购重组审核委员会.{0,18}(?:审议|审核)通过|重组委.{0,10}通过|审核通过', t):
        add('committee_passed', 'ma_audit')
    if not service and ma_context and re.search(r'并购重组审核委员会.{0,18}(?:暂缓|延期审议)|重组委.{0,10}暂缓', t):
        add('committee_deferred', 'ma_audit')
    if not service and ma_context and re.search(r'(?:会议安排|拟召开|计划召开|召开.{0,12}审议会议)', t):
        add('committee_scheduled', 'ma_audit')
    if not service and ma_context and re.search(r'(?:获得|获|予以).{0,8}受理|受理通知', t):
        add('accepted', 'ma_audit')
    if not service and ma_context and re.search(r'恢复审核', t):
        add('resumed', 'ma_audit')
    if not service and ma_context and re.search(r'中止审核|中止审查', t):
        add('suspended', 'ma_audit')
    if not service and ma_context and re.search(r'(?:审核问询函|问询函)(?!.*回复)', t):
        add('inquiry', 'ma_audit')
    if not service and ma_context and re.search(r'(?:问询函|问询).{0,8}回复|回复.{0,8}(?:问询函|问询)', t):
        add('replied', 'ma_audit')
    if re.search(r'标的资产.{0,8}过户完成|过户完成.{0,12}标的资产', t):
        add('completed', 'ma_audit')

    # Official first-disclosure statistics start with a plan (“预案”) or the
    # first full draft (“报告书（草案）”).  A bare suspension/planning notice is
    # retained as a pipeline lead, never promoted into the weekly count.
    draft_excluded = re.search(r'修订|更新|注册稿|上会稿|摘要|说明|法律意见|独立财务顾问|核查意见|评估报告|会计师事务所|律师事务所', t)
    if not service and not draft_excluded and not re.search(r'进展|修订|更新|停牌前|风险提示|投资者说明会|交易对方.{0,8}不存在|相关主体.{0,8}不存在', t):
        is_draft = bool(re.search(r'报告书\s*[（(]草案[）)]', t))
        is_plan = bool(re.search(r'预案', t))
        if is_draft or is_plan:
            milestone = 'first_draft' if is_draft else 'first_disclosed'
            if re.search(r'(?:发行股份(?:及支付现金)?购买资产|发行股份、可转换公司债券及支付现金购买资产)', t):
                add(milestone, 'ma_audit', 'HIGH')
            else:
                if re.search(r'重大资产.{0,20}购买', t):
                    add(milestone, 'ma_cash_purchase', 'HIGH')
                if re.search(r'重大资产.{0,20}出售', t):
                    add(milestone, 'ma_cash_sale', 'HIGH')
        elif not re.search(r'终止筹划', t) and re.search(r'筹划.{0,18}(?:发行股份|重大资产重组)', t):
            add('lead_signal', 'ma_audit', 'LOW')

    control_completion_signal = bool(
        re.search(r'完成过户|完成.{0,8}过户|过户完成|发行完毕|完成工商变更登记|控制权变更完成|实际控制人变更完成|实际控制人变更暨.{0,30}换发营业执照', t)
        and re.search(r'控股股东|实际控制人|控制权', t)
        and re.search(r'变更|变化', t))
    business_reg_change_only = bool('工商变更' in t and not re.search(r'(?:实际控制人|控制权)(?:发生|拟发生)?变更', t))
    if not service and control_completion_signal and not re.search(r'变更名称|经营范围|注册地址', t) and not business_reg_change_only:
        add('control_completed', 'control', 'HIGH')
    if not service and re.search(r'股东协议转让.{0,20}过户完成', t):
        add('control_completed', 'control', 'HIGH')

    substantive_control_first = bool(
        (re.search(r'签署.{0,40}(?:股份转让协议|收购协议|表决权委托协议)', t)
         and (not re.search(r'进展', t) or re.search(r'提示性', t)))
        or re.search(r'一致行动.{0,30}(?:到期终止|到期不再续签|解除)', t))
    control_progress_signal = bool(
        re.search(r'进展', t)
        and re.search(r'控股股东|实际控制人|控制权', t)
        and re.search(r'变更|变化', t)
        and not substantive_control_first
        and not re.search(r'筹划|停牌', t))
    if not service and control_progress_signal and not control_completion_signal:
        add('control_progress', 'control', 'MEDIUM')

    # Explicit control change. A bare equity-change report or takeover report
    # is only a candidate until its body says control changes.
    control_not_changed = bool(re.search(r'实际控制人.{0,10}未(?:发生)?变更|未导致.{0,20}实际控制人.{0,8}变更', t))
    control_stop = bool(re.search(r'意向协议|承诺|进展|完成|终止|继续停牌|筹划|停牌|可能', t))
    if not service and re.search(r'控股股东.{0,30}协议转让.{0,30}公开征集受让方', t):
        add('first_disclosed', 'control', 'HIGH')
    if not service and not control_completion_signal and not control_not_changed and (substantive_control_first or not control_stop):
        if re.search(r'实际控制人.{0,8}(?:发生|拟发生)?变更|控制权.{0,8}(?:发生|拟发生)?变更|控股股东.{0,12}(?:发生|拟发生)?变更', t):
            add('first_disclosed', 'control', 'HIGH')
        elif re.search(r'(?:要约收购报告书|要约收购报告书摘要)$', t):
            add('first_disclosed', 'control', 'MEDIUM')
    # A transfer announcement keeps its own classification.
    if not service and not re.search(r'进展|完成|终止|补充协议|公开征集受让方', t):
        if (not re.search(r'控制权.{0,12}(?:发生|拟发生)?变更', t)
                and re.search(r'控股股东|实际控制人|控制权|一致行动人', t)
                and re.search(r'协议转让.{0,12}(?:部分)?股份|股份转让协议', t)):
            add('first_disclosed', 'partial_transfer', 'HIGH')
    if not service and re.search(r'成为.{0,8}第二大股东', t):
        add('first_disclosed', 'second_shareholder', 'HIGH')
    if not service and not re.search(r'进展|完成|终止', t):
        if re.search(r'筹划重大资产出售|重大资产出售.{0,7}提示性公告', t):
            add('lead_signal', 'ma_cash_sale', 'LOW')
        if re.search(r'筹划.{0,12}现金.{0,8}(?:收购|购买)|现金方式.{0,12}(?:收购|购买).{0,24}重大资产重组', t) and not re.search(r'发行股份|可转换公司债券', t):
            add('lead_signal', 'ma_cash_purchase', 'LOW')
        strategic_signal = (
            re.search(r'引入战略投资者|引入外部投资者|引入投资者', t)
            or re.search(r'老股转让.{0,12}战略投资者', t)
            or re.search(r'投资建设.{0,40}签署《增资协议》', t)
            or re.search(r'全资子公司向其参股公司增资', t)
        )
        strategic_excluded = re.search(r'向全资子公司增资|向全资孙公司增资并完成工商变更|募集资金.*增资|变更注册资本', t)
        parent_level_strategic = bool(re.search(r'引入战略投资者', t) and not re.search(r'子公司|孙公司|参股公司', t))
        if strategic_signal and not strategic_excluded and not parent_level_strategic:
            add('first_disclosed', 'strategic_investment')
        professional_context = re.search(r'专业投资机构|专业机构|产业(?:投资)?基金|股权投资基金|并购基金|母基金|投资基金|私募(?:股权)?基金|创业投资|基金份额|合伙协议', t)
        professional_action = re.search(r'共同投资|合作投资|投资设立|参与设立|参与投资|认购|受让.*份额|发起设立', t)
        professional_excluded = re.search(r'清算注销|延期|放弃.*优先|合作备忘录|进展|完成|终止|募集完毕|募集完成|转让.{0,30}(?:基金|合伙).{0,20}份额', t)
        if professional_context and professional_action and not professional_excluded:
            add('first_disclosed', 'professional_cooperation')
    return hits


def classify(title: str) -> tuple[str, str] | None:
    hits = classify_title(title)
    return (hits[0][0], hits[0][1]) if hits else None


def _utf8_safe_text(value: str) -> str:
    return (value or '').encode('utf-8', 'replace').decode('utf-8')


def _read_pdf_pages(path: Path, cache_dir: Path) -> tuple[list[str], str]:
    cache_dir.mkdir(parents=True, exist_ok=True)
    key = hashlib.sha256(str(path.resolve()).encode() + b'|' + hashlib.sha256(path.read_bytes()).digest()).hexdigest()
    cache = cache_dir / f'{key}.json'
    if cache.is_file():
        try:
            data = json.loads(cache.read_text(encoding='utf-8'))
            return data.get('pages', []), data.get('error', '')
        except (ValueError, OSError):
            cache.unlink(missing_ok=True)
    pages: list[str] = []
    err = ''
    logging.getLogger('pypdf').setLevel(logging.ERROR)
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path), strict=False)
        for page in reader.pages:
            try:
                pages.append(_utf8_safe_text(page.extract_text() or ''))
            except Exception as ex:  # noqa: BLE001
                pages.append('')
                err = f'{type(ex).__name__}: {ex}'
    except Exception as ex:  # noqa: BLE001
        err = f'pypdf: {type(ex).__name__}: {ex}'
    if not any(len(x.strip()) >= 20 for x in pages):
        try:
            import pdfplumber
            pages = []
            with pdfplumber.open(str(path)) as doc:
                for page in doc.pages:
                    pages.append(_utf8_safe_text(page.extract_text() or ''))
            err = ''
        except Exception as ex:  # noqa: BLE001
            err = (err + '; ' if err else '') + f'pdfplumber: {type(ex).__name__}: {ex}'
    tmp = cache.with_suffix(cache.suffix + '.tmp')
    tmp.write_text(json.dumps({'path': str(path.resolve()), 'pages': pages, 'error': err},
                              ensure_ascii=False), encoding='utf-8')
    tmp.replace(cache)
    return pages, err


def _extract_target(text: str) -> str:
    patterns = [
        r'(?:购买|收购)[^。\n]{0,260}?持有的?([^，。；\n]{2,120}(?:股权|股份))',
        r'置入资产为([^。\n]{2,240}?(?:股权|资产))(?=。|；)',
        r'置入资产为([^。\n]{2,160}?股权)',
        r'(?:本次)?标的资产为([^。\n]{2,160}?股权)',
        r'(?:拟|计划|将通过.{0,20})以[^。\n]{0,60}方式(?:购买|收购|取得)[^。\n]{0,80}?([^，。；\n]{2,80}(?:股权|股份|资产))',
        r'标的资产/拟购买资产\s+([^。\n]{2,120})',
        r'标的资产\s*(?:[：:]|指)\s*([^。\n]{2,120})',
        r'(?:标的公司|收购标的)\s*(?:[：:]|指)\s*([^。\n]{2,120})',
        r'购买[^。\n]{0,40}持有的([^。\n]{2,100}(?:股权|股份))',
        r'(?:重大资产出售|重大资产购买)[^。\n]{0,40}?([^，。；\n]{2,100}(?:股权|资产))',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            value = re.sub(r'^(?:其|公司|上市公司|本次)', '', m.group(1)).strip(' ：:，,')
            if 2 <= len(value) <= 100:
                return value
    return ''


def _extract_payment(text: str) -> str:
    if re.search(r'权益变动性质[：:][^。\n]{0,40}协议转让|协议转让方式|股份转让协议', text):
        return '协议转让'
    if re.search(r'发行股份及支付现金', text):
        return '发行股份及支付现金'
    if re.search(r'发行股份、?可转换公司债券(?:及|和)支付现金', text):
        return '发行股份、可转换公司债券及支付现金'
    if re.search(r'发行股份(?:的方式)?购买|发行股份购买资产', text):
        return '发行股份'
    if re.search(r'可转换公司债券(?:及|和)支付现金', text):
        return '可转换公司债券及支付现金'
    if re.search(r'支付现金方式|以现金方式', text):
        return '现金'
    if re.search(r'资产置换', text):
        return '资产置换'
    return ''


def canonical_advisor(value: str) -> str:
    value = clean_text(value)
    if not value:
        return ''
    for name in sorted(ADVISOR_NAMES, key=len, reverse=True):
        if name in value:
            return '银河证券' if name == '中国银河' else name
    value = re.sub(r'(?:股份有限公司|有限责任公司|有限公司)$', '', value)
    return value


def _extract_advisors(text: str) -> list[str]:
    compact = ''.join(text.split())
    windows = []
    for keyword in ('独立财务顾问', '财务顾问'):
        start = 0
        while True:
            pos = compact.find(keyword, start)
            if pos < 0:
                break
            windows.append(compact[max(0, pos - 120):pos + 220])
            start = pos + len(keyword)
    found = []
    for window in windows:
        for name in ADVISOR_NAMES:
            if name in window:
                canonical = '银河证券' if name == '中国银河' else name
                if canonical not in found:
                    found.append(canonical)
    return found[:4]


def _extract_control_facts(text: str) -> dict:
    out: dict = {}
    # PyPDF frequently breaks Chinese phrases across lines (e.g. “实际\n控制人”).
    # Canonicalise whitespace for sentence-level assertions; keep the original
    # page text for evidence excerpts.
    flat = ''.join(text.split())
    alias_map = {alias: full for full, alias in re.findall(
        r'([^，。；]{2,80}?(?:委员会|公司|集团))[（(]以下简称[“\"]?(.{2,30}?)[”\"]?[）)]', flat)}
    current_no_change = re.search(
        r'本次(?:收购|权益变动|交易)[^。]{0,30}完成后[^。]{0,120}(?:仍为[^。]{0,30}实际控制人|不会导致[^。]{0,30}控制权[^。]{0,12}(?:发生变化|变更)|仍无实际控制人)',
        flat)
    m = re.search(r'实际控制人[^。]{0,80}?(?:由|从)([^。]{2,50}?)(?:变更为|变为|调整为|转由)([^。]{2,50})', flat)
    if not m:
        m = re.search(r'([^。]{2,50}?)不再(?:为公司)?实际控制人[^。]{0,60}?实际控制人(?:变更|调整)为([^。]{2,50})', flat)
    def tidy(value: str) -> str:
        value = value.strip(' ，,；;：“”')
        value = re.split(r'[，,；;。]', value)[-1].strip(' ，,；;：“”')
        value = re.split(r'（以下简称|\(以下简称', value, maxsplit=1)[0]
        value = re.sub(r'(?:先生|女士)$', '', value).strip(' ，,；;：“”')
        return alias_map.get(value, value)
    if m:
        out['old_controller'] = tidy(m.group(1))
        out['new_controller'] = tidy(m.group(2))
    else:
        old = re.search(r'本次(?:发行|收购|权益变动|交易)前[^。]{0,140}?实际控制人为([^，。；]+)', flat)
        if not old:
            old = re.search(r'(?:上市公司|公司)(?:原|现有)?控股股东、实际控制人为([^，。；]{2,80})', flat)
        new = re.search(r'本次(?:发行|收购|权益变动|交易)完成后[^。]{0,140}?([\u4e00-\u9fa5A-Za-z·]{2,20})(?:先生|女士)?(?:已经|已)?(?<!将)(?<!拟)成为(?:上市公司|公司)?实际控制人', flat)
        if not new:
            new = re.search(r'([\u4e00-\u9fa5A-Za-z·]{2,20})(?:先生|女士)?(?:已经|已)?(?<!将)(?<!拟)成为(?:上市公司|公司)?实际控制人', flat)
        if not new:
            new = re.search(r'实际控制人(?:变更|调整)为([^，。；]+)', flat)
        if old:
            out['old_controller'] = tidy(old.group(1))
        if new:
            out['new_controller'] = tidy(new.group(1))
    negative_patterns = [
        r'(?:上市公司|公司|标的公司)?[^。]{0,20}仍无实际控制人',
        r'(?:本次[^。]{0,24})?(?:不会|不|未)导致[^。]{0,32}(?:控股股东[、及和]?)?实际控制人[^。]{0,16}(?:变更|变化)',
        r'(?:上市公司|公司)?实际控制人[^。]{0,30}(?:仍为|未发生变化|保持不变)',
        r'(?:不会|未)导致(?:上市公司)?控制权[^。]{0,12}(?:发生变化|变更)',
        r'本次[^。]{0,30}不涉及[^。]{0,20}实际控制人[^。]{0,12}(?:变更|变化)',
    ]
    negative = next((m for p in negative_patterns if (m := re.search(p, flat))), None)
    positive_actual = bool(out.get('new_controller')) or bool(re.search(
        r'实际控制人[^。]{0,30}(?:将|已)?(?:发生)?(?:变更|调整)为', flat))
    planned_control = bool(re.search(r'控制权拟发生变更|将取得.{0,12}控制权|成为.{0,12}控股股东|(?:将|拟)成为.{0,16}实际控制人', flat))
    # An explicit transition of the listed company's actual controller outranks
    # an earlier generic statement that the transaction itself does not change
    # control (often found in transaction history sections).
    strong_transition = bool(re.search(
        r'不再(?:为公司)?实际控制人|实际控制人(?:将|已)?(?:由|从).{0,80}?(?:变更为|调整为)|实际控制人(?:将|已)?(?:变更|调整)为|(?:控股股东|实际控制人)自[^。]{0,30}起由[^。]{0,80}变更为',
        flat))
    if current_no_change or (negative and not strong_transition):
        out['control_result'] = '未导致实际控制人变化'
        out.pop('old_controller', None)
        out.pop('new_controller', None)
    elif positive_actual:
        out['control_result'] = '实际控制人发生变化'
    elif planned_control:
        out['control_result'] = '控制权拟发生变更'
    if out.get('control_result') == '实际控制人发生变化' and re.search(
            r'不再(?:为公司)?实际控制人|已经完成|已完成(?:本次)?(?:变更|过户)|实际控制人(?:已经|已)变更为|发行完毕|(?:股份转让|标的股份).{0,30}已完成过户登记|完成了?[^。]{0,30}工商变更登记', flat):
        out['completion_confirmed'] = True
    completion = re.search(
        r'(?:已于)?(20\d{2})年(\d{1,2})月(\d{1,2})日[^。]{0,45}(?:完成|办理)[^。]{0,20}过户', flat)
    if not completion:
        completion = re.search(
            r'(?:过户日期|过户完成日期)为?(20\d{2})年(\d{1,2})月(\d{1,2})日', flat)
    if not completion:
        completion = re.search(
            r'(20\d{2})年(\d{1,2})月(\d{1,2})日[^。]{0,45}过户(?:登记)?(?:手续)?(?:已)?(?:办理)?(?:完成|完毕)', flat)
    if not completion:
        completion = re.search(
            r'(20\d{2})年(\d{1,2})月(\d{1,2})日[^。]{0,45}完成了?[^。]{0,30}工商变更登记', flat)
    if completion:
        try:
            out['completion_date'] = date(int(completion.group(1)), int(completion.group(2)), int(completion.group(3))).isoformat()
        except ValueError:
            pass
    before_ratio = re.search(r'本次(?:发行|收购|权益变动|交易)前[^。]{0,320}?(?:占比|持股比例|控制)[^\d。]{0,20}(\d{1,3}(?:\.\d+)?)%', flat)
    after_ratio = re.search(r'本次(?:发行|收购|权益变动|交易)完成后[^。]{0,320}?(?:占比|持股比例|控制)[^\d。]{0,20}(\d{1,3}(?:\.\d+)?)%', flat)
    if before_ratio:
        out['controller_ratio_before_pct'] = before_ratio.group(1)
    if after_ratio:
        out['controller_ratio_after_pct'] = after_ratio.group(1)
    return out


def pdf_facts(path: Path, cache_dir: Path | None = None, limit_pages: int | None = None) -> dict:
    """Extract facts from every PDF page; return empty fields when unsupported."""
    cache_dir = cache_dir or path.parent.parent / 'pdf-cache'
    out: dict = {
        'text_pages_scanned': 0, 'text_chars': 0, 'amount_cny_wan': None,
        'supporting_financing_cny_wan': None, 'ratio_pct': None,
        'purchase_amount_cny_wan': None, 'sale_amount_cny_wan': None,
        'purchase_target': '', 'sale_target': '', 'termination_stage': '',
        'external_investor': False, 'strategic_non_third_party': False,
        'target': '', 'payment_method': '', 'advisors': [],
        'control_result': '', 'old_controller': '', 'new_controller': '', 'meeting_date': '',
        'completion_confirmed': False, 'completion_date': '', 'meeting_id': '', 'summary': '',
        'generic_target': '', 'control_transaction_amount_cny_wan': None,
        'fact_pages': {}, 'excerpt': '', 'error': '',
    }
    pages, err = _read_pdf_pages(path, cache_dir)
    if limit_pages:
        pages = pages[:limit_pages]
    out['text_pages_scanned'] = len(pages)
    out['text_chars'] = sum(len(x) for x in pages)
    out['error'] = err
    text_pages = [(i, p) for i, p in enumerate(pages, start=1) if p.strip()]
    full = '\n'.join(p for _, p in text_pages)
    if not full.strip():
        return out
    primary = '\n'.join(p for page_no, p in text_pages if page_no <= 40)
    # A combined transaction often presents an assessment total first and the
    # transaction price last in a table row (e.g. 106,000 ... 93,316.49).
    for page_no, page_text in ((n, p) for n, p in text_pages if n <= 40):
        m = re.search(r'合计\s+[\d,.]+\s+-\s+-\s+([\d,.]+)\s+-', page_text)
        if m:
            try:
                out['amount_cny_wan'] = float(m.group(1).replace(',', ''))
            except ValueError:
                pass
            if out['amount_cny_wan']:
                out['fact_pages']['amount_cny_wan'] = {
                    'page': page_no, 'excerpt': _sentence(page_text, m.start(), m.end())}
                out['purchase_amount_cny_wan'] = out['amount_cny_wan']
                out['excerpt'] = out['excerpt'] or out['fact_pages']['amount_cny_wan']['excerpt']
                break
    for unit_field, patterns, value_field in (
        ('transaction_amount', AMOUNT_PATTERNS, 'amount_cny_wan'),
        ('supporting_financing', SUPPORT_PATTERNS, 'supporting_financing_cny_wan'),
    ):
        if unit_field == 'transaction_amount' and out.get(value_field):
            continue
        scan_pages = [(n, p) for n, p in text_pages if n <= 40] if unit_field == 'supporting_financing' else text_pages
        if unit_field == 'transaction_amount':
            scan_pages = [(n, p) for n, p in text_pages if n <= 40] + [(n, p) for n, p in text_pages if n > 40]
        for page_no, page_text in scan_pages:
            hit = None
            for pat in patterns:
                m = re.search(pat, page_text)
                if m:
                    hit = m
                    break
            if hit:
                value = _wan(hit.group(1), hit.group(2))
                evidence = _sentence(page_text, hit.start(), hit.end())
                if value and value > 0:
                    # Split consideration (“股份对价 3.00亿元、现金对价 1.00亿元”)
                    # is not a disclosed total and must not become transaction_amount.
                    if unit_field == 'transaction_amount' and '暂估' in evidence and ('尚未确定' in evidence or '现金方式' in evidence):
                        continue
                    out[value_field] = value
                    out['fact_pages'][value_field] = {'page': page_no, 'excerpt': evidence}
                    if unit_field == 'transaction_amount':
                        out['purchase_amount_cny_wan'] = value
                    if not out['excerpt']:
                        out['excerpt'] = evidence
                    break
    compact_full = ''.join(full.split())
    issue_amount = re.search(r'募集资金总额为[人民币￥]?([\d,]+(?:\.\d+)?)元', ''.join(full.split()))
    if issue_amount:
        try:
            out['control_transaction_amount_cny_wan'] = float(issue_amount.group(1).replace(',', '')) / 10000
        except ValueError:
            pass
    m = re.search(r'(?:本次协议转让|本次股份转让|转让的?股份|交易标的)[^。\n]{0,36}?(?:占|比例为|合计)[^\d。\n]{0,12}([\d.]+)\s*%', full)
    if not m:
        m = re.search(r'占(?:上市公司|公司)?(?:总股本|已发行股份总数|股份总数)[^\d。\n]{0,12}([\d.]+)\s*%', compact_full)
    if m:
        try:
            out['ratio_pct'] = float(m.group(1))
        except ValueError:
            pass
    first_company = re.search(
        r'(?m)^\s*([^\n：:]{2,50}?(?:股份有限公司|有限公司))\s*[（(]以下简称',
        primary or full)
    subsidiary_target = re.search(
        r'(?:全资子公司|控股子公司|参股公司)(?:为)?([^，。；\n]{1,20}?(?:股份有限公司|有限公司))',
        ''.join((primary or full).split()))
    adjacent_sub = re.search(
        r'([^（(，。；]{2,40}?(?:股份有限公司|有限公司))[（(]以下简称[^）)]{1,30}[）)](?:系|为)[^。]{0,40}?(?:全资|控股|参股)(?:子)?公司',
        ''.join((primary or full).split()))
    if subsidiary_target:
        out['target'] = subsidiary_target.group(1).strip(' ：:')
    elif adjacent_sub:
        adjacent_value = adjacent_sub.group(1).strip(' ：:')
        for marker in ('基本情况', '概述', '情况如下', '如下'):
            if marker in adjacent_value:
                adjacent_value = adjacent_value.rsplit(marker, 1)[-1]
                break
        cleaned_sub = re.search(r'([^，。；:：]{2,40}(?:股份有限公司|有限公司))$', adjacent_value)
        out['target'] = ((first_company.group(1).strip(' ：:') if first_company and len(adjacent_value) < 10 else cleaned_sub.group(1).strip(' ：:'))
                         if cleaned_sub else adjacent_value)
    elif first_company:
        out['target'] = first_company.group(1).strip(' ：:')
    else:
        out['target'] = _extract_target(primary or full)
    out['generic_target'] = (_extract_target(primary or full)
                           or _extract_target(''.join((primary or full).split())))
    if re.search(r'置入资产为', ''.join((primary or full).split())):
        out['purchase_target'] = out['target']
    sale_target = re.search(r'拟将([^。]{2,220}?)(?:出售|转让)给', ''.join((primary or full).split()))
    if sale_target:
        out['sale_target'] = sale_target.group(1).strip('，,；;：:')
    sale_price = re.search(r'出售[^。]{0,100}?作价([\d,]+(?:\.\d+)?)\s*(亿元|万元|元)', ''.join((primary or full).split()))
    if sale_price:
        value = _wan(sale_price.group(1), sale_price.group(2))
        if value is not None:
            out['sale_amount_cny_wan'] = value
    opening = '\n'.join(p for page_no, p in text_pages if page_no <= 3)
    if re.search(r'权益变动性质[：:]?[^。\n]{0,50}(?:增加|减少)?[（(]?协议转让', ''.join(opening.split())):
        out['payment_method'] = '协议转让'
    else:
        out['payment_method'] = _extract_payment(primary or full)
    flat_primary_invest = ''.join((primary or full).split())
    out['external_investor'] = bool(re.search(
        r'引入(?:战略|外部)?投资者|外部投资者|本轮投资[人方]|新投资者|引入新投资人|增资方|认购新增注册资本|基金[^。]{0,40}拟以现金方式[^。]{0,20}增资|拟向[^。]{0,60}投资人民币|拟与[^。]{0,400}(?:基金|创投)[^。]{0,120}共同增资|向[^。]{0,40}基金发行优先股',
        flat_primary_invest))
    external_markers = bool(re.search(r'外部投资者|新投资者|本轮投资者|引入外部|引入战略投资者', flat_primary_invest))
    affiliated_increase = bool(re.search(
        r'引入员工持股平台|控股股东及实际控制人[^。]{0,60}增资|公司控股股东[^。]{0,80}(?:出资|认购)|全资子公司[^。]{0,160}(?:认缴|出资)[^。]{0,50}(?:新增注册资本|增资)|全资子公司[^。]{0,160}以[^。]{0,30}认缴|与关联人共同[^。]{0,40}增资',
        flat_primary_invest))
    self_asset_increase = bool(
        re.search(r'实物资产|作价增资|对子公司增资', flat_primary_invest)
        and not external_markers
        and not re.search(r'引入外部投资者|增资扩股引进外部|新投资者|新增投资方', flat_primary_invest))
    out['strategic_non_third_party'] = (affiliated_increase and not external_markers) or self_asset_increase
    out['advisors'] = _extract_advisors(primary or full)
    out.update(_extract_control_facts(full))
    if out.get('payment_method') == '协议转让':
        controller = re.search(r'控股股东、实际控制人([^，。；]{2,20}?)(?:先生|女士)', compact_full)
        if controller:
            out['old_controller'] = controller.group(1).strip(' ：“”')
        price = re.search(r'(?:以人民币|转让价格为人民币|每股人民币)([\d.]+)元/股', compact_full)
        if price:
            out['price_per_share'] = price.group(1)
        agreement = re.search(r'于(20\d{2})年(\d{1,2})月(\d{1,2})日签署', compact_full)
        if agreement:
            try:
                out['agreement_date'] = date(int(agreement.group(1)), int(agreement.group(2)), int(agreement.group(3))).isoformat()
            except ValueError:
                pass
        lockup = re.search(r'(\d+)个月内[^。]{0,24}?减持', compact_full)
        if lockup:
            out['lockup'] = f'{lockup.group(1)}个月'
        seller_table = compact_full[compact_full.find('转让方合计'):][:180] if '转让方合计' in compact_full else ''
        ratios = re.findall(r'(?:100|[0-9]{1,2})\.[0-9]{2}%', seller_table)
        ratios = [x.rstrip('%') for x in ratios]
        if len(ratios) >= 3:
            out['controller_ratio_before_pct'] = ratios[0]
            out['transfer_ratio_pct'] = (str(out['ratio_pct']) if out.get('ratio_pct') is not None else ratios[1])
            out['controller_ratio_after_pct'] = ratios[2]
    flat_primary = ''.join((primary or full).split())
    if re.search(r'终止|撤回', flat_primary):
        if re.search(r'已受理|受理后|中止审核|审核问询|回复问询|重组委', flat_primary):
            out['termination_stage'] = '在审'
        elif re.search(r'预案|筹划阶段|方案披露阶段', flat_primary):
            out['termination_stage'] = '预案后'
    flat_full = ''.join(full.split())
    meeting = re.search(r'于(20\d{2})年(\d{1,2})月(\d{1,2})日[^。]{0,40}(?:并购重组审核委员会|重组委)', flat_full)
    if not meeting:
        meeting = re.search(r'(?:并购重组审核委员会|重组委)[^。]{0,40}于(20\d{2})年(\d{1,2})月(\d{1,2})日', flat_full)
    if meeting:
        try:
            out['meeting_date'] = date(int(meeting.group(1)), int(meeting.group(2)), int(meeting.group(3))).isoformat()
        except ValueError:
            pass
    meeting_no = re.search(r'(20\d{2})年第(\d+)次[^。]{0,20}(?:并购重组审核委员会|重组委)', flat_full)
    if meeting_no:
        exchange = 'SH' if '上海证券交易所' in flat_full else ('SZ' if '深圳证券交易所' in flat_full else '')
        if exchange:
            out['meeting_id'] = f'{exchange}_{meeting_no.group(1)}_{meeting_no.group(2)}'
    if not out['excerpt']:
        first = next((p for _, p in text_pages if len(p.strip()) >= 40), '')
        out['excerpt'] = SPACE_RE.sub(' ', first)[:500]
    summary_pages = [p for _, p in text_pages[:2] if p.strip()]
    if summary_pages:
        out['summary'] = SPACE_RE.sub(' ', '\n'.join(summary_pages))[:1200]
    return out


class LedgerBuilder:
    def __init__(self, sources_dir: Path, start: str, end: str, backfill_start: str | None = None):
        self.dir = sources_dir
        self.start = start
        self.end = end
        self.backfill_start = backfill_start or start
        self.sources: list[dict] = []
        self.cases: dict[str, dict] = {}
        self.events: list[dict] = []
        self.pending: list[dict] = []
        self.seen_events: set[tuple] = set()
        self.seen_sources: set[str] = set()
        self.code_kind_deals: dict[tuple[str, str], str] = {}
        self.announcements: list[dict] = []
        self.pdf_cache = sources_dir.parent / 'pdf-cache'

    def _iso(self, d8: str) -> str:
        s = (d8 or '').replace('-', '')
        return f'{s[:4]}-{s[4:6]}-{s[6:8]}' if len(s) >= 8 else ''

    def add_source(self, sid: str, locator: str, excerpt: str, url: str,
                   published_on: str, local: str = '') -> str:
        if sid in self.seen_sources:
            existing = next((s for s in self.sources if s['source_id'] == sid), None)
            if existing:
                if local and not existing.get('local_path'):
                    existing['local_path'] = local
                if url and not existing.get('file_or_url'):
                    existing['file_or_url'] = url
                if excerpt and len(excerpt) > len(existing.get('excerpt', '')):
                    existing['excerpt'] = excerpt[:1000]
            return sid
        self.seen_sources.add(sid)
        self.sources.append({'source_id': sid, 'locator': locator,
                             'excerpt': clean_text(excerpt)[:1000], 'file_or_url': url or local,
                             'published_on': published_on, 'local_path': local})
        return sid

    def _deal_id(self, code: str, kind: str, name: str = '') -> str:
        key = (code, kind)
        if key not in self.code_kind_deals:
            self.code_kind_deals[key] = f'CN_{code}_{kind}'
        return self.code_kind_deals[key]

    def add_case(self, deal_id: str, code: str, name: str, kind: str,
                 board: str | None = None, note: str = '') -> dict:
        code = re.sub(r'[^0-9A-Za-z]', '', clean_text(code))
        good_kind = kind if kind in KINDS else 'ma_audit'
        if deal_id not in self.cases:
            self.cases[deal_id] = {
                'deal_id': deal_id, 'code': exchange_code(code), 'name': clean_text(name),
                'kind': good_kind, 'board': board or board_of(code), 'advisors': [],
                'amounts': [], 'supporting_financing': [], 'also_kinds': [],
                'facts': {}, 'fact_evidence': {}, 'source_document_ids': [],
                'note': note,
            }
            self.code_kind_deals.setdefault((code, good_kind), deal_id)
        case = self.cases[deal_id]
        if kind in KINDS and kind != case['kind'] and kind not in case['also_kinds']:
            case['also_kinds'].append(kind)
        return case

    def add_event(self, deal_id: str, ev_type: str, day: str, source_ids: list[str],
                  status: str = 'CONFIRMED', meeting_id: str = '', note: str = '',
                  announced_on: str = '', effective_on: str = '') -> bool:
        unique_types = {'first_disclosed', 'first_draft', 'accepted', 'committee_passed', 'committee_deferred',
                        'committee_rejected', 'simplified_passed', 'registration_submitted',
                        'registered', 'terminated', 'completed', 'control_completed', 'control_progress'}
        key = (deal_id, ev_type, day)
        existing = next((e for e in self.events if e.get('deal_id') == deal_id and
                         e.get('type') == ev_type and
                         (ev_type in unique_types or e.get('date') == day)), None)
        if existing:
            existing['source_ids'] = sorted(set(existing.get('source_ids', [])) | set(source_ids))
            if ev_type in unique_types and day < existing.get('date', day):
                existing['date'] = day
                existing['announced_on'] = announced_on or day
            if status == 'CONFIRMED':
                existing['status'] = 'CONFIRMED'
            if note and note not in existing.get('note', ''):
                existing['note'] = (existing.get('note', '') + '；' + note).strip('；')
            return False
        seq = sum(1 for e in self.events if e['deal_id'] == deal_id and e['date'] == day)
        eid = 'EV_' + hashlib.md5(f'{deal_id}|{ev_type}|{day}|{seq}'.encode()).hexdigest()[:12]
        ev = {'event_id': eid, 'deal_id': deal_id, 'type': ev_type, 'status': status,
              'date': day, 'announced_on': announced_on or day, 'sequence': seq,
              'source_ids': source_ids, 'note': note}
        if effective_on:
            ev['effective_on'] = effective_on
        if ev_type.startswith('committee_') and ev_type != 'committee_scheduled':
            ev['meeting_id'] = meeting_id or f'MEETING_{day.replace("-", "")}_{deal_id[-6:]}'
        self.events.append(ev)
        self.seen_events.add(key)
        return True

    def _existing_audit_deal(self, code: str) -> str | None:
        for did, case in self.cases.items():
            if case['code'].split('.')[0] == code and case['kind'] == 'ma_audit':
                return did
        return None

    def from_cninfo(self):
        f = self.dir / 'cninfo_announcements.json'
        if not f.exists():
            self.pending.append({'item': '巨潮公告清单缺失', 'reason': '采集阶段未产出',
                                 'action': '需补检索'})
            return
        data = json.loads(f.read_text(encoding='utf-8'))
        self.announcements = [
            dict(a, title=_clean_title(a.get('title', '')), sec_name=_clean_title(a.get('sec_name', '')))
            for a in data.get('announcements', [])
            if a.get('published_on') and self._iso(self.backfill_start) <= a['published_on'] <= self._iso(self.end)
        ]
        groups: dict[str, list[dict]] = defaultdict(list)
        for a in self.announcements:
            groups[clean_text(a.get('sec_code', ''))].append(a)

        for code, anns in sorted(groups.items()):
            # Parse representative documents only after title rules identify a
            # real milestone or a genuinely ambiguous control candidate. This
            # still scans every page of every selected PDF.
            parsed: dict[str, dict] = {}
            for a in sorted(anns, key=lambda x: (x['published_on'], x['announcement_id'])):
                title = a['title']
                hits = classify_title(title)
                ambiguous_control = bool(re.search(r'权益变动报告书|收购报告书|要约收购|协议转让.{0,20}股份.{0,30}股东权益变动', title)) and not bool(re.search(r'持续督导|财务顾问|法律意见|核查意见|修订稿|摘要', title))
                generic_reorg = bool(re.search(r'筹划.{0,12}重大资产重组|重大资产重组.{0,8}(?:提示|停牌)', title)) and not bool(re.search(r'进展|风险提示|补充协议|继续停牌', title))
                strategic_candidate = bool(re.search(r'(?:子公司.{0,40}增资扩股|控股孙公司.{0,30}增资|参股公司.{0,40}增资|向参股公司增资|放弃参股公司.{0,30}(?:优先认购权|优先认缴)|子公司.{0,50}向.{0,40}基金发行优先股)', title))
                same_day_has_draft = any(
                    x.get('published_on') == a.get('published_on') and
                    any(h[0] == 'first_draft' for h in classify_title(x.get('title', '')))
                    for x in anns)
                if not hits and not ambiguous_control and not generic_reorg and not strategic_candidate:
                    self.pending.append({
                        'item': f"{a.get('sec_name')}｜{title}", 'reason': '标题不含可确认里程碑；不进入正式统计',
                        'source': a.get('url'), 'action': '仅作为线索；如需纳入须回正文补充证据'})
                    continue
                pdf_path = self.dir / 'cninfo_pdfs' / f"{a['announcement_id']}.pdf"
                facts = pdf_facts(pdf_path, self.pdf_cache) if pdf_path.is_file() else {}
                parsed[a['announcement_id']] = facts
                body = '\n'.join([])  # placeholder to keep logic local
                # For bare control reports, only create a control event if full
                # text explicitly says actual control changes or will change.
                if ambiguous_control and not hits:
                    control_result = facts.get('control_result', '')
                    old = facts.get('old_controller')
                    new = facts.get('new_controller')
                    title_change = bool(re.search(r'(?:实际控制人|控制权).{0,12}(?:拟发生|发生)?变更', title))
                    if re.search(r'要约收购事宜致全体股东(?:的)?报告书', title):
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{title}",
                            'reason': '中期/to全体股东报告书，并非新的首次控制权变更披露',
                            'source': a.get('url'), 'action': '排除：历史要约事项的阶段性文件'})
                        continue
                    summary_flat = ''.join((facts.get('summary') or facts.get('excerpt') or '').split())
                    body_current_change = bool(re.search(
                        r'(?:本次|本次交易|本次发行)[^。]{0,50}完成后[^。]{0,100}(?:实际控制人|控制权)[^。]{0,30}(?:变更|变化)|不再(?:为公司)?实际控制人|实际控制人[^。]{0,30}(?:变更为|调整为)',
                        summary_flat))
                    if title_change or (old and new and old != new and body_current_change) or (
                            control_result == '控制权拟发生变更' and body_current_change):
                        hits = [('first_disclosed', 'control', 'HIGH')]
                    elif control_result == '未导致实际控制人变化':
                        non_agreed_transfer = bool(re.search(r'司法拍卖|执行裁定|被动稀释', title + (facts.get('summary') or ''))
                                                    and not re.search(r'协议转让|股份转让协议', title))
                        summary_subject = bool(re.search(
                            r'控股股东[^。]{0,40}(?:及|和)?实际控制人[^。]{0,60}签订[^。]{0,20}转让协议',
                            ''.join((facts.get('summary') or '').split())))
                        transfer_subject = (bool(re.search(r'控股股东|实际控制人|控制权|一致行动人', title))
                                            or summary_subject) and not non_agreed_transfer
                        if facts.get('payment_method') == '协议转让' and facts.get('ratio_pct') and transfer_subject:
                            hits = [('first_disclosed', 'partial_transfer', 'HIGH')]
                        else:
                            self.pending.append({
                                'item': f"{a.get('sec_name')}｜{title}",
                                'reason': ('仅为一般股东协议转让，转让方不属于控股股东、实际控制人或一致行动人，不进入12/13表'
                                           if facts.get('payment_method') == '协议转让'
                                           else '公告全文明确实际控制人未变化，排除控制权变更表；未发现可量化的协议转让条件'),
                                'source': a.get('url'),
                                'auto_evidence': facts.get('excerpt', '')[:240],
                                'action': '排除：实控人未变化'})
                            continue
                    else:
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{title}",
                            'reason': '仅见权益变动/收购报告书标题，全文未能确认实控人变化',
                            'source': a.get('url'),
                            'auto_evidence': facts.get('excerpt', '')[:240],
                            'action': '人工回链后再纳入'})
                        continue
                strategic_excluded_title = bool(re.search(r'进展|完成|终止|募集资金.*增资|关于向全资(?:子|孙)公司增资', title))
                document_service = bool(re.search(r'持续督导|督导报告|法律意见|核查意见|专项核查|会计师事务所|律师事务所|评估报告|自查报告|问询函回复|回复.*问询|财务顾问报告', title))
                if facts.get('strategic_non_third_party'):
                    hits = [h for h in hits if h[1] != 'strategic_investment']
                if strategic_candidate and not document_service and not strategic_excluded_title and not facts.get('strategic_non_third_party') and not hits and facts.get('external_investor'):
                    hits = [('first_disclosed', 'strategic_investment', 'HIGH')]
                # Body text resolves generic "重大资产重组" into cash or
                # share consideration. It also turns some explicit control /
                # transfer announcements into an overlap transaction.
                if generic_reorg and len(hits) == 1 and hits[0][0] == 'lead_signal':
                    pm = facts.get('payment_method', '')
                    lead_kind = hits[0][1]
                    if '现金' in pm:
                        lead_kind = ('ma_cash_sale' if '出售' in title or '出售' in (facts.get('excerpt') or '')
                                     else 'ma_cash_purchase')
                    hits = [('lead_signal', lead_kind, 'LOW')]
                elif not hits and generic_reorg and not same_day_has_draft:
                    pm = facts.get('payment_method', '')
                    if '发行股份' in pm:
                        hits = [('first_disclosed', 'ma_audit', 'LOW')]
                    elif '现金' in pm:
                        generic_kind = 'ma_cash_sale' if '出售' in title or '出售' in (facts.get('excerpt') or '') else 'ma_cash_purchase'
                        hits = [('first_disclosed', generic_kind, 'LOW')]
                    else:
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{title}",
                            'reason': '重大资产重组方向已披露，但支付方式或交易类型未能从正文确认',
                            'source': a.get('url'), 'auto_evidence': facts.get('excerpt', '')[:240],
                            'action': '确认是否发行股份、现金购买或现金出售后再入表'})
                        continue
                for ev, kind, confidence in hits:
                    body_actual_change = bool(re.search(
                        r'(?:控股股东|实际控制人)自?[^。]{0,40}(?:起)?由[^。]{0,80}变更为|控股股东变更为[^。]{0,60}(?:，|。|；)[^。]{0,60}实际控制人变更为',
                        ''.join((facts.get('summary') or facts.get('excerpt') or '').split())))
                    if body_actual_change:
                        facts['control_result'] = '实际控制人发生变化'
                    if (kind == 'control' and ev == 'first_disclosed' and facts
                            and facts.get('completion_date') and facts.get('control_result') == '实际控制人发生变化'
                            and re.search(r'控股股东.{0,20}(?:及|和)?实际控制人.{0,20}变更的公告', title)
                            and not re.search(r'拟|可能', title)):
                        ev = 'control_completed'
                    if (ev == 'control_completed' and facts
                            and facts.get('control_result') == '未导致实际控制人变化'
                            and not re.search(r'控制权.{0,8}(?:发生)?变更', title)):
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{title}",
                            'reason': '仅为控股股东/实控人协议转让部分股份完成，正文明确未导致控制权变化',
                            'source': a.get('url'),
                            'auto_evidence': facts.get('excerpt', '')[:240],
                            'action': '排除：控制权未变化'})
                        continue
                    progress_only = False
                    if ev == 'control_completed' and facts:
                        title_completion = bool(re.search(
                            r'完成过户|过户完成|发行完毕|控制权变更完成|完成工商变更登记|实际控制人变更暨.{0,30}换发营业执照', title))
                        if title_completion or (body_actual_change and facts.get('completion_date')) or (facts.get('completion_confirmed') and facts.get('completion_date')):
                            facts['control_result'] = '实际控制人发生变化'
                            if not facts.get('completion_date'):
                                facts['completion_date'] = a['published_on']
                        else:
                            # "完成部分过户暨控制权拟发生变更" is a progress
                            # notice, not a completed control transfer.  Keep
                            # the latest-announcement update without moving the
                            # deal into the completion table.
                            ev = 'control_progress'
                            progress_only = True
                            facts['control_result'] = '控制权拟发生变更'
                    affiliated_partnership = bool(re.search(
                        r'子公司.{0,60}控股股东.{0,100}(?:共同增资|增资协议|合伙协议)',
                        title + (facts.get('summary') or '')))
                    if kind == 'professional_cooperation' and affiliated_partnership:
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{title}",
                            'reason': '子公司与控股股东或其下属主体共同增资，属于关联方投入而非引入专业投资机构',
                            'source': a.get('url'),
                            'auto_evidence': facts.get('excerpt', '')[:240],
                            'action': '排除：关联方共同增资，不纳入专业机构合作'})
                        continue
                    if kind in ('control', 'partial_transfer', 'second_shareholder', 'ipo'):
                        deal_id = self._deal_id(code, kind)
                        # Keep control and partial-transfer as separate
                        # filterable deal kinds. The 12 table excludes deals
                        # that change control, while the control table uses
                        # different eligibility and completion tests.
                    elif kind == 'ma_audit' or (ev in AUDIT_EVENTS and kind == 'ma_audit'):
                        deal_id = self._existing_audit_deal(code) or self._deal_id(code, 'ma_audit')
                    else:
                        deal_id = self._deal_id(code, kind)
                    note = '公告标题明确；结论性字段以PDF全文为准'
                    case = self.add_case(deal_id, code, a.get('sec_name', ''), kind, note=note)
                    if a['announcement_id'] not in case['source_document_ids']:
                        case['source_document_ids'].append(a['announcement_id'])
                    if facts:
                        if kind == 'strategic_investment':
                            facts['target'] = facts.get('target') or facts.get('generic_target', '')
                        elif kind in {'ma_audit', 'ma_cash_purchase', 'ma_cash_sale'}:
                            facts['target'] = facts.get('generic_target', '')
                        new_milestone = ev in ('first_disclosed', 'first_draft')
                        if ev == 'control_completed' and not facts.get('completion_date'):
                            facts['completion_date'] = a['published_on']
                        transaction_amount = facts.get('amount_cny_wan')
                        if kind == 'ma_cash_sale' and facts.get('sale_amount_cny_wan') is not None:
                            transaction_amount = facts['sale_amount_cny_wan']
                            facts['target'] = facts.get('sale_target') or facts.get('target', '')
                        elif kind == 'ma_cash_purchase' and facts.get('purchase_target'):
                            facts['target'] = facts['purchase_target']
                        if new_milestone and transaction_amount:
                            case['facts']['transaction_amount_wan'] = transaction_amount
                            case['fact_evidence']['transaction_amount_wan'] = facts['fact_pages'].get('amount_cny_wan', {})
                            if not any(abs(float(x.get('value_wan', 0)) - transaction_amount) < 0.01 for x in case['amounts']):
                                case['amounts'].append({
                                    'basis': 'latest_plan', 'currency': 'CNY',
                                    'value_wan': str(transaction_amount), 'known_on': a['published_on'],
                                    'source_ids': ['CN_' + a['announcement_id']],
                                    'locator': f"{facts['fact_pages'].get('amount_cny_wan', {}).get('page','?')}页｜{facts['excerpt'][:200]}"})
                        if ev == 'control_completed' and facts.get('control_transaction_amount_cny_wan'):
                            actual_amount = float(facts['control_transaction_amount_cny_wan'])
                            case['facts']['actual_transaction_amount_wan'] = actual_amount
                            if not any(abs(float(x.get('value_wan', 0)) - actual_amount) < 0.01
                                       and x.get('basis') == 'actual_completion' for x in case['amounts']):
                                case['amounts'].append({
                                    'basis': 'actual_completion', 'currency': 'CNY',
                                    'value_wan': str(actual_amount), 'known_on': a['published_on'],
                                    'source_ids': ['CN_' + a['announcement_id']]})
                        if new_milestone and facts.get('supporting_financing_cny_wan'):
                            case['facts']['supporting_financing_wan'] = facts['supporting_financing_cny_wan']
                            case['fact_evidence']['supporting_financing_wan'] = facts['fact_pages'].get('supporting_financing_cny_wan', {})
                            case['supporting_financing'].append({
                                'basis': 'latest_plan', 'currency': 'CNY',
                                'value_wan': str(facts['supporting_financing_cny_wan']),
                                'known_on': a['published_on'], 'source_ids': ['CN_' + a['announcement_id']]})
                        if kind in ('strategic_investment', 'professional_cooperation', 'control', 'partial_transfer', 'second_shareholder', 'ma_audit', 'ma_cash_purchase', 'ma_cash_sale') and (facts.get('summary') or facts.get('excerpt')):
                            case['facts']['summary'] = facts.get('summary') or facts['excerpt']
                        if kind == 'partial_transfer' and facts.get('ratio_pct') is not None:
                            facts['transfer_ratio_pct'] = facts['ratio_pct']
                        if kind == 'partial_transfer' and facts.get('controller_ratio_after_pct'):
                            facts['control_ratio_pct'] = facts['controller_ratio_after_pct']
                            case['facts']['control_ratio_pct'] = facts['controller_ratio_after_pct']
                        for key in ('target', 'payment_method', 'ratio_pct', 'control_result',
                                    'old_controller', 'new_controller', 'completion_date', 'termination_stage',
                                    'price_per_share', 'agreement_date', 'lockup',
                                    'controller_ratio_before_pct', 'controller_ratio_after_pct',
                                    'transfer_ratio_pct', 'control_transaction_amount_cny_wan'):
                            if facts.get(key) not in ('', None, []):
                                case['facts'][key] = facts[key]
                        for advisor in facts.get('advisors', []):
                            canonical = canonical_advisor(advisor)
                            if canonical and canonical != case.get('name') and not any(x.get('name') == canonical and x.get('role') == 'independent' for x in case['advisors']):
                                case['advisors'].append({
                                    'name': canonical, 'role': 'independent',
                                    'disclosed_on': a['published_on'],
                                    'source_ids': ['CN_' + a['announcement_id']]})
                    # A title's existence/type/date is confirmed, but ambiguous
                    # transfer eligibility remains a formal pending-review item.
                    sid = self.add_source(
                        'CN_' + a['announcement_id'],
                        f"巨潮公告｜{a['published_on']}｜{a['title']}" +
                        (f"｜正文{facts.get('fact_pages',{}).get('amount_cny_wan',{}).get('page')}页" if facts.get('fact_pages',{}).get('amount_cny_wan',{}).get('page') else ''),
                        facts.get('excerpt') or a['title'], a.get('url', ''), a['published_on'],
                        str(pdf_path) if pdf_path.exists() else '')
                    status = 'CONFIRMED' if confidence == 'HIGH' else 'CANDIDATE'
                    if ev == 'control_completed' and not facts.get('completion_date'):
                        facts['completion_date'] = a['published_on']
                    event_day = facts.get('meeting_date') if ev in {'committee_passed', 'committee_deferred', 'committee_rejected'} else ''
                    self.add_event(deal_id, ev, event_day or a['published_on'], [sid], status=status,
                                   meeting_id=facts.get('meeting_id', '') if ev.startswith('committee_') else '',
                                   announced_on=a['published_on'], note=f'标题规则；置信度{confidence}',
                                   effective_on=facts.get('completion_date', '') if ev == 'control_completed' else '')
                    trusted_body_completion = (facts.get('completion_confirmed')
                                               and facts.get('completion_date')
                                               and facts.get('control_result') == '实际控制人发生变化')
                    if kind == 'control' and ev != 'control_completed' and (
                            trusted_body_completion
                            or (re.search(r'一致行动.{0,20}(?:到期终止|到期不再续签|解除)', title)
                                and facts.get('control_result') == '实际控制人发生变化')
                            or (re.search(r'控制权变更完成', title)
                                and facts.get('control_result') == '实际控制人发生变化')):
                        self.add_event(deal_id, 'control_completed', a['published_on'], [sid],
                                       status='CONFIRMED', announced_on=a['published_on'],
                                       effective_on=facts.get('completion_date', ''),
                                       note='公告正文显示控制权变更已完成')
                    if progress_only:
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{title}",
                            'reason': '仅确认部分过户或控制权拟发生变更，正文未确认本次已完成全部控制权变更',
                            'source': a.get('url'),
                            'auto_evidence': facts.get('excerpt', '')[:240],
                            'action': '核对是否仅为阶段性进展；完成表暂不新增'})
                    if kind in ('control', 'partial_transfer', 'second_shareholder', 'ma_cash_purchase', 'ma_cash_sale'):
                        self.pending.append({
                            'item': f"{a.get('sec_name')}｜{a['title']}",
                            'reason': '已确认公告存在和标题所示事件；是否纳入对应专表仍需核对控制权、持股排序、协议转让比例等正文条件',
                            'source': a.get('url'),
                            'auto_evidence': (facts.get('fact_pages', {}).get('transaction_amount_wan', {}).get('excerpt') or facts.get('excerpt', ''))[:240],
                            'action': '按专表资格二次筛选'})

    def from_bse(self):
        f = self.dir / 'bse_momentum.json'
        if not f.is_file():
            self.pending.append({'item': '北交所并购重组项目动态', 'reason': '未产出结构化快照',
                                 'action': '通过浏览器核验项目页后导出 bse_momentum.json'})
            return
        try:
            obj = json.loads(f.read_text(encoding='utf-8'))
        except ValueError:
            self.pending.append({'item': '北交所并购重组项目动态', 'reason': 'JSON解析失败',
                                 'action': '重新核验项目页'})
            return
        rows = obj.get('data', []) if isinstance(obj, dict) else obj
        for r in rows:
            day = str(r.get('updtdt') or r.get('updtDt') or '')[:10]
            if not day or not (self._iso(self.start) <= day <= self._iso(self.end)):
                continue
            code = re.sub(r'\D', '', str(r.get('cmpcode') or r.get('zqdm') or ''))[:6]
            name = clean_text(r.get('cmpsnm') or r.get('gfmc') or '')
            state = clean_text(r.get('prjst') or r.get('shzt') or '')
            ev = lookup_state(state, BSE_STATE_MAP)
            if not ev:
                self.pending.append({'item': f'{name}｜{state}',
                                     'reason': '北交所审核状态未映射，需人工归类',
                                     'source': 'BSE_MOMENTUM', 'action': '补状态映射'})
                continue
            sid = self.add_source(
                'BSE_' + hashlib.md5(f'{code}|{day}|{state}'.encode()).hexdigest()[:10],
                f'北交所并购重组项目动态｜{code} {name}｜{state}',
                json.dumps(r, ensure_ascii=False)[:900],
                'https://www.bse.cn/audit/t_project_news_r.html', day, str(f))
            deal_id = self._existing_audit_deal(code) or self._deal_id(code, 'ma_audit')
            case = self.add_case(deal_id, code, name, 'ma_audit', '北交所',
                                 note='北交所结构化项目动态；与公司公告按证券代码合并')
            advisor = canonical_advisor(r.get('sprinsts') or r.get('sprinst') or '')
            if advisor and advisor != name and '未披露' not in advisor and len(advisor) < 80:
                if not any(x.get('name') == advisor for x in case['advisors']):
                    case['advisors'].append({'role': 'independent', 'name': advisor,
                                             'disclosed_on': day, 'source_ids': [sid]})
            self.add_event(deal_id, ev, day, [sid], note=f'北交所审核状态：{state}',
                           meeting_id=f'BSE_{day}' if ev.startswith('committee_') else '')

    def from_szse(self):
        f = self.dir / 'szse_momentum.json'
        if not f.exists():
            self.pending.append({'item': '深交所项目动态本期窗口', 'reason': '未产出',
                                 'action': '需补检索'})
            return
        try:
            obj = json.loads(f.read_text(encoding='utf-8'))
        except ValueError:
            self.pending.append({'item': '深交所项目动态', 'reason': 'JSON解析失败', 'action': '重新采集'})
            return
        rows = obj.get('data', []) if isinstance(obj, dict) else []
        for r in rows:
            day = str(r.get('updtdt') or '')[:10]
            accepted_day = str(r.get('acptdt') or '')[:10]
            accepted_in_window = bool(accepted_day and self._iso(self.start) <= accepted_day <= self._iso(self.end))
            update_in_window = bool(day and self._iso(self.start) <= day <= self._iso(self.end))
            if not update_in_window and not accepted_in_window:
                continue
            code = re.sub(r'\D', '', str(r.get('cmpcode') or ''))[:6]
            name = clean_text(r.get('cmpsnm') or r.get('cmpnm') or '')
            state = clean_text(r.get('prjst') or '')
            ev = lookup_state(state, SZSE_STATE_MAP)
            if not ev:
                self.pending.append({'item': f'{name}｜{state}',
                                     'reason': '深交所审核状态未映射，需人工归类',
                                     'source': 'SZSE_MOMENTUM', 'action': '补状态映射'})
                continue
            sid = self.add_source(
                'SZ_' + hashlib.md5(f'{code}|{day}|{state}'.encode()).hexdigest()[:10],
                f'深交所并购重组项目动态｜{code} {name}｜{state}',
                json.dumps(r, ensure_ascii=False)[:900],
                'https://www.szse.cn/listing/projectdynamic/reproperty/index.html', day,
                str(f))
            deal_id = self._existing_audit_deal(code) or self._deal_id(code, 'ma_audit')
            board = '创业板' if str(r.get('boardCode')) == '16' else '深主板'
            case = self.add_case(deal_id, code, name, 'ma_audit', board,
                                 note='深交所结构化项目动态；与公司公告按证券代码合并')
            advisor = canonical_advisor(r.get('sprinsts') or r.get('sprinst') or '')
            if advisor and advisor != name and '未披露' not in advisor and len(advisor) < 80:
                if not any(x.get('name') == advisor for x in case['advisors']):
                    case['advisors'].append({'role': 'independent', 'name': advisor,
                                             'disclosed_on': day, 'source_ids': [sid]})
            terminal_state = ev in {'committee_passed', 'committee_rejected', 'simplified_passed',
                                    'registration_submitted', 'registered', 'terminated'}
            if accepted_in_window and day != accepted_day and not terminal_state:
                acc_sid = self.add_source(
                    'SZ_ACC_' + hashlib.md5(f'{code}|{accepted_day}'.encode()).hexdigest()[:10],
                    f'深交所并购重组项目动态｜{code} {name}｜受理日期',
                    json.dumps(r, ensure_ascii=False)[:900],
                    'https://www.szse.cn/listing/projectdynamic/reproperty/index.html',
                    accepted_day, str(f))
                self.add_event(deal_id, 'accepted', accepted_day, [acc_sid],
                               note='深交所结构化项目动态：受理日期')
            if update_in_window:
                self.add_event(deal_id, ev, day, [sid], note=f'深交所审核状态：{state}',
                               meeting_id=f"SZSE_{day}" if ev.startswith('committee_') else '')

    def from_sse(self):
        f = self.dir / 'sse_momentum_window.json'
        if not f.exists():
            self.pending.append({'item': '上交所项目动态本期窗口', 'reason': '未产出',
                                 'action': '需补检索'})
            return
        rows = json.loads(f.read_text(encoding='utf-8'))
        for r in rows:
            day = (r.get('更新日期') or '').replace('/', '-')[:10]
            if not day or not (self._iso(self.start) <= day <= self._iso(self.end)):
                continue
            code = re.sub(r'\D', '', r.get('公司代码') or '')[:6]
            name = clean_text(r.get('公司简称') or '')
            state = r.get('审核状态') or ''
            advisor = canonical_advisor(r.get('独立财务顾问') or '')
            ev = lookup_state(state, SSE_STATE_MAP)
            if not ev:
                self.pending.append({'item': f'{name}｜{state}',
                                     'reason': '审核状态未映射，需人工归类', 'source': 'SSE_MOMENTUM',
                                     'action': '补状态映射'})
                continue
            sid = self.add_source('SSE_' + hashlib.md5(
                f'{code}|{day}|{state}'.encode()).hexdigest()[:10],
                f'上交所并购重组项目动态｜{code} {name}｜{state}',
                json.dumps(r, ensure_ascii=False)[:500],
                'https://www.sse.com.cn/listing/renewal/ma/', day,
                str(self.dir / 'sse_momentum.xls'))
            deal_id = self._existing_audit_deal(code) or self._deal_id(code, 'ma_audit')
            board = '科创板' if r.get('板块') == '科创板' else '沪主板'
            case = self.add_case(deal_id, code, name, 'ma_audit', board,
                                 note='上交所结构化项目动态；与公司公告按证券代码合并')
            if advisor and advisor != name and '未披露' not in advisor and len(advisor) < 80:
                if not any(x.get('name') == advisor for x in case['advisors']):
                    case['advisors'].append({'role': 'independent', 'name': advisor,
                                             'disclosed_on': day, 'source_ids': [sid]})
            meeting_id = ''
            if ev.startswith('committee_'):
                meeting_id = f"SSE_{day.replace('-', '')}"
            self.add_event(deal_id, ev, day, [sid], note=f'上交所审核状态：{state}',
                           meeting_id=meeting_id)

    def build(self) -> dict:
        self.from_sse()
        self.from_szse()
        self.from_bse()
        self.from_cninfo()
        # Same-day duplicate milestones can arise after normalized titles; keep
        # one event per deal/type/date and union source evidence.
        dedup: dict[tuple, dict] = {}
        for e in self.events:
            key = (e['deal_id'], e['type'], e['date'])
            if key not in dedup:
                dedup[key] = e
            else:
                dedup[key]['source_ids'] = sorted(set(dedup[key]['source_ids']) | set(e['source_ids']))
                if len(e.get('note', '')) > len(dedup[key].get('note', '')):
                    dedup[key]['note'] = e['note']
        self.events = list(dedup.values())
        period = {'start': self._iso(self.start), 'end': self._iso(self.end),
                  'cutoff': self._iso(self.end)}
        cov = json.loads((self.dir.parent / 'source-manifest.json').read_text(encoding='utf-8'))['coverage'] \
            if (self.dir.parent / 'source-manifest.json').exists() else [
                {'group': 'ACQUIRE', 'status': 'NOT_RUN', 'evidence': '采集清单缺失'}]
        ledger = {'period': period, 'sources': self.sources,
                  'cases': list(self.cases.values()), 'events': self.events, 'coverage': cov,
                  'limitations': [
                      '事件存在性与类型以官方公告清单/交易所结构化数据为据',
                      '金额、比例、控制权等结论只取PDF全文命中，抽不到留空；对应事项见待核清单',
                      '同一公司同类型交易在本期按证券代码+交易类型归并；后续加入标的名称指纹处理并行交易']}
        return {'ledger': ledger, 'pending': self.pending}


def write_sqlite(db: Path, ledger: dict):
    if db.exists():
        db.unlink()
    con = sqlite3.connect(db)
    con.executescript('''
    CREATE TABLE deals(deal_id TEXT PRIMARY KEY, code TEXT, name TEXT, board TEXT, kind TEXT, note TEXT);
    CREATE TABLE events(event_id TEXT PRIMARY KEY, deal_id TEXT, type TEXT, status TEXT,
                        date TEXT, announced_on TEXT, sequence INTEGER, meeting_id TEXT, note TEXT);
    CREATE TABLE event_sources(event_id TEXT, source_id TEXT, PRIMARY KEY(event_id, source_id));
    CREATE TABLE evidence(source_id TEXT PRIMARY KEY, locator TEXT, excerpt TEXT,
                          file_or_url TEXT, published_on TEXT, local_path TEXT);
    CREATE TABLE coverage("group" TEXT PRIMARY KEY, status TEXT, evidence TEXT, note TEXT);
    ''')
    con.executemany('INSERT INTO deals VALUES(?,?,?,?,?,?)',
                    [(c['deal_id'], c['code'], c['name'], c['board'], c['kind'], c.get('note', ''))
                     for c in ledger['cases']])
    con.executemany('INSERT INTO events VALUES(?,?,?,?,?,?,?,?,?)',
                    [(e['event_id'], e['deal_id'], e['type'], e['status'], e['date'],
                      e['announced_on'], e['sequence'], e.get('meeting_id', ''), e.get('note', ''))
                     for e in ledger['events']])
    con.executemany('INSERT INTO event_sources VALUES(?,?)',
                    [(e['event_id'], s) for e in ledger['events'] for s in e['source_ids']])
    con.executemany('INSERT INTO evidence VALUES(?,?,?,?,?,?)',
                    [(s['source_id'], s['locator'], s['excerpt'], s['file_or_url'],
                      s['published_on'], s.get('local_path', '')) for s in ledger['sources']])
    con.executemany('INSERT INTO coverage VALUES(?,?,?,?)',
                    [(c['group'], c['status'], c['evidence'], c.get('note', '')) for c in ledger['coverage']])
    con.commit()
    con.close()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--sources', required=True, type=Path)
    ap.add_argument('--start', required=True)
    ap.add_argument('--end', required=True)
    ap.add_argument('--out', required=True, type=Path)
    ap.add_argument('--db', type=Path)
    ap.add_argument('--pending', type=Path)
    ap.add_argument('--backfill-start', help='Optional earlier source range start for late backfill; metrics remain bounded by --start')
    args = ap.parse_args()
    result = LedgerBuilder(args.sources, args.start, args.end, args.backfill_start).build()
    blocking_actions = ('按专表资格二次筛选', '人工回链后再纳入',
                        '确认是否发行股份、现金购买或现金出售后再入表',
                        '补状态映射', '重新采集', '重新核验')
    items = result['pending']
    for item in items:
        item['blocking'] = any(str(item.get('action', '')).startswith(a) for a in blocking_actions)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(result['ledger'], ensure_ascii=False, indent=2), encoding='utf-8')
    (args.pending or args.out.with_name('pending-review.json')).write_text(
        json.dumps({'items': items, 'count': len(items),
                    'blocking_count': sum(bool(x.get('blocking')) for x in items)},
                   ensure_ascii=False, indent=2), encoding='utf-8')
    if args.db:
        write_sqlite(args.db, result['ledger'])
    print(json.dumps({'sources': len(result['ledger']['sources']),
                      'cases': len(result['ledger']['cases']),
                      'events': len(result['ledger']['events']),
                      'pending': len(result['pending']), 'out': str(args.out)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())