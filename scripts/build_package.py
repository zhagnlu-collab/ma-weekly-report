#!/usr/bin/env python3
"""从上期 16 个原生文件生成本期 16 个原生文件（OOXML 局部补丁，保真）。

三件事：
  1. 文件名、期号、统计区间、页眉页脚与口径锚点的安全文本替换；
  2. 年度累计数字按"上期基数 + 本期增量"桥接，而不是用本期台账重算覆盖；
  3. 本期事件按真实表头列名追加到对应 Excel，并以 RGB(226,239,218) 标记新增行。
上期目录只读；业务正文中的历史案例日期、图表、批注、外链、合并区一律保留。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import ooxml  # noqa: E402
from lib import celltext  # noqa: E402
from lib.period import Period, map_filename, text_replacements  # noqa: E402

import openpyxl  # noqa: E402

NATIVE = ('.xlsx', '.xlsm', '.docx', '.pptx')
COLUMN_ALIASES = {
    'name': ['股票简称', '证券简称', '公司名称', '公司简称', '上市公司', '简称'],
    'code': ['股票代码', '证券代码', '公司代码', '代码'],
    'date': ['首次披露时间', '首次公告日', '提示性公告披露日期', '实控人变更提示性公告披露日期', '过会时间', '公告日期', '公告时间', '披露日期', '更新日期'],
    'latest': ['最新公告时间', '更新日期', '最新披露日期'],
    'completion': ['完成日期'],
    'seq': ['序号'],
    'in_review': ['是否在审（已通过重组委审核的都不纳入计算）', '是否在审'],
    'board': ['所属板块', '板块'],
    'title': ['公告名称', '公告标题', '项目名称', '交易概述', '交易概要', '备注'],
    'target': ['收购标的', '标的资产', '子公司名称', '标的公司'],
    'method': ['支付方式', '交易方式'],
    'amount': ['交易金额合计（万元）', '交易金额（万元）', '交易金额'],
    'support': ['募集配套资金（万元）', '配套融资金额（万元）', '配套融资（万元）'],
    'advisor': ['独立财务顾问', '财务顾问'],
    'summary': ['交易概述', '交易概要', '交易方案', '合作具体情况', '引战具体情况'],
    'progress': ['进展', '审核状态'],
    'old_controller': ['原实控人', '上市公司实际控制人'],
    'new_controller': ['新实控人'],
    'control_ratio': ['新实控人控制股比', '转让后上市公司实控人控制股比'],
    'control_ratio_before': ['原实控人控制股比', '转让前上市公司实控人控制股比'],
    'transfer_ratio': ['协议转让股比'],
    'price_per_share': ['交易价格（元/股）', '交易价格'],
    'agreement_date': ['股份转让协议签署日', '协议签署日'],
    'lockup': ['受让方股份锁定期', '股份锁定期'],
    'related': ['是否构成关联交易', '关联交易'],
    'major_reorg': ['重大资产重组'],
    'reverse_merger': ['重组上市'],
    'termination_stage': ['终止/撤回阶段'],
    'termination_reason': ['终止原因'],
    'note': ['备注'],
}
BRIDGE_PATTERNS = [
    ('in_review', r'在审项目共\s*(\d+)\s*单'),
    ('meetings', r'本年度共召开\s*(\d+)\s*次并购重组委会议'),
    ('registered', r'本年度共\s*(\d+)\s*单重组项目取得证监会注册批复'),
    ('first_disclosed', r'本年度审核类首次披露项目共计\s*(\d+)\s*单'),
]
BOARD_PAT = re.compile(r'(沪主板|深主板|科创板|创业板|北交所)\s*(\d+)\s*单')


def prev_native(prev: Path) -> list[Path]:
    return sorted(p for p in prev.iterdir()
                  if p.is_file() and p.suffix.lower() in NATIVE and not p.name.startswith('~$'))


def joined_text(files: list[Path]) -> str:
    return ''.join(ooxml.collect_text(p) for p in files
                   if p.suffix.lower() in ('.docx', '.pptx'))


def deltas(ledger: dict) -> dict:
    cases = {c['deal_id']: c for c in ledger.get('cases', [])}
    confirmed = [e for e in ledger.get('events', []) if e.get('status') == 'CONFIRMED']
    def ma(e):
        return cases.get(e['deal_id'], {}).get('kind') == 'ma_audit'
    # A transaction disclosed as a plan and then as a draft in the same week is
    # still one new transaction. Count unique deal ids, not event rows.
    first_ids = {e['deal_id'] for e in confirmed
                 if e['type'] in ('first_disclosed', 'first_draft') and ma(e)}
    reg_ids = {e['deal_id'] for e in confirmed if e['type'] == 'registered' and ma(e)}
    term_ids = {e['deal_id'] for e in confirmed if e['type'] == 'terminated' and ma(e)}
    passed = [e for e in confirmed if e['type'] in ('committee_passed', 'simplified_passed') and ma(e)]
    passed_ids = {e['deal_id'] for e in passed}
    accepted_ids = {e['deal_id'] for e in confirmed if e['type'] == 'accepted' and ma(e)}
    meetings = len({e.get('meeting_id') for e in passed if e.get('meeting_id')})
    by_board = Counter(cases.get(did, {}).get('board', '') for did in first_ids)
    # In-review stock bridge: a new acceptance adds one; a passed/registered
    # transaction leaves the in-review pool. Deduplicate by deal id so a same-day
    # pass plus registration submission is not subtracted twice.
    in_review_delta = len(accepted_ids - passed_ids - reg_ids) - len(passed_ids - accepted_ids)
    return {
        'first_disclosed': len(first_ids),
        'registered': len(reg_ids),
        'meetings': meetings,
        'in_review': in_review_delta,
        'by_board': dict(by_board),
        'event_total': len(confirmed),
        'committee_passed': len(passed_ids),
        'terminated': len(term_ids),
    }


def bridge_replacements(files: list[Path], ledger: dict) -> tuple[dict[str, str], dict]:
    """从上期 Word/PPT 读累计基数，按本期事件增量桥接出生成本期数字。"""
    text = joined_text(files)
    dl = deltas(ledger)
    repl: dict[str, str] = {}
    report: dict = {'deltas': dl, 'applied': [], 'skipped': []}
    for key, pat in BRIDGE_PATTERNS:
        m = re.search(pat, text)
        if not m:
            report['skipped'].append({'key': key, 'reason': '上期正文未找到该统计句式'})
            continue
        old, base = m.group(0), int(m.group(1))
        new_val = max(0, base + dl.get(key, 0))
        new = old.replace(str(base), str(new_val), 1)
        if new != old:
            repl[old] = new
            report['applied'].append({'key': key, 'old': old, 'new': new,
                                      'base': base, 'delta': dl.get(key, 0)})
    return repl, report


def header_map(path: Path, sheet: str, scan_rows: int = 3) -> dict[str, str]:
    """只读 openpyxl 找真实表头行，返回 {语义: 列字母}。不保存文件。"""
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    if sheet not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[sheet]
    best: dict[str, str] = {}
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        row = {}
        for c in range(1, min(ws.max_column, 80) + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = re.sub(r'\s+', '', str(v).strip())
            for key, aliases in COLUMN_ALIASES.items():
                if key not in row and any(re.sub(r'\s+', '', a) in s for a in aliases):
                    row[key] = openpyxl.utils.get_column_letter(c)
        if len(row) > len(best):
            best = row
    wb.close()
    return best


def last_data_row(path: Path, sheet: str) -> int:
    wb = openpyxl.load_workbook(path, read_only=False, data_only=False)
    if sheet not in wb.sheetnames:
        wb.close()
        return 0
    ws = wb[sheet]
    rows = [c.row for c in ws._cells.values() if c.value is not None]
    wb.close()
    return max(rows) if rows else 0


def first_sheet(path: Path) -> str:
    wb = openpyxl.load_workbook(path, read_only=True)
    name = wb.sheetnames[0]
    wb.close()
    return name


def _event_source(events: list[dict], sources: dict[str, dict], deal_id: str) -> dict:
    ids = next((e.get('source_ids', []) for e in events if e['deal_id'] == deal_id and e.get('source_ids')), [])
    return sources.get(ids[0], {}) if ids else {}


def _latest_amount(case: dict, *, actual: bool = False) -> float | None:
    vals = []
    for item in case.get('amounts', []):
        if item.get('currency') != 'CNY':
            continue
        if actual and item.get('basis') != 'actual_completion':
            continue
        if not actual and item.get('basis') == 'actual_completion':
            continue
        try:
            vals.append((item.get('known_on', ''), float(item.get('value_wan'))))
        except (TypeError, ValueError):
            continue
    return sorted(vals, key=lambda x: x[0])[-1][1] if vals else None


def _row_values(event: dict, case: dict, source: dict, cols: dict[str, str]) -> dict:
    facts = case.get('facts', {})
    row: dict[str, str | float] = {}
    clean_locator = re.sub(r'^巨潮公告｜[^｜]+｜', '', source.get('locator', ''))
    summary_value = facts.get('summary') or clean_locator or event.get('note', '')
    if event.get('type') == 'control_completed':
        pieces = []
        if facts.get('old_controller') or facts.get('new_controller'):
            pieces.append('实际控制人由%s变更为%s' %
                          (facts.get('old_controller') or '原实际控制人',
                           facts.get('new_controller') or '新实际控制人'))
        pieces.append('控制权变更已完成')
        if facts.get('completion_date'):
            pieces.append('完成日期为%s' % facts['completion_date'])
        actual_amount = _latest_amount(case, actual=True)
        if actual_amount is not None:
            pieces.append('相关发行或交易金额为%s万元' % actual_amount)
        summary_value = '；'.join(pieces) + '。'
    vals: dict[str, str | float | None] = {
        'name': case.get('name', ''), 'code': case.get('code', ''),
        'date': event.get('date', ''), 'latest': event.get('announced_on') or event.get('date', ''),
        'completion': fact_or(event, facts, 'completion_date'),
        'board': case.get('board', ''), 'title': clean_locator or event.get('note', ''),
        'target': facts.get('target', ''), 'method': facts.get('payment_method', ''),
        'amount': _latest_amount(case, actual=event.get('type') in ('control_completed', 'completed')),
        'support': facts.get('supporting_financing_wan'),
        'advisor': '、'.join(a.get('name', '') for a in case.get('advisors', []) if a.get('role') == 'independent'),
        'summary': summary_value,
        'progress': {'accepted':'已受理','inquiry':'已问询','replied':'已回复','resumed':'已恢复审核',
                     'committee_scheduled':'已安排会议','committee_passed':'通过重组委',
                     'registration_submitted':'提交注册','registered':'注册生效',
                     'terminated':'终止','completed':'已完成','control_completed':'已完成',
                     'control_progress':'进行中','first_disclosed':'进行中','first_draft':'进行中','lead_signal':'筹划中'}.get(event.get('type'), ''),
        'in_review': '否' if event.get('type') in ('committee_passed','registration_submitted','registered','terminated','completed','control_completed') else ('是' if event.get('type') in ('accepted','inquiry','replied','suspended','resumed') else ''),
        'old_controller': facts.get('old_controller', ''), 'new_controller': facts.get('new_controller', ''),
        'control_ratio': _ratio(facts.get('control_ratio_pct')),
        'control_ratio_before': _ratio(facts.get('controller_ratio_before_pct')),
        'transfer_ratio': _ratio(facts.get('transfer_ratio_pct')),
        'price_per_share': facts.get('price_per_share', ''),
        'agreement_date': facts.get('agreement_date', ''),
        'lockup': facts.get('lockup', ''),
        'related': facts.get('related_party', ''), 'major_reorg': facts.get('major_reorg', ''),
        'reverse_merger': facts.get('reverse_merger', ''), 'termination_stage': facts.get('termination_stage', ''),
        'termination_reason': facts.get('termination_reason', ''),
        'note': (('暂未公告预案；' if event.get('type') == 'lead_signal' else '') +
                 (f"官方公告：{clean_locator}" if clean_locator else '')),
    }
    for key, value in vals.items():
        if key in cols and value not in (None, ''):
            if key == 'name':
                # PDF 抽取会在 CJK 之间插空格：'英 力 特' -> '英力特'。
                value = celltext.normalize_name(value)
            elif isinstance(value, str):
                # v4.3.8：证据层与展示层分离。facts['summary'] 是前两页
                # 公告原文（业务口径分类依赖它），绝不能原样写进单元格；
                # 这里统一净化为有界摘要，兜底退回公告标题。
                value = celltext.safe_cell_value(value, fallback=clean_locator)
            row[cols[key]] = value
    return row


def fact_or(event: dict, facts: dict, key: str):
    return facts.get(key, '')


def _ratio(value):
    try:
        return float(value) / 100
    except (TypeError, ValueError):
        return None


def last_seq_value(path: Path, sheet: str, col: str, row_no: int) -> int:
    """读取上期最后一条数据行的序号值。

    追加行绝不能写 '=A149+1' 这类公式字符串：OOXML 局部补丁按 inlineStr
    写单元格，Excel 会把它显示成普通文字，序号列从此失去数值语义，也无法
    参与排序/筛选。直接写数值序号，不依赖 Excel 公式重算即可保真。
    """
    from openpyxl.utils.cell import column_index_from_string
    try:
        idx = column_index_from_string(col)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[sheet]
            for cells in ws.iter_rows(min_row=row_no, max_row=row_no,
                                      min_col=idx, max_col=idx):
                v = cells[0].value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    return int(v)
        finally:
            wb.close()
    except Exception:  # noqa: BLE001 - 读不到就退回行号，宁可近似也不错写文本
        pass
    return row_no


def _spec(path: Path, sheet: str, selected: list[dict], ledger: dict, max_rows: int,
          control_history: dict[str, dict[str, object]] | None = None,
          workbook_history: dict[str, dict[str, object]] | None = None) -> dict | None:
    cols = header_map(path, sheet)
    if not cols.get('name') and not cols.get('code'):
        return None
    cases = {c['deal_id']: c for c in ledger.get('cases', [])}
    sources = {s['source_id']: s for s in ledger.get('sources', [])}
    rows = []
    for event in selected:
        case = cases.get(event['deal_id'])
        if not case:
            continue
        row = _row_values(event, case, _event_source(selected, sources, event['deal_id']), cols)
        code = code_digits(case.get('code', ''))
        if path.name.startswith('7、') and sheet == '2026年已完成':
            base = mirrored_row_values(path, '2026年首次披露', sheet, case.get('code', ''))
            if not base and control_history:
                base = map_headers_to_sheet(path, sheet, control_history.get(code, {}))
            overlay = {cols[key] for key in
                       ('latest', 'completion', 'progress', 'note', 'summary', 'amount',
                        'control_ratio_before', 'control_ratio') if key in cols}
            for col, value in row.items():
                if col in overlay or col not in base:
                    base[col] = value
            facts = case.get('facts', {})
            for key in ('old_controller', 'new_controller'):
                col = cols.get(key)
                current = facts.get(key, '')
                if col and current and len(str(current)) > len(str(base.get(col, ''))):
                    base[col] = current
            row = base
        elif workbook_history and code in workbook_history:
            base = map_headers_to_sheet(path, sheet, workbook_history[code])
            overlay = {cols[key] for key in
                       ('date', 'latest', 'completion', 'progress', 'in_review',
                        'termination_stage', 'termination_reason', 'title', 'note') if key in cols}
            for col, value in row.items():
                if col in overlay or col not in base:
                    base[col] = value
            row = base
        if row:
            rows.append(row)
        if len(rows) >= max_rows:
            break
    if not rows:
        return None
    tpl = last_data_row(path, sheet)
    if tpl < 2:
        return None
    if cols.get('seq'):
        base = last_seq_value(path, sheet, cols['seq'], tpl)
        for idx, row in enumerate(rows):
            row[cols['seq']] = base + idx + 1
    return {'sheet': sheet, 'columns': cols, 'template_row': tpl, 'rows': rows, 'count': len(rows)}


def code_digits(value: str) -> str:
    return re.sub(r'\D', '', str(value or ''))[:6]


def sheet_codes(path: Path, sheet: str) -> dict[str, int]:
    """Return security code -> first matching row in one worksheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    if sheet not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[sheet]
    found: dict[str, int] = {}
    try:
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 5000),
                                min_col=1, max_col=min(ws.max_column, 80)):
            for cell in row:
                value = cell.value
                if isinstance(value, str):
                    m = re.fullmatch(r'\s*(\d{6})(?:\.(?:SH|SZ|BJ))?\s*', value, re.I)
                    if m:
                        found.setdefault(m.group(1), cell.row)
                        break
                elif isinstance(value, int) and 100000 <= value <= 999999:
                    found.setdefault(str(value), cell.row)
                    break
    finally:
        wb.close()
    return found


def _normalized_headers(ws) -> dict[str, int]:
    return {re.sub(r'\s+', '', str(ws.cell(1, c).value or '')): c
            for c in range(1, min(ws.max_column, 100) + 1)
            if ws.cell(1, c).value is not None}


def row_values_by_header(path: Path, sheet: str, code: str) -> dict[str, object]:
    wanted = code_digits(code)
    wb = openpyxl.load_workbook(path, read_only=False, data_only=False)
    if sheet not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[sheet]
    headers = _normalized_headers(ws)
    source_row = None
    for r in range(2, min(ws.max_row, 5000) + 1):
        for c in headers.values():
            if code_digits(ws.cell(r, c).value) == wanted:
                source_row = r
                break
        if source_row:
            break
    if not source_row:
        wb.close()
        return {}
    out = {name: ws.cell(source_row, c).value for name, c in headers.items()}
    wb.close()
    return out


def map_headers_to_sheet(path: Path, sheet: str, source_values: dict[str, object]) -> dict[str, object]:
    wb = openpyxl.load_workbook(path, read_only=False, data_only=False)
    if sheet not in wb.sheetnames:
        wb.close()
        return {}
    headers = _normalized_headers(wb[sheet])
    wb.close()
    out = {}
    for name, target_col in headers.items():
        if name == '序号':
            continue
        value = source_values.get(name)
        if value not in (None, ''):
            out[openpyxl.utils.get_column_letter(target_col)] = value
    return out


def mirrored_row_values(path: Path, source_sheet: str, target_sheet: str, code: str) -> dict[str, object]:
    """Carry a historical transaction's shared fields into a derived table."""
    return map_headers_to_sheet(path, target_sheet,
                                row_values_by_header(path, source_sheet, code))


def workbook_code_index(path: Path) -> dict[str, dict[str, object]]:
    """Index the most complete row for each issuer across one workbook.

    A later milestone can move from a first-disclosure sheet into a review /
    committee sheet.  Shared columns should follow the issuer even when the
    target sheet has no existing row yet.
    """
    index: dict[str, dict[str, object]] = {}
    wb = openpyxl.load_workbook(path, read_only=False, data_only=False)
    for ws in wb.worksheets:
        headers = _normalized_headers(ws)
        code_cols = [c for name, c in headers.items() if name in {'代码', '股票代码', '证券代码'}]
        if not code_cols:
            continue
        for r in range(2, min(ws.max_row, 5000) + 1):
            code = next((code_digits(ws.cell(r, c).value) for c in code_cols
                         if code_digits(ws.cell(r, c).value)), '')
            if not code:
                continue
            values = {name: ws.cell(r, c).value for name, c in headers.items()
                      if ws.cell(r, c).value not in (None, '')}
            if len(values) > len(index.get(code, {})):
                index[code] = values
    wb.close()
    return index


def control_first_disclosure_index(files: list[Path]) -> dict[str, dict[str, object]]:
    """Index the newest first-disclosure row across 2026/2025/2024 workbooks."""
    index: dict[str, dict[str, object]] = {}
    ranked = sorted(files, key=lambda p: int((re.match(r'^(\d+)、', p.name) or [None, '99'])[1]))
    # 7 is the current-year register; then 8 (prior year), then 9 (older year).
    for path in ranked:
        m = re.match(r'^(\d+)、', path.name)
        if not m or m.group(1) not in {'7', '8', '9'}:
            continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        sheets = [s for s in wb.sheetnames if s.endswith('首次披露')]
        wb.close()
        for sheet in sheets:
            wb = openpyxl.load_workbook(path, read_only=False, data_only=False)
            ws = wb[sheet]
            headers = _normalized_headers(ws)
            code_cols = [c for name, c in headers.items() if name in {'代码', '股票代码', '证券代码'}]
            for r in range(2, min(ws.max_row, 5000) + 1):
                code = next((code_digits(ws.cell(r, c).value) for c in code_cols
                             if code_digits(ws.cell(r, c).value)), '')
                if not code:
                    continue
                index.setdefault(code, {
                    name: ws.cell(r, c).value for name, c in headers.items()
                    if ws.cell(r, c).value not in (None, '')})
            wb.close()
    return index


def sheet_name_rows(path: Path, sheet: str) -> dict[str, tuple[str, int]]:
    """Map normalized company name to existing code/row for code-migration cases."""
    wb = openpyxl.load_workbook(path, read_only=False, data_only=False)
    if sheet not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[sheet]
    header_row, code_cols, name_cols = 1, [], []
    for r in range(1, min(4, ws.max_row) + 1):
        cs, ns = [], []
        for c in range(1, min(ws.max_column, 100) + 1):
            h = re.sub(r'\s+', '', str(ws.cell(r, c).value or ''))
            if '代码' in h:
                cs.append(c)
            if any(k in h for k in ('简称', '公司名称', '证券名称')):
                ns.append(c)
        if cs and ns:
            header_row, code_cols, name_cols = r, cs, ns
            break
    out = {}
    for r in range(header_row + 1, min(ws.max_row, 5000) + 1):
        code = next((code_digits(ws.cell(r, c).value) for c in code_cols), '')
        name = next((re.sub(r'\s+', '', str(ws.cell(r, c).value or '')) for c in name_cols), '')
        if code and name:
            out.setdefault(name, (code, r))
    wb.close()
    return out


def unique_events(events: list[dict], preferred_types: set[str] | None = None) -> list[dict]:
    """Keep one row per deal, preferring the latest/highest-value state in a week."""
    rank = {'control_progress': 15, 'first_disclosed': 10, 'first_draft': 20, 'accepted': 30, 'inquiry': 40,
            'replied': 41, 'suspended': 45, 'resumed': 46, 'committee_scheduled': 50,
            'committee_passed': 60, 'simplified_passed': 61, 'registration_submitted': 70,
            'registered': 80, 'terminated': 90, 'completed': 95, 'control_completed': 96}
    best: dict[str, dict] = {}
    for event in events:
        if preferred_types and event.get('type') not in preferred_types:
            continue
        key = event['deal_id']
        old = best.get(key)
        score = (event.get('date', ''), event.get('announced_on', ''), rank.get(event.get('type'), 0))
        old_score = (old.get('date', ''), old.get('announced_on', ''), rank.get(old.get('type'), 0)) if old else ('', '', -1)
        if score >= old_score:
            best[key] = event
    return sorted(best.values(), key=lambda e: (e.get('date', ''), e['deal_id']))


def update_cells_for_existing(path: Path, sheet: str, selected: list[dict],
                              ledger: dict, *, preserve_first_date: bool = False) -> dict[str, str]:
    """Build minimal cell updates for codes already present in a sheet."""
    cols = header_map(path, sheet)
    row_map = sheet_codes(path, sheet)
    name_rows = sheet_name_rows(path, sheet)
    cases = {c['deal_id']: c for c in ledger.get('cases', [])}
    sources = {s['source_id']: s for s in ledger.get('sources', [])}
    updates: dict[str, str] = {}
    safe_keys = {'name', 'code', 'board', 'progress', 'in_review', 'latest', 'advisor',
                 'target', 'method', 'amount', 'support', 'summary', 'related',
                 'major_reorg', 'reverse_merger', 'old_controller', 'new_controller',
                 'control_ratio', 'control_ratio_before', 'transfer_ratio', 'price_per_share',
                 'agreement_date', 'lockup', 'completion', 'termination_stage',
                 'termination_reason', 'note'}
    state_keys = {'progress', 'in_review', 'latest', 'completion',
                  'termination_stage', 'termination_reason'}
    wb_existing = openpyxl.load_workbook(path, read_only=False, data_only=False)
    ws_existing = wb_existing[sheet]
    for event in unique_events(selected):
        case = cases.get(event['deal_id'])
        if not case:
            continue
        code = code_digits(case.get('code', ''))
        name_key = re.sub(r'\s+', '', case.get('name', ''))
        row_no = row_map.get(code) or (name_rows.get(name_key, (None, None))[1])
        if not row_no:
            continue
        vals = _row_values(event, case, _event_source(selected, sources, event['deal_id']), cols)
        for key in safe_keys:
            col = cols.get(key)
            if not col or col not in vals or vals[col] in (None, ''):
                continue
            current = ws_existing[f'{col}{row_no}'].value
            if key == 'note' and current not in (None, '') and sheet.endswith('在审'):
                appended_note = str(vals[col])
                if appended_note not in str(current):
                    updates[f'{col}{row_no}'] = f'{current}；{appended_note}'
            elif key in state_keys or current in (None, ''):
                updates[f'{col}{row_no}'] = vals[col]
        if not preserve_first_date and 'date' in cols and cols['date'] in vals:
            updates[f"{cols['date']}{row_no}"] = event.get('announced_on') or event.get('date', '')
    wb_existing.close()
    return updates


def merge_sheet_patches(store: dict[str, dict[str, str]], sheet: str,
                        patches: dict[str, str]) -> None:
    """合并同一 sheet 的单元格补丁。

    in_review_sheets 把多个板块映射到同一张在审表（沪主板/科创板 → 上交所在审），
    excel_plan 会带不同 selected 多次进入该 sheet；若直接赋值，后一批次会把
    前一批次的坐标补丁全部冲掉（v4.3.6 实测：9 条沪主板中止状态回写丢失）。
    """
    if patches:
        store.setdefault(sheet, {}).update(patches)


def workbook_codes(path: Path) -> set[str]:
    """Collect security codes already represented anywhere in a workbook.

    A first-disclosure event is an append only when the issuer/transaction code
    is absent. This prevents a new draft/progress announcement for an existing
    transaction (e.g. 2026-W43 节能铁汉) from being duplicated as a new deal.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    found: set[str] = set()
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 2000),
                                    min_col=1, max_col=min(ws.max_column, 80),
                                    values_only=True):
                for value in row:
                    if isinstance(value, str):
                        m = re.fullmatch(r'\s*(\d{6})(?:\.(?:SH|SZ|BJ))?\s*', value, re.I)
                        if m:
                            found.add(m.group(1))
    finally:
        wb.close()
    return found

def workbook_role(path: Path) -> str:
    """Classify workbook function from its real filename, not just its number."""
    m = re.match(r'^(\d+)、', path.name)
    number = m.group(1) if m else ''
    name = path.name
    if '上市公司子公司引战及与专业机构合作' in name:
        return 'invest'
    if 'IPO终止' in name:
        return 'ipo'
    if '控股股东、实际控制人协议转让部分股份' in name:
        return 'partial'
    if '协议受让成为上市公司二股东' in name:
        return 'second'
    return {'0': 'base', '4': 'review_raw', '7': 'control2026',
            '8': 'control2025', '9': 'control2024'}.get(number, f'xlsx_{number}')


def excel_plan(files: list[Path], ledger: dict, max_rows: int = 50) -> dict[str, dict]:
    """Map confirmed events to appends and updates using per-sheet identities."""
    cases = {c['deal_id']: c for c in ledger.get('cases', [])}
    control_history = control_first_disclosure_index(files)
    all_events = ledger.get('events', [])
    confirmed = [e for e in all_events if e.get('status') == 'CONFIRMED']
    cash_leads = [e for e in all_events if e.get('type') == 'lead_signal'
                  and e.get('status') in {'CONFIRMED', 'CANDIDATE'}
                  and cases.get(e['deal_id'], {}).get('kind') in {'ma_cash_purchase', 'ma_cash_sale'}]
    events = [e for e in confirmed if e.get('type') != 'lead_signal']
    mapped_events = events + cash_leads

    def pick(kind, types):
        kinds = {kind} if isinstance(kind, str) else set(kind)
        return [e for e in mapped_events
                if ({cases.get(e['deal_id'], {}).get('kind')} | set(cases.get(e['deal_id'], {}).get('also_kinds', []))) & kinds
                and e.get('type') in types]

    in_review_sheets = {'沪主板': '上交所在审', '科创板': '上交所在审',
                        '深主板': '深交所在审', '北交所': '北交所在审'}
    by_file: dict[str, list[tuple[str, list[dict]]]] = {
        'base': [
            ('现金重大购买', pick('ma_cash_purchase', {'first_disclosed', 'first_draft', 'lead_signal'})),
            ('现金重大出售', pick('ma_cash_sale', {'first_disclosed', 'first_draft', 'lead_signal'})),
            ('重组委通过', pick('ma_audit', {'committee_passed', 'simplified_passed'})),
            ('证监会注册', pick('ma_audit', {'registered'})),
            ('撤回终止', pick({'ma_audit', 'ma_cash_purchase', 'ma_cash_sale'}, {'terminated'})),
        ],
        'review_raw': [('重组委审核原始数据', pick('ma_audit', {'committee_passed', 'simplified_passed', 'committee_deferred', 'committee_rejected'}))],
        'control2026': [('2026年首次披露', pick('control', {'first_disclosed', 'control_progress', 'control_completed'})),
              ('2026年已完成', pick('control', {'control_completed'}))],
        'ipo': [('2026', pick('ipo', {'terminated'}))],
        'invest': [('上市公司子公司引战情况', pick('strategic_investment', {'first_disclosed'})),
               ('上市公司与专业机构合作情况', pick('professional_cooperation', {'first_disclosed'}))],
        'partial': [('2026年首次披露（2026.9.13起统计）', pick('partial_transfer', {'first_disclosed'}))],
        'second': [('2026年首次披露（2026.9.13起统计）', pick('second_shareholder', {'first_disclosed'}))],
    }
    # The 34–37 era split first disclosures into Shanghai/Shenzhen and Beijing
    # sheets; later issues merged them.  Resolve the actual layout from the template.
    base_file = next((p for p in files if workbook_role(p) == 'base'), None)
    if base_file:
        probe = openpyxl.load_workbook(base_file, read_only=True)
        base_sheets = set(probe.sheetnames)
        probe.close()
        first_all = pick('ma_audit', {'first_disclosed', 'first_draft'})
        if '沪深北交易所首次披露' in base_sheets:
            by_file['base'].insert(0, ('沪深北交易所首次披露', first_all))
        else:
            if '沪深交易所首次披露' in base_sheets:
                by_file['base'].insert(0, ('沪深交易所首次披露',
                    [e for e in first_all if cases[e['deal_id']].get('board') != '北交所']))
            if '北交所首次披露' in base_sheets:
                by_file['base'].insert(1, ('北交所首次披露',
                    [e for e in first_all if cases[e['deal_id']].get('board') == '北交所']))

    # A control event may complete in the current year but first-disclose in a
    # prior-year register.  Resolve the actual historical sheet names rather
    # than assuming fixed labels, and update those rows in place.
    control_first_events = pick('control', {'first_disclosed', 'control_progress', 'control_completed'})
    for path in files:
        role = workbook_role(path)
        if role not in {'control2025', 'control2024'}:
            continue
        probe = openpyxl.load_workbook(path, read_only=True)
        first_sheets = [name for name in probe.sheetnames if name.endswith('首次披露')]
        probe.close()
        for first_sheet in first_sheets:
            by_file.setdefault(role, []).append((first_sheet, control_first_events))

    # Milestone status is applied back to the relevant in-review sheet.
    state_events = pick('ma_audit', {'accepted', 'inquiry', 'replied', 'suspended', 'resumed',
                                     'committee_passed', 'simplified_passed',
                                     'registration_submitted', 'registered', 'terminated'})
    for board, sheet in in_review_sheets.items():
        by_file.setdefault('base', []).append((sheet, [e for e in state_events if cases[e['deal_id']].get('board') == board]))

    plan: dict[str, dict] = {}
    for path in files:
        if path.suffix.lower() not in ('.xlsx', '.xlsm'):
            continue
        role = workbook_role(path)
        if role not in by_file:
            continue
        specs: list[dict] = []
        updates: dict[str, dict[str, str]] = {}
        workbook_history = workbook_code_index(path) if role in {'base', 'review_raw', 'control2026', 'ipo'} else {}
        for sheet, selected in by_file[role]:
            if not selected:
                continue
            probe = openpyxl.load_workbook(path, read_only=True)
            has_sheet = sheet in probe.sheetnames
            probe.close()
            if not has_sheet:
                continue
            present = sheet_codes(path, sheet)
            name_rows = sheet_name_rows(path, sheet)
            unique = unique_events(selected)
            def absent(case):
                return (code_digits(case.get('code', '')) not in present
                        and re.sub(r'\s+', '', case.get('name', '')) not in name_rows)
            # Investment/cooperation announcements are event-level, not a
            # perpetual one-row-per-issuer register: one issuer may launch two
            # different funds in the same period.  Other sheets are deal/state
            # registers where an existing code identifies the row to update.
            transaction_level = role == 'invest'
            control_first = role in {'control2026', 'control2025', 'control2024'} and sheet.endswith('首次披露')
            if control_first:
                first_events = unique_events([e for e in selected
                                              if e.get('type') in {'first_disclosed', 'first_draft'}])
                appended = ([e for e in first_events
                             if absent(cases[e['deal_id']])]
                            if role == 'control2026' else [])
            elif sheet.endswith('在审'):
                terminal = {'committee_passed', 'committee_rejected', 'simplified_passed',
                            'registration_submitted', 'registered', 'terminated', 'completed'}
                terminal_deals = {e['deal_id'] for e in unique if e.get('type') in terminal}
                appended = [e for e in unique if e.get('type') not in terminal
                            and e['deal_id'] not in terminal_deals
                            and absent(cases[e['deal_id']])]
            elif role in {'partial', 'second'}:
                appended = unique
            else:
                appended = unique if transaction_level else [
                    e for e in unique if absent(cases[e['deal_id']])]
            # Existing deals stay in the same row. Requested patches only cover
            # columns actually present; no guess is made for historical columns.
            cell_updates = {} if transaction_level else update_cells_for_existing(
                path, sheet, unique, ledger,
                preserve_first_date=sheet.endswith('首次披露') or sheet in {'沪深北交易所首次披露',
                                                                         '现金重大购买', '现金重大出售'})
            if cell_updates:
                # in_review_sheets 把多个板块映射到同一张在审表（沪主板/科创板
                # → 上交所在审），本循环会带着不同 selected 多次进入同一 sheet。
                # 必须合并坐标补丁；直接赋值会让后一批次把前一批次全部冲掉。
                merge_sheet_patches(updates, sheet, cell_updates)
            spec = _spec(path, sheet, appended, ledger, max_rows, control_history, workbook_history)
            if spec:
                prior = next((s for s in specs if s.get('sheet') == sheet), None)
                if prior is None:
                    specs.append(spec)
                else:
                    prior['rows'] = list(prior.get('rows') or []) + list(spec.get('rows') or [])
                    prior['count'] = len(prior['rows'])
        if specs or updates:
            plan[path.name] = {'sheets': specs, 'updates': updates,
                               'count': sum(x['count'] for x in specs),
                               'update_cells': sum(len(v) for v in updates.values())}
    return plan

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--prev', required=True, type=Path)
    ap.add_argument('--out', required=True, type=Path)
    ap.add_argument('--period', required=True, type=Path)
    ap.add_argument('--ledger', required=True, type=Path)
    ap.add_argument('--report', required=True, type=Path)
    ap.add_argument('--max-rows', type=int, default=200)
    args = ap.parse_args()

    pd = json.loads(args.period.read_text(encoding='utf-8'))
    period = Period(**pd)
    ledger = json.loads(args.ledger.read_text(encoding='utf-8'))
    files = prev_native(args.prev)
    if not files:
        raise RuntimeError(f'上期目录无原生文件: {args.prev}')

    base_repl = text_replacements(period)
    bridge_repl, bridge_report = bridge_replacements(files, ledger)
    all_repl = {**base_repl, **bridge_repl}
    plan = excel_plan(files, ledger, args.max_rows)
    prefix_counts = Counter((re.match(r'^(\d+)、', p.name) or [None, ''])[1]
                            for p in files if p.suffix.lower() in ('.xlsx', '.xlsm'))
    normalize_roles = any(n > 1 for k, n in prefix_counts.items() if k)
    role_numbers = {'ipo': 10, 'invest': 11, 'partial': 12, 'second': 13}

    args.out.mkdir(parents=True, exist_ok=True)
    report: dict = {'prev_dir': str(args.prev), 'out': str(args.out),
                    'period': {'start': period.start, 'end': period.end, 'issue_no': period.issue_no},
                    'base_replacements': base_repl, 'bridge': bridge_report,
                    'excel_plan': {k: {'sheets': [{
                        kk: vv for kk, vv in spec.items() if kk != 'rows'
                    } for spec in v['sheets']], 'updates': {sn: sorted(cells) for sn, cells in v.get('updates', {}).items()},
                    'count': v['count'], 'update_cells': v.get('update_cells', 0)} for k, v in plan.items()},
                    'files': [], 'unmatched_text': []}
    for src in files:
        dest_name = map_filename(src.name, period)
        role = workbook_role(src)
        if normalize_roles and role in role_numbers:
            dest_name = re.sub(r'^\d+、', f'{role_numbers[role]}、', dest_name, count=1)
        dst = args.out / dest_name
        if src.suffix.lower() in ('.xlsx', '.xlsm') and src.name in plan:
            specs = plan[src.name]['sheets']
            r = ooxml.patch_workbook(
                src, dst, all_repl,
                sheet_updates=plan[src.name].get('updates'),
                sheet_appends={spec['sheet']: (spec['template_row'], spec['rows']) for spec in specs})
        else:
            r = {'text_hits': ooxml.patch_text(src, dst, all_repl)}
        hits = {k: v for k, v in r.get('text_hits', {}).items() if v}
        entry = {'source': src.name, 'dest': dst.name, 'bytes': dst.stat().st_size,
                 'text_hits': hits, 'cells': r.get('cells', {}),
                 'appended': r.get('appended', {}), 'green_xf': r.get('green_xf')}
        report['files'].append(entry)
        for k in all_repl:
            if not hits.get(k):
                report['unmatched_text'].append({'file': src.name, 'pattern': k})

    # v4.3.8 硬门禁：成品不得含「整段公告正文塞进单元格」。
    # 现场人工稿中偶有引用证券代码的长摘要（只命中 1 个抬头特征），
    # is_raw_filing_body 要求 >=2 个特征，不会误伤；第 40 期现场稿实测 0 命中。
    generated_xlsx = sorted(p for p in args.out.glob('*.xlsx')
                            if not p.name.startswith('~$'))
    violations = celltext.scan_raw_filing_cells(generated_xlsx)
    report['cell_text_gate'] = {
        'version': '4.3.8', 'scanned_files': len(generated_xlsx),
        'violations': violations, 'status': 'PASS' if not violations else 'BLOCKED',
        'rule': '单元格文本长度>200 且前80字命中>=2个公告抬头特征即阻断',
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'native_files': len(report['files']),
                      'excel_with_new_rows': len(plan),
                      'new_rows': sum(v['count'] for v in plan.values()),
                      'updated_cells': sum(v.get('update_cells', 0) for v in plan.values()),
                      'bridge_applied': len(bridge_report['applied']),
                      'cell_text_gate': report['cell_text_gate']['status'],
                      'cell_text_violations': len(violations),
                      'out': str(args.out)}, ensure_ascii=False))
    if violations:
        print(json.dumps({'error': 'CELL_TEXT_GATE_BLOCKED',
                          'detail': '成品单元格含整段公告正文，禁止出件',
                          'violations': violations[:8]}, ensure_ascii=False))
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
