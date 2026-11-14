#!/usr/bin/env python3
"""期次推导。唯一人工输入是 as-of 日期，其余全部由最新完整上期成果推导。

现场约定（由第41/42/43/44期目录名实测得到）：
  - 统计窗口为周日至周六，本期起点 = 上期结束日 + 1 天，终点 = as-of；
  - 当年期号 +1；历史年度库的自身期号同样 +1（如 55->56、119->120、151->152）；
  - 文件名中的日期串统一替换为 <start>至<end>；
  - 输出目录名 = 并购周报基础数据-{年份}年第{N}期-{start}至{end}。
跨年初始化属于口径变更，脚本继续出件但在 plan 中标记需确认，绝不静默假定。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from pathlib import Path

DIR_RE = re.compile(r'并购周报基础数据-(?P<year>\d{4})年第(?P<no>\d+)期-(?P<start>\d{8})至(?P<end>\d{8})')
DATE_SPAN_RE = re.compile(r'(?P<start>\d{8})至(?P<end>\d{8,})')
ISSUE_RE = re.compile(r'(?P<year>\d{4})年第(?P<no>\d+)期')
NATIVE_SUFFIXES = {'.xlsx', '.xlsm', '.docx', '.pptx'}


def d8(d: date) -> str:
    return d.strftime('%Y%m%d')


def parse_d8(s: str) -> date:
    return datetime.strptime(s, '%Y%m%d').date()


@dataclass
class Period:
    as_of: str
    start: str
    end: str
    year: int
    issue_no: int
    prev_dir: str
    prev_year: int
    prev_issue_no: int
    prev_start: str
    prev_end: str
    out_dirname: str
    notes: list[str]

    @property
    def span_token(self) -> str:
        return f'{self.start}至{self.end}'

    @property
    def prev_span_token(self) -> str:
        return f'{self.prev_start}至{self.prev_end}'


def find_latest_complete(root: Path, before: date) -> tuple[Path, dict]:
    """扫描工作区，取结束日不晚于 as-of 且包含完整角色集的最新一期目录。"""
    candidates = []
    for p in root.iterdir():
        if not p.is_dir():
            continue
        m = DIR_RE.search(p.name)
        if not m:
            continue
        try:
            end = parse_d8(m.group('end'))
        except ValueError:
            continue
        if end >= before:
            continue
        files = [f for f in p.iterdir() if f.is_file() and f.suffix.lower() in NATIVE_SUFFIXES
                 and not f.name.startswith('~$')]
        roles = {_role(f.name) for f in files}
        roles.discard(None)
        required = {f'xlsx_{i}' for i in range(14) if i != 5} | {'weekly_docx', 'dynamic_docx', 'market_ppt'}
        missing = sorted(required - roles)
        candidates.append((end, p, {'year': int(m.group('year')), 'no': int(m.group('no')),
                                    'start': m.group('start'), 'end': m.group('end'),
                                    'missing': missing, 'complete': not missing}))
    if not candidates:
        raise RuntimeError(f'在 {root} 未找到 as-of 之前的期次目录')
    # The immediately preceding period defines the next window even when it is
    # incomplete.  Choosing an older "complete" package silently skips a week.
    latest_end = max(c[0] for c in candidates)
    latest = [c for c in candidates if c[0] == latest_end]
    end, path, meta = next((c for c in latest if c[2]['complete']), max(latest, key=lambda c: c[0]))
    return path, meta


def _role(name: str) -> str | None:
    low = name.lower()
    if name.startswith('并购重组市场周报') and low.endswith('.pptx'):
        return 'market_ppt'
    if name.startswith('A股市场并购重组动态') and low.endswith('.docx'):
        return 'dynamic_docx'
    m = re.match(r'^(\d+)、', name)
    if m:
        if low.endswith(('.xlsx', '.xlsm')):
            return 'xlsx_' + m.group(1)
        if low.endswith('.docx') and m.group(1) == '5':
            return 'weekly_docx'
    return None


def resolve(root: Path, as_of: str, prev_override: str | None = None) -> Period:
    cutoff = date.fromisoformat(as_of)
    if prev_override:
        prev = Path(prev_override).resolve()
        if not prev.is_dir():
            raise RuntimeError(f'指定上期目录不存在: {prev}')
        m = DIR_RE.search(prev.name)
        if not m:
            raise RuntimeError(f'上期目录名无法解析期号: {prev.name}')
        meta = {'year': int(m.group('year')), 'no': int(m.group('no')),
                'start': m.group('start'), 'end': m.group('end'), 'missing': [], 'complete': True}
    else:
        prev, meta = find_latest_complete(root, cutoff)

    prev_end = parse_d8(meta['end'])
    start = prev_end + timedelta(days=1)
    if start > cutoff:
        raise RuntimeError(f'本期起点 {start} 晚于 as-of {cutoff}，请检查日期')
    notes: list[str] = []
    if not meta['complete']:
        notes.append(f'最新上期缺角色 {meta["missing"]}，已按可得文件出件，需在说明中标注')
    cross_year = cutoff.year != meta['year']
    if cross_year:
        issue_no = 1
        notes.append('跨年：当年期号重置为第1期，历史年度库编号沿用上期，需确认新年度初始化口径')
    else:
        issue_no = meta['no'] + 1
    if cutoff.weekday() != 5:
        notes.append(f'as-of {cutoff} 为周{cutoff.weekday()+1}，非现场惯例的周六；按指定日期闭区间执行')
    out_name = f'并购周报基础数据-{cutoff.year}年第{issue_no}期-{d8(start)}至{d8(cutoff)}'
    return Period(as_of=as_of, start=d8(start), end=d8(cutoff), year=cutoff.year,
                  issue_no=issue_no, prev_dir=str(prev), prev_year=meta['year'],
                  prev_issue_no=meta['no'], prev_start=meta['start'], prev_end=meta['end'],
                  out_dirname=out_name, notes=notes)


def map_filename(name: str, p: Period) -> str:
    """把上期文件名映射为本期文件名。历史年度库自身期号 +1，不统一改成本年期号。"""
    out = name
    out = DATE_SPAN_RE.sub(p.span_token, out)

    def bump_issue(m: re.Match) -> str:
        y, n = int(m.group('year')), int(m.group('no'))
        if y == p.prev_year and n == p.prev_issue_no and y == p.year:
            return f'{y}年第{p.issue_no}期'
        return f'{y}年第{n + 1}期'

    out = ISSUE_RE.sub(bump_issue, out)
    return out


def text_replacements(p: Period) -> dict[str, str]:
    """OOXML 文本层的安全替换表：只匹配完整、明确的本期口径锚点。

    不替换业务正文中的历史日期（如案例解析里的 2026年10月30日），
    也不替换年份数字本身（如独立的 2026）。
    """
    ps, pe = parse_d8(p.prev_start), parse_d8(p.prev_end)
    s, e = parse_d8(p.start), parse_d8(p.end)

    def cn(d: date, pad: bool = False) -> str:
        return (f'{d.year}年{d.month:02d}月{d.day:02d}日' if pad
                else f'{d.year}年{d.month}月{d.day}日')

    r: dict[str, str] = {}

    def put(old: str, new: str):
        """首次写入优先：补零与不补零写法在月/日≥10时会生成同一个键，
        若让补零写法覆盖，就会把「2026年11月1日」擅自改成「2026年11月01日」。
        保留先登记的不补零写法，维持源文档的日期风格。"""
        if old != new and old not in r:
            r[old] = new

    # 文件名/页眉中的紧凑区间与三种常见日期写法
    put(p.prev_span_token, p.span_token)
    put(f'{ps.year}.{ps.month:02d}.{ps.day:02d}-{pe.year}.{pe.month:02d}.{pe.day:02d}',
        f'{s.year}.{s.month:02d}.{s.day:02d}-{e.year}.{e.month:02d}.{e.day:02d}')
    put(f'{ps.year}/{ps.month:02d}/{ps.day:02d}-{pe.year}/{pe.month:02d}/{pe.day:02d}',
        f'{s.year}/{s.month:02d}/{s.day:02d}-{e.year}/{e.month:02d}/{e.day:02d}')
    put(f'{cn(ps)}至{cn(pe)}', f'{cn(s)}至{cn(e)}')
    put(f'{cn(ps, True)}至{cn(pe, True)}', f'{cn(s, True)}至{cn(e, True)}')
    put(f'{cn(ps)}-{cn(pe)}', f'{cn(s)}-{cn(e)}')
    # 截止日口径句
    put(f'截至{cn(pe)}', f'截至{cn(e)}')
    put(f'截至 {cn(pe)}', f'截至 {cn(e)}')
    put(f'截至{cn(pe, True)}', f'截至{cn(e, True)}')
    put(f'截至{pe.month}月{pe.day}日', f'截至{e.month}月{e.day}日')
    put(f'截至{pe.year}/{pe.month:02d}/{pe.day:02d}',
        f'截至{e.year}/{e.month:02d}/{e.day:02d}')
    # 期号：完整写法精确匹配；独立「第N期」用负向后顾避免改动历史年度期号
    put(f'{pe.year}年第{p.prev_issue_no}期', f'{e.year}年第{p.issue_no}期')
    # 正则键不参与「old==new」判定，单独追加
    r[f're:(?<!\\d{{4}}年)第{p.prev_issue_no}期'] = f'第{p.issue_no}期'
    return r


def as_dict(p: Period) -> dict:
    return asdict(p)
