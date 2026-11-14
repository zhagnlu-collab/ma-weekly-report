#!/usr/bin/env python3
"""Excel-only 盲测审计：当期新增事件代码集合 + 汇总区缓存差异。

由 2026-11-08 v4.3.1 五轮回放 Round 34→35 固化而来。原临时分析把
`2026-09-01` 压成 `202609`、把金额 `851081.38` 压成 `851081`，
误报为股票代码；本工具用带交易所后缀优先 + 裸码排除小数/年份的规则消噪。

用法（仅在生成稿冻结哈希、现场稿拆封之后调用）：

    python scripts/audit_excel_blind.py \
      --generated <生成稿目录> --reference <现场稿目录> \
      --start 20260830 --end 20260905 \
      --compare <compare_excel.py 产出的 json> --out <输出 json>

只读两边文件，不修改任何成果，不把现场结果回填生成链。

v4.3.6 新增失败关闭：参考稿/生成稿任一 xlsx 不可读（如 OneDrive 云端占位）
时直接 `BLOCKED` 并返回退出码 2，不产出任何代码集合或汇总区对比结论，
避免把「参考侧空集」误写成「生成稿全部有差异」。
"""
from __future__ import annotations
import argparse, json, re, zipfile
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook

CODE_SUFFIX = re.compile(r'\b([03689]\d{5})\.(?:SH|SZ|BJ|BSE)\b', re.I)
CODE_BARE = re.compile(r'(?<![\d.])([03689]\d{5})(?![\d.])')


def codes_in_row(row: tuple) -> set[str]:
    s = ' | '.join(str(v) for v in row if v is not None)
    found = set(m.upper() for m in CODE_SUFFIX.findall(s))
    t = CODE_SUFFIX.sub(' ', s)
    found |= set(CODE_BARE.findall(t))
    # 日期压缩残留：2026xx / 2025xx / 2024xx 不是证券代码；
    # 80xxxx 是占位/内部编号（BSE 实际为 43x/83x/87x/92x），全同数字也不是代码。
    return {c for c in found
            if not c.startswith('20') and not c.startswith('80')
            and len(set(c)) > 1}


def window_patterns(start: str, end: str) -> list[re.Pattern]:
    """支持 YYYYMMDD / YYYY-MM-DD / YYYY/MM/DD / YYYY年M月D日 及跨月窗口。"""
    y, m, d = start[:4], int(start[4:6]), int(start[6:8])
    y2, m2, d2 = end[:4], int(end[4:6]), int(end[6:8])
    pats = [re.compile(r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?')]
    comp = re.compile(r'\d{8}')
    out = [comp]
    # 宽松：只要行内出现窗口起止月内的日期样式即视为候选，精确过滤由 in_window 完成
    return [comp, pats[0]]


def row_in_window(row: tuple, start: str, end: str) -> bool:
    s = ' | '.join(str(v) for v in row if v is not None)
    candidates: list[str] = []
    candidates += re.findall(r'\b(20\d{6})\b', s)
    for y, m, d in re.findall(r'\b(20\d{2})[-/年](\d{1,2})[-/月](\d{1,2})', s):
        candidates.append(f'{y}{int(m):02d}{int(d):02d}')
    for v in row:
        if isinstance(v, (datetime, date)):
            candidates.append(v.strftime('%Y%m%d'))
    return any(start <= c <= end for c in candidates)


# 历史年度库（1/2/3/8/9号）常含「更新日期」列，其值为本期，
# 但行内是往年案例；把它们计入当期代码集合会系统性放大伪差异。
HISTORY_ROLES = {'1', '2', '3', '8', '9'}

# 模板世代差异：第36期及以前把首次披露拆成「沪深」+「北交所」两个 sheet，
# 第37期起合并为「沪深北交易所首次披露」。按原始表名硬对齐会把同一批代码
# 误判成双向差异，因此代码集合比较必须先做语义归一化。
SHEET_ALIASES = {
    '沪深交易所首次披露': '首次披露(沪深北)',
    '北交所首次披露': '首次披露(沪深北)',
    '沪深北交易所首次披露': '首次披露(沪深北)',
    '沪深北首次披露': '首次披露(沪深北)',
}


def canon_sheet(name: str) -> str:
    n = name.strip()
    for k, v in SHEET_ALIASES.items():
        if k in n:
            return v
    return n[:30]


def collect(d: Path, start: str, end: str,
            include_history: bool = False,
            normalize: bool = True) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for p in sorted(d.glob('*.xlsx')):
        if p.name.startswith('~$'):
            continue
        m = re.match(r'^(\d+)、', p.name)
        if not m:
            continue
        role = m.group(1)
        if role in HISTORY_ROLES and not include_history:
            continue
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
        except Exception as exc:
            # 不得静默跳过：云端占位/损坏工作簿会让参考侧代码集合变成空集，
            # 进而把「所有代码都只是生成独有」误报为盲后对比结论。
            raise RuntimeError(f'工作簿不可读: {p}: {type(exc).__name__}: {exc}') from exc
        for ws in wb.worksheets:
            hit: set[str] = set()
            for row in ws.iter_rows(values_only=True):
                if row_in_window(row, start, end):
                    hit |= codes_in_row(row)
            if hit:
                key = f'{role}|{canon_sheet(ws.title) if normalize else ws.title[:30]}'
                out.setdefault(key, set()).update(hit)
        wb.close()
    return out


def role_map(d: Path) -> dict[str, Path]:
    out = {}
    for p in sorted(d.glob('*.xlsx')):
        if p.name.startswith('~$'):
            continue
        m = re.match(r'^(\d+)、', p.name)
        if m:
            out[m.group(1)] = p
    return out


def preflight_workbooks(d: Path, label: str) -> dict:
    """失败关闭预检：参考稿不可读时禁止产出盲后对比结论。

    2026-11-08 Round 5（43→44）实测：`第44期周报/` 与原件快照的 12 份 xlsx
    全部被 OneDrive 回收为云端占位，`read_bytes()` / `zipfile` 抛 ETIMEDOUT。
    旧版 `collect()` 用 `except: continue` 静默跳过，导致参考侧集合为空，
    审计报告错误地显示「仅生成有代码 = 全部，仅现场有代码 = 无」，
    看起来像生成稿全错，实际是根本没有对比。此类结果必须阻断。
    """
    files = sorted(p for p in d.glob('*.xlsx') if not p.name.startswith('~$'))
    failures: list[dict] = []
    for p in files:
        try:
            with zipfile.ZipFile(p) as z:
                names = z.namelist()
            if 'xl/workbook.xml' not in names:
                failures.append({'file': p.name, 'reason': 'MISSING_WORKBOOK_XML'})
        except Exception as exc:
            failures.append({'file': p.name,
                             'reason': f'{type(exc).__name__}: {exc}'})
    return {'label': label, 'dir': str(d), 'xlsx_count': len(files),
            'readable': len(files) - len(failures), 'unreadable': failures,
            'ok': bool(files) and not failures}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--generated', type=Path, required=True)
    ap.add_argument('--reference', type=Path, required=True)
    ap.add_argument('--start', required=True)
    ap.add_argument('--end', required=True)
    ap.add_argument('--compare', type=Path, help='compare_excel.py 的 json，用于提取汇总区差异')
    ap.add_argument('--include-history', action='store_true',
                    help='把 1/2/3/8/9 号历史年度库也计入当期代码集合（默认排除，避免更新日期噪声）')
    ap.add_argument('--no-normalize', action='store_true',
                    help='不做首披 sheet 世代归一化（默认归一化）')
    ap.add_argument('--out', type=Path, required=True)
    a = ap.parse_args()

    pg = preflight_workbooks(a.generated, 'generated')
    pr = preflight_workbooks(a.reference, 'reference')
    if not pg['ok'] or not pr['ok']:
        verdict = ('REFERENCE_UNREADABLE' if not pr['ok']
                   else 'GENERATED_UNREADABLE')
        blocked = {
            'status': 'BLOCKED',
            'verdict': verdict,
            'start': a.start, 'end': a.end,
            'preflight': {'generated': pg, 'reference': pr},
            'note': '参考稿/生成稿含不可读工作簿（常见为 OneDrive 云端占位）。'
                    '已失败关闭，未产出代码集合或汇总区对比；'
                    '此前基于不可读参考目录得到的对比结果一律无效。'
                    '应先物化现场文件，或改用现场仍可读时冻结的派生证据，'
                    '并明确标注为「派生证据级交叉验证」，不得冒充原件级盲后对比。',
        }
        a.out.write_text(json.dumps(blocked, ensure_ascii=False, indent=2),
                         encoding='utf-8')
        a.out.with_suffix('.md').write_text(
            '# Excel 盲测审计：BLOCKED（参考稿不可读）\n\n'
            f'- 判定：`{verdict}`\n'
            f'- 生成稿：{pg["readable"]}/{pg["xlsx_count"]} 可读\n'
            f'- 现场稿：{pr["readable"]}/{pr["xlsx_count"]} 可读\n\n'
            f'{blocked["note"]}\n\n'
            '## 不可读文件\n\n'
            + '\n'.join(f'- {x["file"]}: {x["reason"]}'
                        for x in (pg['unreadable'] + pr['unreadable']))
            + '\n', encoding='utf-8')
        print(json.dumps({'status': 'BLOCKED', 'verdict': verdict,
                          'generated_unreadable': len(pg['unreadable']),
                          'reference_unreadable': len(pr['unreadable'])},
                         ensure_ascii=False))
        return 2

    norm = not a.no_normalize
    g = collect(a.generated, a.start, a.end, a.include_history, norm)
    r = collect(a.reference, a.start, a.end, a.include_history, norm)
    groles, rroles = set(role_map(a.generated)), set(role_map(a.reference))
    boundary = {
        'only_reference_roles': sorted(rroles - groles),
        'only_generated_roles': sorted(groles - rroles),
        'is_structure_exception': bool(rroles - groles),
        'rule': '只用上期模板时，模板未提供的角色不得凭空创建；缺角色应显式标结构例外，'
                '不得从第 n+1 期结果反向构造模板。',
    }
    keys = sorted(set(g) | set(r))
    sheets = []
    for k in keys:
        x, y = g.get(k, set()), r.get(k, set())
        if x == y:
            continue
        sheets.append({'sheet': k, 'only_generated': sorted(x - y),
                       'only_reference': sorted(y - x),
                       'gen_count': len(x), 'ref_count': len(y)})
    only_g = sorted({c for e in sheets for c in e['only_generated']})
    only_r = sorted({c for e in sheets for c in e['only_reference']})

    summary = []
    if a.compare and a.compare.is_file():
        cmp = json.loads(a.compare.read_text(encoding='utf-8'))
        for e in cmp.get('sheets', []):
            if e.get('summary_diff'):
                summary.append({'role': e['role'], 'sheet': e['sheet'],
                                'gen_summary': e.get('gen_summary', [])[:6],
                                'ref_summary': e.get('ref_summary', [])[:6]})
        totals = cmp.get('totals', {})
    else:
        totals = {}

    result = {
        'status': 'OK',
        'start': a.start, 'end': a.end,
        'generated': str(a.generated), 'reference': str(a.reference),
        'preflight': {'generated': pg, 'reference': pr},
        'roles': {'generated': sorted(groles), 'reference': sorted(rroles)},
        'template_boundary': boundary,
        'code_set': {'sheets_with_diff': len(sheets), 'sheets_total': len(keys),
                     'only_generated_codes': only_g, 'only_reference_codes': only_r,
                     'sheets': sheets},
        'summary_cache_diff': {'sheets': len(summary), 'items': summary},
        'compare_totals': totals,
    }
    a.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    md = a.out.with_suffix('.md')
    lines = [f'# Excel 盲测审计 {a.start}–{a.end}', '',
             f'- 角色：生成 {len(result["roles"]["generated"])} / 现场 {len(result["roles"]["reference"])}',
             f'- 模板边界：{"**结构例外**" if boundary["is_structure_exception"] else "一致"}；'
             f'现场多出角色 {boundary["only_reference_roles"] or "无"}（模板未提供，不得凭空创建）',
             f'- 当期代码集合统计范围：{"全部角色" if a.include_history else "排除 1/2/3/8/9 号历史年度库（避免更新日期噪声）"}',
             f'- sheet 语义归一化：{"关闭" if a.no_normalize else "开启（沪深/北交所首披 → 沪深北首披）"}',
             f'- 当期代码集合有差异的 sheet：**{len(sheets)} / {len(keys)}**',
             f'- 仅生成有代码：{only_g or "无"}', f'- 仅现场有代码：{only_r or "无"}',
             f'- 汇总区缓存差异 sheet：**{len(summary)}**', '',
             '## 代码集合差异明细', '', '| 角色\\|sheet | 仅生成 | 仅现场 |', '|---|---|---|']
    for e in sheets:
        lines.append(f'| {e["sheet"]} | {e["only_generated"] or "-"} | {e["only_reference"] or "-"} |')
    lines += ['', '## 汇总区缓存差异（公式缓存未随明细重算）', '']
    for e in summary:
        lines.append(f'- [{e["role"]}] {e["sheet"]}')
        for x, y in zip(e['gen_summary'], e['ref_summary']):
            if x != y:
                lines.append(f'  - 生成 `{x[:160]}`')
                lines.append(f'  - 现场 `{y[:160]}`')
    md.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    print(json.dumps({'code_diff_sheets': len(sheets), 'sheets_total': len(keys),
                      'only_generated_codes': only_g, 'only_reference_codes': only_r,
                      'summary_cache_diff_sheets': len(summary)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
