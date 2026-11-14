#!/usr/bin/env python3
"""第43期「已知答案回放」回归门禁。

来源：2026-11-08 严格盲测回放（见 references/blind-replay-43.md）。
现场第43期人工成果的口径事实在**生成完成后**才拆封读取。GOLDEN 同时
记录官方稿明确事实和工作簿勾稽事实；二者冲突时保留差异检查，不用文本替换
掩盖。v4.1 的核心事实检查通过，但 Word/PPT 模板体量仍以 WARN 标记，因此
该门禁不会把“数据正确”冒称为“可直接发布”。

用法：

    python scripts/test_known_answer_43.py \
        --run-dir  <auto_run 的 runs/2026-10-31> \
        --package  <发布出去的成品目录> \
        [--reference <现场第43期目录>] \
        [--report   <输出 json 路径>]

退出码：0=无 FAIL（可有 WARN）；1=存在 FAIL；2=参数/文件缺失。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# 现场第43期《A股市场并购重组动态》与《并购重组周报》正文口径（拆封后读取）
GOLDEN_43 = {
    'weekly.first_disclosed': 2,          # 科瑞思、博盈特焊（预案）
    'weekly.initial_projects': 3,         # 再加节能铁汉现金重大交易草案
    'weekly.first_draft': 1,              # 节能铁汉（购买/出售双表，市场按代码去重）
    'weekly.accepted': 1,                 # 秉扬科技（开源证券）
    'weekly.committee_passed': 3,         # 国风新材、东方证券、永达股份
    'weekly.registered': 1,               # 创远信科（国泰海通）
    'weekly.terminated': 1,               # 天壕能源，终止表内计 1 单
    'weekly.terminated_in_review': 0,     # 天壕能源为预案阶段，不计在审终止
}
GOLDEN_WORKBOOK = {
    'first_disclosed': 76,
    'passed': 39,
    'registered': 33,
    'terminated': 67,
    'monthly_reviews': 40,                 # 38单普通程序 + 2单简易程序
    'control_first': 149,
    'control_completed': 123,
    'strategic': 228,
    'current_strategic': 9,
    'current_professional': 23,
    'current_partial_transfer': 4,
    'in_review': 48,                       # 工作簿“是否在审=是”勾稽值
}

# 注册批复事件必须命中并购重组口径，不能把再融资「向特定对象发行股票」计入
GOLDEN_REGISTERED_CODES = {'920961'}      # 创远信科
FORBIDDEN_REGISTERED_TOKENS = ('向特定对象发行股票', '向特定对象发行A股',
                               '再融资', '非公开发行股票')
# 现场文档体量（OOXML 原始段落/形状文本行数）
GOLDEN_TEXT_VOLUME = {'weekly_docx': 1620, 'dynamic_docx': 141, 'market_ppt': 187}
STALE_TOKENS = ('20261018至20261024', '第42期', '截至10月24日', '2026年10月18日至2026年10月24日')


def jload(p: Path, default=None):
    if not p.is_file():
        return default
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except (ValueError, OSError):
        return default


class Checks:
    def __init__(self):
        self.rows: list[dict] = []

    def eq(self, name: str, got, want, note: str = ''):
        self.rows.append({'check': name, 'got': got, 'want': want,
                          'status': 'PASS' if got == want else 'FAIL', 'note': note})

    def ok(self, name: str, cond: bool, got, note: str = ''):
        self.rows.append({'check': name, 'got': got, 'want': True,
                          'status': 'PASS' if cond else 'FAIL', 'note': note})

    def warn(self, name: str, cond: bool, got, note: str = ''):
        self.rows.append({'check': name, 'got': got, 'want': True,
                          'status': 'PASS' if cond else 'WARN', 'note': note})

    @property
    def failed(self):
        return [r for r in self.rows if r['status'] == 'FAIL']

    @property
    def warnings(self):
        return [r for r in self.rows if r['status'] == 'WARN']


def metric(m: dict, dotted: str):
    cur = m
    for part in dotted.split('.'):
        if not isinstance(cur, dict):
            return None
        cur = cur.get(part)
    if isinstance(cur, dict):
        return cur.get('projects')
    return cur


def docx_lines(p: Path) -> int:
    from xml.etree import ElementTree as ET
    from zipfile import ZipFile
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    n = 0
    with ZipFile(p) as z:
        for name in z.namelist():
            if name.startswith('word/') and name.endswith('.xml') and 'document' in name:
                try:
                    root = ET.fromstring(z.read(name))
                except ET.ParseError:
                    continue
                for para in root.iter(W + 'p'):
                    t = ''.join(x.text or '' for x in para.iter(W + 't')).strip()
                    if t:
                        n += 1
    return n


def pptx_lines(p: Path) -> int:
    from pptx import Presentation
    prs = Presentation(p)
    return sum(1 for s in prs.slides for sh in s.shapes
               if sh.has_text_frame and sh.text.strip())


def role_files(pkg: Path) -> dict:
    out = {}
    for p in pkg.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith('并购重组市场周报') and p.suffix == '.pptx':
            out['market_ppt'] = p
        elif p.name.startswith('A股市场并购重组动态') and p.suffix == '.docx':
            out['dynamic_docx'] = p
        elif re.match(r'^5、', p.name) and p.suffix == '.docx':
            out['weekly_docx'] = p
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--run-dir', required=True, type=Path)
    ap.add_argument('--package', required=True, type=Path)
    ap.add_argument('--report', type=Path)
    a = ap.parse_args()
    if not a.run_dir.is_dir() or not a.package.is_dir():
        print('run-dir / package 不存在', file=sys.stderr)
        return 2
    m = jload(a.run_dir / 'metrics.json', {})
    led = jload(a.run_dir / 'event-ledger.json', {})
    c = Checks()

    for dotted, want in GOLDEN_43.items():
        c.eq('口径 ' + dotted, metric(m, dotted), want)

    # Annual and as-of totals require the updated base workbook, not the current
    # week's event cohort.  Compare against the workbook produced by BUILD.
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from build_reports import workbook_facts
    wf = workbook_facts(a.package)
    for key, want in GOLDEN_WORKBOOK.items():
        c.eq('工作簿勾稽 ' + key, wf.get(key), want)
    c.ok('年度累计不得等同本期事件', m.get('annual') != m.get('weekly'),
         'annual == weekly' if m.get('annual') == m.get('weekly') else 'independent')
    c.ok('官方正文49单与工作簿48单的内部差异已显式登记',
         wf.get('in_review') == 48, wf.get('in_review'),
         '官方正文写49；工作簿“是否在审=是”为48，须由业务确认')

    # 注册批复口径：不得混入再融资类公告
    reg = (m.get('weekly') or {}).get('registered', {}) or {}
    reg_ids = reg.get('deal_ids', []) or []
    codes = {re.sub(r'\D', '', x)[:6] for x in reg_ids}
    c.ok('注册批复命中并购重组标的（创远信科 920961）',
         bool(GOLDEN_REGISTERED_CODES & codes), sorted(codes))
    blob = json.dumps(led, ensure_ascii=False)
    hits = [t for t in FORBIDDEN_REGISTERED_TOKENS if t in blob]
    c.ok('事件台账未把再融资公告当作并购注册批复', not hits, hits[:6])

    # 追加行序号必须是数值，不能是 '=A149+1' 文本
    import openpyxl
    seq_text_cells = []
    for p in sorted(a.package.glob('*.xlsx')):
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=False)
        except Exception:
            continue
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows(min_col=1, max_col=1):
                    cell = row[0]
                    # 真公式 data_type='f'；被写成文本的伪公式 data_type='s'
                    if (cell.data_type == 's' and isinstance(cell.value, str)
                            and re.fullmatch(r'=A\d+\+1', cell.value.strip())):
                        seq_text_cells.append(f'{p.name}/{ws.title}/{cell.coordinate}')
        finally:
            wb.close()
    c.ok('追加行序号为数值而非公式文本', not seq_text_cells,
         seq_text_cells[:8])
    seq_date_cells = []
    current_roles = {'0','4','7','10','11','12','13'}
    for p in sorted(a.package.glob('*.xlsx')):
        m=re.match(r'^(\d+)、',p.name)
        if not m or m.group(1) not in current_roles:continue
        try: wb = openpyxl.load_workbook(p, read_only=True, data_only=False)
        except Exception: continue
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows(min_col=1, max_col=1):
                    cell = row[0]
                    if hasattr(cell.value, 'year') and cell.value.year == 1900:
                        seq_date_cells.append(f'{p.name}/{ws.title}/{cell.coordinate}')
        finally: wb.close()
    c.ok('新增序号未被日期样式显示为1900年', not seq_date_cells, seq_date_cells[:8])

    # 文档体量：现场稿是重表格/重图表文档，从零重写会塌缩
    rf = role_files(a.package)
    for role, want in GOLDEN_TEXT_VOLUME.items():
        p = rf.get(role)
        if not p:
            c.eq('文档体量 ' + role, None, f'>={int(want * 0.7)}')
            continue
        try:
            got = pptx_lines(p) if role == 'market_ppt' else docx_lines(p)
        except Exception as ex:  # noqa: BLE001
            got = f'ERROR {type(ex).__name__}'
        c.warn(f'文档体量 {role} >= 现场 70%（{int(want * 0.7)} 行）',
               isinstance(got, int) and got >= want * 0.7, got, f'现场 {want} 行；体量不足属于版式保真例外')

    # 旧期次残留
    stale = []
    for p in rf.values():
        try:
            from zipfile import ZipFile
            with ZipFile(p) as z:
                text = ' '.join(z.read(n).decode('utf-8', 'ignore')
                                for n in z.namelist() if n.endswith('.xml')
                                and ('slide' in n or 'document' in n))
        except Exception:
            continue
        stale += [t for t in STALE_TOKENS if t in text]
    c.ok('Word/PPT 无上期残留', not stale, sorted(set(stale)))

    out = {'as_of': '2026-10-31', 'golden_source': '官方原文与第43期人工成果（盲测拆封后核对；补充项经公告原件验证）',
           'passed': len([r for r in c.rows if r['status'] == 'PASS']),
           'warnings': len(c.warnings), 'failed': len(c.failed), 'checks': c.rows}
    if a.report:
        a.report.parent.mkdir(parents=True, exist_ok=True)
        a.report.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
    for r in c.rows:
        print(f"[{r['status']}] {r['check']}: got={r['got']} want={r['want']}")
    result = 'KNOWN_ANSWER_PASS' if not c.failed else 'KNOWN_ANSWER_FAIL'
    if not c.failed and c.warnings:
        result = 'KNOWN_ANSWER_PASS_WITH_WARNINGS'
    print(f"\nPASS {out['passed']} / WARN {out['warnings']} / FAIL {out['failed']}  -> {result}")
    return 1 if c.failed else 0


if __name__ == '__main__':
    raise SystemExit(main())
