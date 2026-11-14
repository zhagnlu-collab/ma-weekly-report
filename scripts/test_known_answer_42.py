#!/usr/bin/env python3
"""第42期「已知答案回放」回归门禁。

来源：2026-11-08 严格盲测回放（见 references/blind-replay-42.md）。
现场第42期人工成果在**生成达到终态并冻结哈希后**才拆封读取；GOLDEN 取自
官方稿正文与现场工作簿勾稽，生成链在拆封前没有接触过这些数值。

用法：

    python scripts/test_known_answer_42.py \
        --run-dir  <auto_run 的 runs/2026-10-24> \
        --package  <发布出去的成品目录> \
        [--reference <现场第42期目录>] \
        [--report   <输出 json 路径>]

退出码：0=无 FAIL（可有 WARN）；1=存在 FAIL；2=参数/文件缺失。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------- 已知答案
# 现场第42期《2026年并购重组周报》正文，统计区间 2026-10-18 至 2026-10-24。
# 注意：现场把「本期新增首次披露草案」（天津港）单列于「八、本期新披露案例
# 简介」，但**不计入本期首次披露单数**；生成稿若把二者相加写成 3 单，属表述
# 口径分歧，事件台账本身 first_disclosed=2 / first_draft=1 是对的。
GOLDEN_42_WEEKLY = {
    'weekly.first_disclosed': 2,       # 三峡新材（沪主板）、新亚制程（深主板）
    'weekly.first_draft': 1,           # 天津港，单列、不并入首次披露单数
    'weekly.accepted': 0,              # 各板块「本期无新增受理」
    'weekly.committee_passed': 1,      # 光库科技，深交所 2026 年第 16 次会议
    'weekly.registered': 0,            # 「本期无新增取得证监会注册批复的交易」
    'weekly.terminated': 0,            # 各板块「本期无新增终止的在审交易」
    'weekly.simplified_passed': 0,
}

# 现场工作簿勾稽（用 build_reports.workbook_facts 对现场目录计算得到）。
# 前 5 项是年度累计核心口径，v4.1.0 已能桥接正确；后半是明细台账口径。
GOLDEN_WORKBOOK_42 = {
    'first_disclosed': 74,
    'passed': 36,
    'registered': 32,
    'terminated': 66,
    'monthly_reviews': 37,
    'cash_purchase': 26,
    'cash_sale': 7,
    'control_first': 146,
    'control_completed': 122,
    'strategic': 222,
    'professional': 482,
    'partial_transfer': 16,
    'current_strategic': 1,
    'current_professional': 7,
    'current_partial_transfer': 3,
    'in_review': 50,                   # 现场工作簿勾稽 50；现场正文分板块合计 51
}

# 现场正文的年度累计与叙述性口径（宽松正则，去除空白后匹配生成稿 Word）。
NARRATIVE_FACTS = {
    '年度首次披露74单': r'首次披露74单|共披露74单',
    '年度注册批复32单': r'共32单取得证监会注册批复|注册批复.{0,12}32单',
    '重组委会议37次': r'召开37次并购重组委会议|37次.{0,6}重组委',
    '合计审核37单': r'合计共审核37单|已审核37单',
    '审核通过36单': r'审核通过36单|通过36单',
    '否决1单': r'否决1单',
    '独立财务顾问44家次': r'44家次',
    '分板块25/16/9/21/3': r'沪主板25单.{0,10}深主板16单',
    '沪主板在审17单': r'沪主板在审项目17单|沪主板.{0,6}17单在审',
    '在审分板块17/16/3/15': r'深主板在审项目16单',
    '券商汇总表华泰联合1065.32': r'1,065\.32|1065\.32',
    '下周会议10月29日': r'2026年10月29日',
    '下周审议国风新材': r'国风新材',
    '下周审议东方证券': r'东方证券',
    '本周通过光库科技': r'光库科技',
    '本期首次披露三峡新材': r'三峡新材',
    '本期首次披露新亚制程': r'新亚制程',
    '本期首次披露草案天津港': r'天津港',
}

# 现场文档体量（OOXML 原始段落 / PPT 形状文本行数）。第42期现场**没有**
# 《A股市场并购重组动态》Word，故不设该角色基准。
GOLDEN_TEXT_VOLUME_42 = {'weekly_docx': 1597, 'market_ppt': 159}
STALE_TOKENS_42 = ('20261011至20261017', '第41期', '截至10月17日',
                   '2026年10月11日至2026年10月17日')


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

    def _add(self, status, name, got, want, note=''):
        self.rows.append({'check': name, 'got': got, 'want': want,
                          'status': status, 'note': note})

    def eq(self, name, got, want, note=''):
        self._add('PASS' if got == want else 'FAIL', name, got, want, note)

    def ok(self, name, cond, got, note=''):
        self._add('PASS' if cond else 'FAIL', name, got, True, note)

    def warn(self, name, cond, got, note=''):
        self._add('PASS' if cond else 'WARN', name, got, True, note)

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


def ooxml_text(p: Path) -> str:
    """拼接 Word 全部 <w:t>，表格/页眉/脚注一并纳入，避免 python-docx 部件报错。"""
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    chunks = []
    with zipfile.ZipFile(p) as z:
        for name in z.namelist():
            if not (name.startswith('word/') and name.endswith('.xml')):
                continue
            try:
                root = ET.fromstring(z.read(name))
            except ET.ParseError:
                continue
            for para in root.iter(W + 'p'):
                t = ''.join(x.text or '' for x in para.iter(W + 't'))
                if t.strip():
                    chunks.append(t)
    return re.sub(r'\s+', '', '\n'.join(chunks))


def docx_lines(p: Path) -> int:
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    n = 0
    with zipfile.ZipFile(p) as z:
        for name in z.namelist():
            if name.startswith('word/') and name.endswith('.xml') and 'document' in name:
                try:
                    root = ET.fromstring(z.read(name))
                except ET.ParseError:
                    continue
                for para in root.iter(W + 'p'):
                    if ''.join(x.text or '' for x in para.iter(W + 't')).strip():
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
    ap.add_argument('--reference', type=Path)
    ap.add_argument('--report', type=Path)
    a = ap.parse_args()
    if not a.run_dir.is_dir() or not a.package.is_dir():
        print('run-dir / package 不存在', file=sys.stderr)
        return 2
    m = jload(a.run_dir / 'metrics.json', {})
    c = Checks()

    # 1) 本期事件口径
    for dotted, want in GOLDEN_42_WEEKLY.items():
        c.eq('本期口径 ' + dotted, metric(m, dotted), want)
    c.ok('年度累计不得等同本期事件', m.get('annual') != m.get('weekly'),
         'annual == weekly' if m.get('annual') == m.get('weekly') else 'independent')

    # 2) 年度累计与明细台账勾稽（对 BUILD 后的成品工作簿）
    for cand in (Path(__file__).resolve().parent,
                 a.run_dir.parent.parent / 'skill' / 'scripts'):
        if (cand / 'build_reports.py').is_file():
            sys.path.insert(0, str(cand))
            break
    try:
        from build_reports import workbook_facts  # noqa: E402
        wf = workbook_facts(a.package)
    except Exception as ex:  # noqa: BLE001
        wf = {}
        c.ok('workbook_facts 可调用', False, f'{type(ex).__name__}: {ex}')
    for key, want in GOLDEN_WORKBOOK_42.items():
        note = ('现场正文分板块合计51，工作簿勾稽50，属现场稿内部差异，'
                '需业务确认' if key == 'in_review' else '')
        c.eq('工作簿勾稽 ' + key, wf.get(key), want, note)

    # 3) 叙述层：现场正文明确事实是否在生成稿 Word 中保留
    rf = role_files(a.package)
    wp = rf.get('weekly_docx')
    text = ooxml_text(wp) if wp else ''
    covered = 0
    for label, pat in NARRATIVE_FACTS.items():
        hit = bool(re.search(pat, text))
        covered += int(hit)
        c.ok('叙述口径 ' + label, hit, 'hit' if hit else 'missing')
    c.ok(f'叙述口径覆盖率 >= 80%（{covered}/{len(NARRATIVE_FACTS)}）',
         covered >= len(NARRATIVE_FACTS) * 0.8, f'{covered}/{len(NARRATIVE_FACTS)}')

    # 4) 追加行序号必须是数值
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
                    if (cell.data_type == 's' and isinstance(cell.value, str)
                            and re.fullmatch(r'=A\d+\+1', cell.value.strip())):
                        seq_text_cells.append(f'{p.name}/{ws.title}/{cell.coordinate}')
        finally:
            wb.close()
    c.ok('追加行序号为数值而非公式文本', not seq_text_cells, seq_text_cells[:8])

    # 5) 文档体量（WARN：从零重写导致塌缩，属版式保真例外而非数据错误）
    for role, want in GOLDEN_TEXT_VOLUME_42.items():
        p = rf.get(role)
        if not p:
            c.eq('文档体量 ' + role, None, f'>={int(want * 0.7)}')
            continue
        try:
            got = pptx_lines(p) if role == 'market_ppt' else docx_lines(p)
        except Exception as ex:  # noqa: BLE001
            got = f'ERROR {type(ex).__name__}'
        c.warn(f'文档体量 {role} >= 现场 70%（{int(want * 0.7)} 行）',
               isinstance(got, int) and got >= want * 0.7, got,
               f'现场 {want} 行；体量不足属于版式保真例外')

    # 6) 旧期次残留
    stale = []
    for p in rf.values():
        try:
            with zipfile.ZipFile(p) as z:
                blob = ' '.join(z.read(n).decode('utf-8', 'ignore')
                                for n in z.namelist() if n.endswith('.xml')
                                and ('slide' in n or 'document' in n))
        except Exception:
            continue
        stale += [t for t in STALE_TOKENS_42 if t in blob]
    c.ok('Word/PPT 无第41期残留', not stale, sorted(set(stale)))

    # 7) 文件角色结构
    native = [p for p in a.package.iterdir()
              if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm', '.docx', '.pptx')]
    c.ok('原生 Office 文件 >= 16 个', len(native) >= 16, len(native))

    out = {
        'as_of': '2026-10-24', 'issue': 42,
        'golden_source': '现场第42期人工成果正文与工作簿（盲测拆封后读取；生成链未接触）',
        'passed': len([r for r in c.rows if r['status'] == 'PASS']),
        'warnings': len(c.warnings), 'failed': len(c.failed),
        'narrative_covered': covered, 'narrative_total': len(NARRATIVE_FACTS),
        'workbook_facts': wf, 'checks': c.rows,
    }
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
