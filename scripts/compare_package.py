#!/usr/bin/env python3
"""独立成品 vs 参考成品：逐角色审计。

建议在独立成品发布、冻结并记录哈希后再执行。
只读两边文件，不修改任何一方。
输出 compare43.json（机器可读）与 compare43.md（人读摘要）。
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
from collections import Counter
from pathlib import Path

import openpyxl
from docx import Document
from pypdf import PdfReader

NUM_RE = re.compile(r'(?<![\d.])\d+(?:\.\d+)?\s*(?:单|次|家|家次|亿元|万元|起|笔|项|%)')
DATE_RE = re.compile(r'20\d{2}[-年]\d{1,2}[-月]\d{1,2}日?|\d{8}')


def role(name: str) -> str | None:
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


def norm(s) -> str:
    if s is None:
        return ''
    return re.sub(r'\s+', '', str(s)).strip()


def cell_rows(path: Path, max_rows: int = 4000) -> dict:
    out: dict = {'sheets': {}, 'errors': []}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as ex:  # noqa: BLE001
        out['errors'].append(f'{type(ex).__name__}: {ex}')
        return out
    try:
        for ws in wb.worksheets:
            rows = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= max_rows:
                    break
                vals = [norm(v) for v in row]
                while vals and not vals[-1]:
                    vals.pop()
                if any(vals):
                    rows.append('|'.join(vals))
            out['sheets'][ws.title] = rows
    finally:
        wb.close()
    return out


def _raw_docx_text(path: Path) -> list[str]:
    """OOXML 原始兜底：现场 Word 含 word/NULL 等 Word 特有部件，
    python-docx 可能直接抛错；直接读 document.xml 的 <w:t> 才不漏内容。"""
    import zipfile
    from xml.etree import ElementTree as ET
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    out = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not (name.startswith('word/') and name.endswith('.xml')
                    and ('document' in name or 'header' in name or 'footer' in name
                         or 'footnote' in name or 'endnote' in name)):
                continue
            try:
                root = ET.fromstring(z.read(name))
            except ET.ParseError:
                continue
            for para in root.iter(W + 'p'):
                txt = ''.join(t.text or '' for t in para.iter(W + 't')).strip()
                if txt:
                    out.append(txt)
    return out


def docx_text(path: Path) -> list[str]:
    try:
        d = Document(path)
    except Exception:  # noqa: BLE001 - 现场稿常有 Word 特有部件
        return _raw_docx_text(path)
    chunks = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    for t in d.tables:
        for r in t.rows:
            line = '|'.join(c.text.strip() for c in r.cells)
            if line.strip(' |'):
                chunks.append(line)
    raw = _raw_docx_text(path)
    return raw if len(raw) > len(chunks) * 2 else chunks


def pptx_text(path: Path) -> list[str]:
    try:
        prs = __import__('pptx').Presentation(path)
    except Exception as ex:  # noqa: BLE001
        return [f'<ERROR {type(ex).__name__}: {ex}>']
    chunks = []
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text.strip():
                chunks.append(f'[p{i}] {sh.text.strip()}')
            if getattr(sh, 'has_table', False):
                for r in sh.table.rows:
                    chunks.append(f'[p{i}] ' + '|'.join(c.text.strip() for c in r.cells))
        if prs.slides[i - 1].has_notes_slide:
            nt = prs.slides[i - 1].notes_slide.notes_text_frame.text.strip()
            if nt:
                chunks.append(f'[p{i} notes] {nt}')
    return chunks


def pdf_info(path: Path) -> dict:
    try:
        r = PdfReader(str(path))
        pages = []
        for i, pg in enumerate(r.pages, 1):
            try:
                t = pg.extract_text() or ''
            except Exception:
                t = ''
            pages.append(t)
        return {'pages': len(pages), 'blank': [i for i, t in enumerate(pages, 1)
                                               if len(t.strip()) < 10],
                'chars': [len(t) for t in pages]}
    except Exception as ex:  # noqa: BLE001
        return {'error': f'{type(ex).__name__}: {ex}'}


def numbers(lines: list[str]) -> Counter:
    c: Counter = Counter()
    for ln in lines:
        for m in NUM_RE.findall(ln):
            c[re.sub(r'\s+', '', m)] += 1
    return c


def collection(root: Path) -> dict:
    out: dict = {'roles': {}, 'pdfs': {}, 'other': []}
    for p in sorted(root.iterdir()):
        if not p.is_file() or p.name.startswith('~$'):
            continue
        r = role(p.name)
        if p.suffix.lower() in ('.xlsx', '.xlsm') and r:
            out['roles'][r] = {'name': p.name, 'kind': 'xlsx', 'data': cell_rows(p)}
        elif p.suffix.lower() == '.docx' and r:
            out['roles'][r] = {'name': p.name, 'kind': 'docx', 'data': docx_text(p)}
        elif p.suffix.lower() == '.pptx' and r:
            out['roles'][r] = {'name': p.name, 'kind': 'pptx', 'data': pptx_text(p)}
        elif p.suffix.lower() == '.pdf':
            out['pdfs'][p.name] = pdf_info(p)
        else:
            out['other'].append(p.name)
    return out


def compare(gend: Path, refd: Path) -> dict:
    g, r = collection(gend), collection(refd)
    res: dict = {
        'generated_dir': str(gend), 'reference_dir': str(refd),
        'role_coverage': {}, 'role_diffs': {}, 'pdf_diffs': {},
        'number_diffs': {}, 'summary': {},
    }
    allroles = sorted(set(g['roles']) | set(r['roles']))
    res['role_coverage'] = {
        'generated': sorted(g['roles']), 'reference': sorted(r['roles']),
        'missing_in_generated': sorted(set(r['roles']) - set(g['roles'])),
        'extra_in_generated': sorted(set(g['roles']) - set(r['roles'])),
        'names': {k: {'generated': g['roles'].get(k, {}).get('name'),
                      'reference': r['roles'].get(k, {}).get('name')} for k in allroles},
    }
    for k in sorted(set(g['roles']) & set(r['roles'])):
        gi, ri = g['roles'][k], r['roles'][k]
        if gi['kind'] != ri['kind']:
            res['role_diffs'][k] = {'kind_mismatch': [gi['kind'], ri['kind']]}
            continue
        d: dict = {}
        if gi['kind'] == 'xlsx':
            gs, rs = gi['data']['sheets'], ri['data']['sheets']
            d['sheets_generated'] = sorted(gs)
            d['sheets_reference'] = sorted(rs)
            d['sheets_missing'] = sorted(set(rs) - set(gs))
            d['sheets_extra'] = sorted(set(gs) - set(rs))
            for s in sorted(set(gs) & set(rs)):
                a, b = gs[s], rs[s]
                ca, cb = Counter(a), Counter(b)
                add, rem = ca - cb, cb - ca
                d.setdefault('sheets_detail', {})[s] = {
                    'rows_generated': len(a), 'rows_reference': len(b),
                    'added_rows': sum(add.values()), 'removed_rows': sum(rem.values()),
                    'added_samples': [x[:220] for x in list(add.elements())[:8]],
                    'removed_samples': [x[:220] for x in list(rem.elements())[:8]],
                }
        else:
            a, b = gi['data'], ri['data']
            ca, cb = Counter(a), Counter(b)
            add, rem = ca - cb, cb - ca
            d.update({
                'lines_generated': len(a), 'lines_reference': len(b),
                'added_lines': sum(add.values()), 'removed_lines': sum(rem.values()),
                'added_samples': [x[:300] for x in list(add.elements())[:12]],
                'removed_samples': [x[:300] for x in list(rem.elements())[:12]],
                'similarity': round(difflib.SequenceMatcher(None, '\n'.join(a), '\n'.join(b)).ratio(), 4),
            })
            gn, rn = numbers(a), numbers(b)
            d['numbers_only_generated'] = dict((gn - rn).most_common(25))
            d['numbers_only_reference'] = dict((rn - gn).most_common(25))
        res['role_diffs'][k] = d
    gpn = {re.sub(r'\d{8}', 'DATE', n): i for n, i in g['pdfs'].items()}
    rpn = {re.sub(r'\d{8}', 'DATE', n): i for n, i in r['pdfs'].items()}
    for k in sorted(set(gpn) | set(rpn)):
        res['pdf_diffs'][k] = {
            'generated_pages': gpn.get(k, {}).get('pages'),
            'reference_pages': rpn.get(k, {}).get('pages'),
            'generated_blank': gpn.get(k, {}).get('blank'),
            'reference_blank': rpn.get(k, {}).get('blank'),
        }
    res['summary'] = {
        'roles_generated': len(g['roles']), 'roles_reference': len(r['roles']),
        'pdfs_generated': len(g['pdfs']), 'pdfs_reference': len(r['pdfs']),
        'other_generated': g['other'], 'other_reference': r['other'],
    }
    return res


def md_report(res: dict) -> str:
    L = ['# 并购周报独立重建对照报告', '',
         f"- 生成稿：`{res['generated_dir']}`", f"- 现场稿：`{res['reference_dir']}`", '',
         '## 1. 文件角色', '']
    rc = res['role_coverage']
    L += [f"- 生成稿角色 {len(rc['generated'])} 个 / 现场稿角色 {len(rc['reference'])} 个",
          f"- 生成稿缺失角色：{rc['missing_in_generated'] or '无'}",
          f"- 生成稿多出角色：{rc['extra_in_generated'] or '无'}", '',
          '| 角色 | 生成稿文件名 | 现场稿文件名 |', '|---|---|---|']
    for k, v in rc['names'].items():
        L.append(f"| {k} | {v['generated'] or '—'} | {v['reference'] or '—'} |")
    L += ['', '## 2. 逐角色内容差异', '']
    for k, d in res['role_diffs'].items():
        L += [f'### {k}', '']
        if 'kind_mismatch' in d:
            L += [f"- 类型不一致：{d['kind_mismatch']}", '']
            continue
        if 'sheets_detail' in d:
            L += [f"- Sheet：生成 {d['sheets_generated']} / 现场 {d['sheets_reference']}",
                  f"- 缺 Sheet：{d['sheets_missing'] or '无'}；多 Sheet：{d['sheets_extra'] or '无'}", '',
                  '| Sheet | 生成行数 | 现场行数 | 新增行 | 缺失行 |', '|---|---|---|---|---|']
            for s, x in d['sheets_detail'].items():
                L.append(f"| {s} | {x['rows_generated']} | {x['rows_reference']} | "
                         f"{x['added_rows']} | {x['removed_rows']} |")
            for s, x in d['sheets_detail'].items():
                if x['removed_samples']:
                    L += ['', f'**{s} 现场有而生成稿无（样例）**', '']
                    L += [f'- `{v}`' for v in x['removed_samples'][:5]]
                if x['added_samples']:
                    L += ['', f'**{s} 生成稿有而现场无（样例）**', '']
                    L += [f'- `{v}`' for v in x['added_samples'][:5]]
        else:
            L += [f"- 行数：生成 {d['lines_generated']} / 现场 {d['lines_reference']}；"
                  f"新增 {d['added_lines']} 行、缺失 {d['removed_lines']} 行；"
                  f"文本相似度 {d['similarity']}",
                  f"- 仅生成稿出现的数字口径：{d['numbers_only_generated']}",
                  f"- 仅现场稿出现的数字口径：{d['numbers_only_reference']}", '']
            if d['removed_samples']:
                L += ['**现场有而生成稿无（样例）**', ''] + [f'- `{v}`' for v in d['removed_samples'][:8]]
            if d['added_samples']:
                L += ['', '**生成稿有而现场无（样例）**', ''] + [f'- `{v}`' for v in d['added_samples'][:8]]
        L.append('')
    L += ['## 3. PDF 页数', '', '| PDF | 生成页数 | 现场页数 | 生成空白页 | 现场空白页 |',
          '|---|---|---|---|---|']
    for k, v in res['pdf_diffs'].items():
        L.append(f"| {k} | {v['generated_pages']} | {v['reference_pages']} | "
                 f"{v['generated_blank']} | {v['reference_blank']} |")
    L += ['', '## 4. 其它文件', '',
          f"- 生成稿：`{res['summary']['other_generated']}`",
          f"- 现场稿：`{res['summary']['other_reference']}`", '']
    return '\n'.join(L)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--generated', required=True, type=Path)
    ap.add_argument('--reference', required=True, type=Path)
    ap.add_argument('--out', required=True, type=Path)
    a = ap.parse_args()
    res = compare(a.generated, a.reference)
    a.out.parent.mkdir(parents=True, exist_ok=True)
    a.out.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding='utf-8')
    a.out.with_suffix('.md').write_text(md_report(res), encoding='utf-8')
    print(json.dumps(res['summary'], ensure_ascii=False))
    print(json.dumps(res['role_coverage']['missing_in_generated'], ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
