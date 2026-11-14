#!/usr/bin/env python3
"""只读盘点Office模板；不重算、不更改文件、不把错误值直接判为业务错误。"""
import argparse
import collections
import hashlib
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'

def inspect(path):
    result = {'file': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'bytes': path.stat().st_size}
    if path.suffix.lower() == '.pdf':
        result['kind'] = 'pdf'
        return result
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        result['package_parts'] = names
        result['broken_null_relationships'] = [n for n in names if n.endswith('.rels') and any(x.attrib.get('Target', '').upper() in ('NULL', '') for x in ET.fromstring(z.read(n)).findall(R+'Relationship'))]
        result['embedded_parts'] = [n for n in names if '/embeddings/' in n or 'externalLink' in n]
        suffix = path.suffix.lower()
        if suffix in ('.xlsx', '.xlsm'):
            try:
                import openpyxl
            except ImportError as ex:
                raise RuntimeError('需openpyxl才能盘点工作簿；请使用已定位的可用Python环境') from ex
            wb = openpyxl.load_workbook(path, data_only=False, keep_vba=suffix == '.xlsm')
            values = openpyxl.load_workbook(path, data_only=True)
            result.update(kind='workbook', sheets=[], defined_names=list(wb.defined_names), external_link_count=len(wb._external_links))
            for ws in wb:
                cells = [c for c in ws._cells.values() if c.value is not None]
                rows = collections.defaultdict(list)
                for c in cells:
                    if c.row <= 8:
                        rows[c.row].append({'cell': c.coordinate, 'value': str(c.value)[:250]})
                errors = [{'cell': c.coordinate, 'value': c.value} for c in values[ws.title]._cells.values() if c.data_type == 'e']
                result['sheets'].append(dict(name=ws.title, state=ws.sheet_state, dimension=ws.calculate_dimension(), actual_cells=len(cells), header_candidates=dict(rows), formulas=sum(c.data_type == 'f' for c in cells), formula_missing_cache=[c.coordinate for c in cells if c.data_type == 'f' and values[ws.title][c.coordinate].value is None], cached_errors=errors, merged=list(map(str, ws.merged_cells.ranges)), charts=len(ws._charts), images=len(ws._images), comments=sum(c.comment is not None for c in cells), external_formula_cells=[c.coordinate for c in cells if c.data_type == 'f' and '[' in str(c.value)]))
        elif suffix == '.docx':
            root = ET.fromstring(z.read('word/document.xml'))
            revisions = collections.Counter(x.tag.split('}')[-1] for x in root.iter() if x.tag in {W+'ins', W+'del', W+'moveFrom', W+'moveTo'})
            result.update(kind='word', revisions=dict(revisions), tables=sum(1 for _ in root.iter(W+'tbl')), media=[n for n in names if n.startswith('word/media/')])
            # 净稿候选，不改原文；只作为标题及段落抽取，不声称完整Word渲染。
            for parent in root.iter():
                for child in list(parent):
                    if child.tag in {W+'del', W+'moveFrom'}:
                        parent.remove(child)
            result['opening_text'] = [''.join(t.text or '' for t in p.iter(W+'t')) for p in root.iter(W+'p')][:35]
        elif suffix == '.pptx':
            slides = sorted([n for n in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)], key=lambda n:int(re.search(r'(\d+)\.xml', n).group(1)))
            result.update(kind='presentation', slides=[])
            for n in slides:
                text = [t.text or '' for t in ET.fromstring(z.read(n)).iter(A+'t')]
                result['slides'].append({'part': n, 'text': text})
            result['chart_parts'] = [n for n in names if re.fullmatch(r'ppt/charts/chart\d+\.xml', n)]
    return result

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('directory', type=Path)
    ap.add_argument('--out', required=True, type=Path)
    args = ap.parse_args()
    files = sorted(p for p in args.directory.iterdir() if p.is_file() and p.suffix.lower() in {'.xlsx','.xlsm','.docx','.pptx','.pdf'} and not p.name.startswith('~$'))
    if not files:
        ap.error('目录中未找到可盘点Office/PDF文件')
    if args.out.resolve() in {p.resolve() for p in files}:
        ap.error('输出不得覆盖输入文件')
    report = {'directory':str(args.directory.resolve()), 'files':[], 'limitations':['只读结构盘点，未重算公式、未验证公告事实、未执行视觉核验']}
    for p in files:
        try:
            report['files'].append(inspect(p))
        except Exception as ex:
            report['files'].append({'file':p.name, 'error':str(ex)})
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(report,ensure_ascii=False,indent=2), encoding='utf-8')
    errors = sum('error' in f for f in report['files'])
    print(json.dumps({'files':len(files),'parse_errors':errors,'out':str(args.out)},ensure_ascii=False))
    raise SystemExit(1 if errors else 0)

if __name__ == '__main__':
    main()
