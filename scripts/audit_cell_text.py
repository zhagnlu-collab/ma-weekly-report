#!/usr/bin/env python3
"""单元格展示文本质量审计（v4.3.8）。

拆封后只读工具。背景：compare_excel.py 以「行键/日期集合/非空列数」为口径，
只能看到行是否多/少，看不到同一行里「交易概要」是 1200 字公告正文还是
有界摘要。第 39→40 期实测：净化 14 处公告正文后，行级 totals 一个数字都没变。
本工具补足单元格级指标，供以后每轮盲测记录「改动前后」的真实效果。

用法：
    python scripts/audit_cell_text.py --generated <dir> --reference <dir> --out <file.json>
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import openpyxl  # noqa: E402
from lib import celltext  # noqa: E402

LONG_THRESHOLDS = (200, 500, 1000)


def workbook_paths(directory: Path) -> list[Path]:
    return sorted(p for p in directory.glob('*.xlsx')
                  if not p.name.startswith('~$') and p.is_file())


def audit_dir(directory: Path) -> dict:
    files = workbook_paths(directory)
    if not files:
        raise ValueError(f'目录中没有可读 xlsx: {directory}')
    raw_hits = celltext.scan_raw_filing_cells(files)
    lengths: list[int] = []
    longest: list[dict] = []
    for path in files:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        value = cell.value
                        if not isinstance(value, str) or len(value) < LONG_THRESHOLDS[0]:
                            continue
                        lengths.append(len(value))
                        longest.append({'file': path.name, 'sheet': ws.title,
                                        'cell': cell.coordinate, 'length': len(value),
                                        'head': value[:60]})
        finally:
            wb.close()
    longest.sort(key=lambda x: -x['length'])
    return {
        'directory': str(directory),
        'files': len(files),
        'raw_filing_body_cells': len(raw_hits),
        'raw_filing_body_examples': raw_hits[:8],
        'long_text_cells': len(lengths),
        'over_thresholds': {str(t): sum(1 for n in lengths if n > t)
                            for t in LONG_THRESHOLDS},
        'max_cell_length': max(lengths) if lengths else 0,
        'top_longest': longest[:8],
        'status': 'PASS' if not raw_hits else 'BLOCKED',
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--generated', type=Path, required=True)
    ap.add_argument('--reference', type=Path, required=True)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--before', type=Path, default=None,
                    help='可选：修复前生成稿，用于记录改动前后量化指标')
    args = ap.parse_args()
    report = {'rule': '单元格文本长度>200 且前80字命中>=2个公告抬头特征',
              'generated': audit_dir(args.generated),
              'reference': audit_dir(args.reference)}
    if args.before:
        before = audit_dir(args.before)
        report['before'] = before
        report['delta'] = {
            'raw_filing_body_cells': (before['raw_filing_body_cells'],
                                      report['generated']['raw_filing_body_cells']),
            'long_text_cells': (before['long_text_cells'],
                                report['generated']['long_text_cells']),
            'max_cell_length': (before['max_cell_length'],
                                report['generated']['max_cell_length']),
        }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'generated_raw_body': report['generated']['raw_filing_body_cells'],
                      'reference_raw_body': report['reference']['raw_filing_body_cells'],
                      'before_raw_body': report.get('before', {}).get('raw_filing_body_cells'),
                      'out': str(args.out)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
