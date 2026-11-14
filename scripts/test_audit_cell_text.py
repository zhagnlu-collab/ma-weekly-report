#!/usr/bin/env python3
"""v4.3.8 单元格级审计工具回归测试。"""
from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import openpyxl  # noqa: E402
from audit_cell_text import audit_dir, main  # noqa: E402

RAW = ('证券代码：000635 证券简称：英力特 公告编号：2026-059 '
       '本公司及董事会全体成员保证信息披露内容真实、准确、完整，'
       '没有虚假记载、误导性陈述或者重大遗漏。重要内容提示：'
       '国能英力特能源化工集团股份有限公司拟通过公开征集受让方方式'
       '转让其持有的宁夏英力特化工股份有限公司155,322,687股股份，'
       '占公司总股本的39.41%，转让价格不低于6.27元/股。' * 3)
HUMAN_LONG = ('本次南京三超新材料股份有限公司（证券代码：300554）控股股东'
              '拟以协议转让方式向受让方转让其所持部分股份，交易完成后公司'
              '控制权将发生变更，受让方承诺36个月内不转让所受让股份，'
              '相关款项分期支付，尚需取得交易所合规性确认。' * 4)


def make_book(path: Path, texts: list[str]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '测试表'
    ws['A1'] = '标题'
    for i, text in enumerate(texts, start=2):
        ws.cell(i, 1, text)
    wb.save(path)
    wb.close()


class AuditCellTextTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.gen = self.root / 'gen'
        self.ref = self.root / 'ref'
        self.before = self.root / 'before'
        for d in (self.gen, self.ref, self.before):
            d.mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def test_detects_raw_filing_body(self):
        make_book(self.before / 'a.xlsx', [RAW, RAW])
        result = audit_dir(self.before)
        self.assertEqual(result['raw_filing_body_cells'], 2)
        self.assertEqual(result['status'], 'BLOCKED')
        self.assertGreater(result['over_thresholds']['500'], 0)

    def test_human_long_summary_not_flagged(self):
        make_book(self.ref / 'a.xlsx', [HUMAN_LONG])
        result = audit_dir(self.ref)
        self.assertEqual(result['raw_filing_body_cells'], 0)
        self.assertEqual(result['status'], 'PASS')
        # 合法长摘要仍计入「长文本」分布，但绝不能计入「公告正文污染」。
        self.assertGreaterEqual(result['long_text_cells'], 1)

    def test_cli_records_before_after_delta(self):
        make_book(self.before / 'a.xlsx', [RAW])
        make_book(self.gen / 'a.xlsx', ['有界摘要：拟协议转让39.41%股份'])
        make_book(self.ref / 'a.xlsx', [HUMAN_LONG[:300]])
        out = self.root / 'cell-quality.json'
        argv = ['audit_cell_text.py', '--generated', str(self.gen),
                '--reference', str(self.ref), '--before', str(self.before),
                '--out', str(out)]
        old = sys.argv
        sys.argv = argv
        try:
            self.assertEqual(main(), 0)
        finally:
            sys.argv = old
        report = json.loads(out.read_text(encoding='utf-8'))
        self.assertEqual(report['delta']['raw_filing_body_cells'], [1, 0])
        self.assertEqual(report['generated']['status'], 'PASS')

    def test_empty_directory_raises(self):
        with self.assertRaisesRegex(ValueError, '没有可读'):
            audit_dir(self.gen)


if __name__ == '__main__':
    unittest.main(verbosity=2)
