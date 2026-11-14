#!/usr/bin/env python3
"""单元格展示文本安全层回归测试（v4.3.8）。

回归背景：第 39→40 期 Excel 盲测发现 14 个单元格被塞入整段公告正文
（最长 1200 字，以「证券代码：…公告编号：…公司及董事会全体成员保证…
重要内容提示」开头），现场稿 0 处。根因是 evidence 层（pdf_facts 的前两页
原文 summary）与 display 层（_row_values 写入交易概要）未分离。
"""
from __future__ import annotations
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import celltext  # noqa: E402
from build_package import _row_values  # noqa: E402

RAW_000635 = (
    '— 1 —\n证券代码：000635 证券简称：英力特 公告编号：2026-059\n'
    '宁夏英力特化工股份有限公司\n关于控股股东拟协议转让公司股份\n公开征集受让方的公告\n'
    '公司及董事会全体成员保证公告内容的真实、准确和完整，没有虚假记载、\n'
    '误导性陈述或者重大遗漏。\n重要内容提示：\n\uf06c 转让价格及比例\n'
    '国能英力特能源化工集团股份有限公司（以下简称"英力特集团"）拟通过\n'
    '公开征集受让方的方式一次性整体协议转让其持有的宁夏英力特化工股份有限\n'
    '公司（以下简称"公司"或"上市公司"或"英力特化工"）155,322,687 股无限售\n'
    '流通股份（占公司总股本的 39.41%），转让价格不低于人民币 6.27 元/股。\n'
    '\uf06c 控制权变更\n若本次公开征集转让实施完成，公司控制权将发生变更。'
)
RAW_603682 = (
    '证券代码：603682 证券简称：锦和商管 公告编号：2026-035\n'
    '上海锦和商业经营管理（集团）股份有限公司\n关于公司控股股东协议转让部分股份的公告\n'
    '本公司董事会、全体董事及相关股东保证本公告内容不存在任何虚假记载、\n'
    '误导性陈述或者重大遗漏，并对其内容的真实性、准确性和完整性承担法律责任。\n'
    '重要内容提示：\n\uf06c 上海锦和商业经营管理（集团）股份有限公司（以下简称"锦和商管"\n'
    '或"公司"）控股股东上海锦和投资集团有限公司（以下简称"锦和集团"）与杭\n'
    '州汇皋企业管理合伙企业（有限合伙）（以下简称"杭州汇皋"）签署了《股份转\n'
    '让协议》，锦和集团通过协议转让的方式将其持有的锦和商管 70,874,000 股（约\n'
    '占锦和商管总股本的 14.9998%）无限售流通股转让给杭州汇皋。'
)
# 现场人工撰写的长摘要：引用了证券代码，但只命中 1 个抬头特征，不得误判
HUMAN_LONG = (
    '本次南京三超新材料股份有限公司（证券代码：300554，以下简称"公司"）'
    '拟发生的控制权变动，主要涉及控股股东邹余耀、持股5%以上股东刘建勋'
    '与相关受让方签署股份转让协议，转让完成后公司实际控制人将发生变更，'
    '受让方承诺自股份过户登记完成之日起36个月内不转让所持股份，'
    '本次交易尚需通过深圳证券交易所合规性确认并办理过户登记手续，'
    '最终能否实施完成及完成时间尚存在不确定性。' * 2
)


class RawFilingBodyDetection(unittest.TestCase):
    def test_detects_real_announcement_bodies(self):
        for raw in (RAW_000635, RAW_603682):
            self.assertTrue(celltext.is_raw_filing_body(raw), raw[:40])

    def test_human_long_summary_not_flagged(self):
        self.assertFalse(celltext.is_raw_filing_body(HUMAN_LONG))

    def test_short_text_and_non_string_not_flagged(self):
        self.assertFalse(celltext.is_raw_filing_body('证券代码：000635 证券简称：英力特'))
        self.assertFalse(celltext.is_raw_filing_body(12345))
        self.assertFalse(celltext.is_raw_filing_body(None))


class FilingBriefExtraction(unittest.TestCase):
    def test_brief_is_bounded_and_boilerplate_free(self):
        for raw in (RAW_000635, RAW_603682):
            b = celltext.filing_brief(raw)
            self.assertGreaterEqual(len(b), 25)
            self.assertLessEqual(len(b), celltext.MAX_BRIEF)
            self.assertNotIn('证券代码', b)
            self.assertNotIn('公告编号', b)
            self.assertNotIn('误导性陈述', b)
            self.assertFalse(celltext.is_raw_filing_body(b))

    def test_brief_keeps_transaction_facts(self):
        b = celltext.filing_brief(RAW_000635)
        self.assertIn('英力特', b)
        self.assertIn('155,322,687', b)


class SafeCellValue(unittest.TestCase):
    def test_raw_body_never_reaches_cell(self):
        out = celltext.safe_cell_value(RAW_000635, fallback='关于控股股东拟协议转让公司股份公开征集受让方的公告')
        self.assertFalse(celltext.is_raw_filing_body(out))
        self.assertLessEqual(len(out), celltext.MAX_BRIEF)
        self.assertNotIn('证券代码', out)

    def test_falls_back_to_title_when_no_brief(self):
        out = celltext.safe_cell_value(
            '证券代码：000635 证券简称：英力特 公告编号：2026-059\n' + '保证。' * 80,
            fallback='关于控股股东拟协议转让公司股份的公告')
        self.assertEqual(out, '关于控股股东拟协议转让公司股份的公告')

    def test_clean_short_text_unchanged(self):
        self.assertEqual(celltext.safe_cell_value('进行中'), '进行中')
        self.assertEqual(celltext.safe_cell_value(12345.6), 12345.6)

    def test_clean_locator_strips_source_prefix(self):
        self.assertEqual(
            celltext.clean_locator('巨潮公告｜2026-10-10｜关于股份转让的公告'),
            '关于股份转让的公告')


class NameNormalization(unittest.TestCase):
    def test_spaced_cjk_name(self):
        self.assertEqual(celltext.normalize_name('英 力 特'), '英力特')
        self.assertEqual(celltext.normalize_name('ST 智知'), 'ST 智知')
        self.assertEqual(celltext.normalize_name(' 锦和商管 '), '锦和商管')


class RowValuesIntegration(unittest.TestCase):
    """真实写入路径：整段公告正文不得出现在任何单元格值里。"""

    def setUp(self):
        self.cols = {'name': 'C', 'code': 'B', 'date': 'D', 'summary': 'G',
                     'title': 'F', 'progress': 'E', 'note': 'H'}

    def test_polluted_summary_is_sanitized(self):
        case = {'name': '英 力 特', 'code': '000635.SZ', 'board': '深主板',
                'facts': {'summary': RAW_000635}}
        event = {'type': 'first_disclosed', 'date': '2026-10-10', 'note': '标题规则'}
        source = {'locator': '巨潮公告｜2026-10-10｜关于控股股东拟协议转让公司股份公开征集受让方的公告'}
        row = _row_values(event, case, source, self.cols)
        self.assertEqual(row['C'], '英力特')
        for col, value in row.items():
            self.assertFalse(celltext.is_raw_filing_body(value), f'{col}: {str(value)[:60]}')
            if isinstance(value, str) and col in ('G', 'F'):
                self.assertLessEqual(len(value), celltext.MAX_BRIEF_FALLBACK)

    def test_human_long_summary_is_preserved(self):
        """现场人工长摘要不应被净化层改写（保真优先）。"""
        out = celltext.safe_cell_value(HUMAN_LONG)
        self.assertEqual(out, HUMAN_LONG)


class WorkbookScanGate(unittest.TestCase):
    def test_scan_detects_only_polluted_cells(self):
        import openpyxl
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / 'x.xlsx'
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = '交易概要'
            ws['A1'] = RAW_000635
            ws['A2'] = HUMAN_LONG
            ws['A3'] = '进行中'
            wb.save(p)
            hits = celltext.scan_raw_filing_cells([p])
            self.assertEqual(len(hits), 1)
            self.assertEqual(hits[0]['cell'], 'A1')
            self.assertEqual(hits[0]['sheet'], '交易概要')


if __name__ == '__main__':
    unittest.main(verbosity=2)
