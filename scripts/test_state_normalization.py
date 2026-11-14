#!/usr/bin/env python3
"""交易所项目动态状态映射回归测试（v4.3.7）。

回归背景：上交所 2026-09-30 批量状态为「中止（财报更新）」（全角括号），
而 SSE_STATE_MAP 键为半角「中止(财报更新)」，导致 10 条窗口内状态全部
落入 pending、不生成 suspended 事件，进而在审既有行不回写、汇总源数据少计。
"""
from __future__ import annotations
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract_filings import (BSE_STATE_MAP, SSE_STATE_MAP, SZSE_STATE_MAP,
                             LedgerBuilder, lookup_state, normalize_state)


class StateNormalizationTest(unittest.TestCase):
    def test_fullwidth_parentheses_maps_to_suspended(self):
        self.assertEqual(normalize_state('中止（财报更新）'), '中止(财报更新)')
        self.assertEqual(lookup_state('中止（财报更新）', SSE_STATE_MAP), 'suspended')

    def test_halfwidth_exact_still_works(self):
        self.assertEqual(lookup_state('中止(财报更新)', SSE_STATE_MAP), 'suspended')
        self.assertEqual(lookup_state('已回复', SSE_STATE_MAP), 'replied')

    def test_whitespace_and_fullwidth_punctuation(self):
        self.assertEqual(normalize_state(' 中止（ 财报更新 ） '), '中止(财报更新)')
        self.assertEqual(lookup_state(' 中止（ 财报更新 ） ', SSE_STATE_MAP), 'suspended')

    def test_termination_not_confused_with_suspension(self):
        self.assertEqual(lookup_state('终止(撤回)', SZSE_STATE_MAP), 'terminated')
        self.assertEqual(lookup_state('终止（撤回）', SZSE_STATE_MAP), 'terminated')
        self.assertEqual(lookup_state('终止', BSE_STATE_MAP), 'terminated')
        self.assertEqual(lookup_state('中止', BSE_STATE_MAP), 'suspended')

    def test_unknown_state_returns_none(self):
        self.assertIsNone(lookup_state('未定义的新状态', SSE_STATE_MAP))
        self.assertIsNone(lookup_state('中止及财报更新额外后缀', SSE_STATE_MAP))

    def test_from_sse_emits_events_and_no_unmapped_pending(self):
        rows = [
            {'公司代码': '600095', '公司简称': '湘财股份', '板块': '主板',
             '业务类型': '发行股份购买资产', '审核状态': '中止（财报更新）',
             '独立财务顾问': '财通证券股份有限公司', '更新日期': '2026-09-30',
             '受理日期': '2025-10-23'},
            {'公司代码': '603306', '公司简称': '华懋科技', '板块': '主板',
             '业务类型': '发行股份购买资产', '审核状态': '中止(财报更新)',
             '独立财务顾问': '申港证券股份有限公司', '更新日期': '2026-09-30',
             '受理日期': '2025-10-24'},
            {'公司代码': '688766', '公司简称': '普冉股份', '板块': '科创板',
             '业务类型': '发行股份购买资产', '审核状态': '已回复',
             '独立财务顾问': '国泰海通证券股份有限公司', '更新日期': '2026-09-28',
             '受理日期': '2026-06-01'},
            # 窗口外，不应进台账
            {'公司代码': '601969', '公司简称': '海南矿业', '板块': '主板',
             '业务类型': '发行股份购买资产', '审核状态': '已回复',
             '独立财务顾问': '', '更新日期': '2026-11-05',
             '受理日期': '2026-01-01'},
        ]
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            (d / 'sse_momentum_window.json').write_text(
                json.dumps(rows, ensure_ascii=False), encoding='utf-8')
            b = LedgerBuilder(d, '20260927', '20261003')
            b.from_sse()
            self.assertEqual(len(b.events), 3, f'events={b.events}')
            self.assertEqual(sum(e['type'] == 'suspended' for e in b.events), 2)
            self.assertEqual(sum(e['type'] == 'replied' for e in b.events), 1)
            unmapped = [p for p in b.pending if '审核状态未映射' in str(p.get('reason', ''))]
            self.assertEqual(unmapped, [], f'unexpected pending: {unmapped}')
            self.assertTrue(all(e['status'] == 'CONFIRMED' for e in b.events))


if __name__ == '__main__':
    unittest.main(verbosity=2)
