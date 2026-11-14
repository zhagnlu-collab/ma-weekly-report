#!/usr/bin/env python3
"""在审表多板块坐标补丁合并回归测试（v4.3.7）。

回归背景：沪主板与科创板都映射到「上交所在审」，excel_plan 分两批处理，
原实现 `updates[sheet] = cell_updates` 使科创板批次覆盖沪主板批次。
"""
from __future__ import annotations
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_package import merge_sheet_patches


class MergeSheetPatchesTest(unittest.TestCase):
    def test_two_batches_same_sheet_are_merged(self):
        store: dict[str, dict[str, str]] = {}
        merge_sheet_patches(store, '上交所在审', {'E67': '中止', 'K67': '2026-09-30'})
        merge_sheet_patches(store, '上交所在审', {'E71': '中止', 'K71': '2026-09-30'})
        self.assertEqual(store['上交所在审'],
                         {'E67': '中止', 'K67': '2026-09-30',
                          'E71': '中止', 'K71': '2026-09-30'})

    def test_different_sheets_are_independent(self):
        store: dict[str, dict[str, str]] = {}
        merge_sheet_patches(store, '上交所在审', {'E67': '中止'})
        merge_sheet_patches(store, '深交所在审', {'F100': '是'})
        self.assertEqual(set(store), {'上交所在审', '深交所在审'})
        self.assertEqual(store['深交所在审'], {'F100': '是'})

    def test_same_coordinate_later_batch_wins(self):
        store: dict[str, dict[str, str]] = {}
        merge_sheet_patches(store, '上交所在审', {'E67': '已问询'})
        merge_sheet_patches(store, '上交所在审', {'E67': '中止'})
        self.assertEqual(store['上交所在审']['E67'], '中止')

    def test_empty_patch_does_not_create_sheet(self):
        store: dict[str, dict[str, str]] = {}
        merge_sheet_patches(store, '上交所在审', {})
        self.assertEqual(store, {})

    def test_three_batches_simulating_four_boards(self):
        store: dict[str, dict[str, str]] = {}
        for coords in ({'E1': 'a'}, {'E2': 'b'}, {'E3': 'c'}):
            merge_sheet_patches(store, '上交所在审', coords)
        self.assertEqual(len(store['上交所在审']), 3)


if __name__ == '__main__':
    unittest.main(verbosity=2)
