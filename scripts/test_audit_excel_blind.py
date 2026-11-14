#!/usr/bin/env python3
"""v4.3.6 回归：参考稿不可读必须失败关闭，不得静默产出对比结论。

背景：2026-11-08 Round 5（43→44）中 `第44期周报/` 12 份 xlsx 全为
OneDrive 云端占位；旧版 collect() 用 `except: continue` 静默跳过，
参考侧集合为空，审计报告错误显示「所有代码都是生成独有」。
"""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from audit_excel_blind import (collect, preflight_workbooks, role_map)


class PreflightTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix='audit-blind-'))
        self.good = self.tmp / 'generated'
        self.ref = self.tmp / 'reference'
        self.good.mkdir()
        self.ref.mkdir()
        wb = Workbook()
        ws = wb.active
        ws.title = '测试表'
        ws.append(['股票代码', '首次披露时间', '备注'])
        ws.append(['600500', '2026-11-03', '受理'])
        wb.save(self.good / '0、并购周报基础数据_test.xlsx')
        wb.save(self.ref / '0、并购周报基础数据_test.xlsx')

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_readable_workbooks_pass_preflight(self):
        g = preflight_workbooks(self.good, 'generated')
        r = preflight_workbooks(self.ref, 'reference')
        self.assertTrue(g['ok'])
        self.assertTrue(r['ok'])
        self.assertEqual((g['xlsx_count'], g['readable']), (1, 1))
        self.assertEqual(r['unreadable'], [])

    def test_cloud_placeholder_is_detected(self):
        # OneDrive 占位：文件名/尺寸看似正常，但不是可读 OOXML 包
        placeholder = self.ref / '7、2026年控制权变更_test.xlsx'
        placeholder.write_bytes(b'\x00' * 4096)
        r = preflight_workbooks(self.ref, 'reference')
        self.assertFalse(r['ok'])
        self.assertEqual(r['xlsx_count'], 2)
        self.assertEqual(r['readable'], 1)
        self.assertEqual(len(r['unreadable']), 1)
        self.assertIn('7、2026年控制权变更', r['unreadable'][0]['file'])

    def test_empty_reference_dir_is_blocked(self):
        for p in self.ref.glob('*.xlsx'):
            p.unlink()
        r = preflight_workbooks(self.ref, 'reference')
        self.assertFalse(r['ok'], '空参考目录不得视为通过')
        self.assertEqual(r['xlsx_count'], 0)

    def test_collect_raises_instead_of_silent_skip(self):
        (self.ref / '7、2026年控制权变更_test.xlsx').write_bytes(b'not a zip')
        with self.assertRaises(RuntimeError) as ctx:
            collect(self.ref, '20261101', '20261107')
        self.assertIn('工作簿不可读', str(ctx.exception))

    def test_role_map_ignores_lock_files(self):
        (self.ref / '~$0、并购周报基础数据_test.xlsx').write_bytes(b'lock')
        self.assertEqual(list(role_map(self.ref)), ['0'])


class FailClosedCliTest(unittest.TestCase):
    def test_main_returns_2_and_writes_blocked_result(self):
        import subprocess
        import sys
        tmp = Path(tempfile.mkdtemp(prefix='audit-blind-cli-'))
        try:
            good, ref, out = tmp / 'gen', tmp / 'ref', tmp / 'result.json'
            good.mkdir()
            ref.mkdir()
            wb = Workbook()
            wb.active['A1'] = 'x'
            wb.save(good / '0、基础数据_test.xlsx')
            (ref / '0、基础数据_test.xlsx').write_bytes(b'\x00' * 128)
            script = Path(__file__).with_name('audit_excel_blind.py')
            cp = subprocess.run(
                [sys.executable, str(script),
                 '--generated', str(good), '--reference', str(ref),
                 '--start', '20261101', '--end', '20261107', '--out', str(out)],
                capture_output=True, text=True, timeout=60)
            self.assertEqual(cp.returncode, 2, cp.stderr)
            data = json.loads(out.read_text(encoding='utf-8'))
            self.assertEqual(data['status'], 'BLOCKED')
            self.assertEqual(data['verdict'], 'REFERENCE_UNREADABLE')
            self.assertNotIn('code_set', data)
            self.assertNotIn('summary_cache_diff', data)
            self.assertTrue(out.with_suffix('.md').exists())
            self.assertIn('BLOCKED', out.with_suffix('.md').read_text(encoding='utf-8'))
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == '__main__':
    unittest.main(verbosity=2)
