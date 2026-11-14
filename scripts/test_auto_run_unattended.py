#!/usr/bin/env python3
"""端到端离线回归：一次调用出件 + 同命令重跑幂等。

默认跳过（渲染 3 份 PDF 较慢）；设置环境变量后运行：

    MA_WEEKLY_E2E=1 python scripts/test_auto_run_unattended.py

不联网、不改真实期次目录；运行目录与成品目录都放在临时目录下。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
SKILL = SCRIPTS.parent
AS_OF = '2026-11-07'
OUT_NAME = '并购周报基础数据-2026年第44期-20261101至20261107'
PREV = '并购周报基础数据-2026年第43期-20261025至20261031'


def unreadable_template_files(tpl: Path) -> list[str]:
    """返回模板目录中不可读的原生 Office 文件（OneDrive 云端占位）。"""
    bad = []
    for p in sorted(tpl.iterdir()):
        if not p.is_file() or p.name.startswith('~$'):
            continue
        if p.suffix.lower() not in ('.xlsx', '.xlsm', '.docx', '.pptx'):
            continue
        try:
            with zipfile.ZipFile(p) as z:
                names = z.namelist()
            required = ('xl/workbook.xml' if p.suffix.lower().startswith('.xl')
                        else '[Content_Types].xml')
            if required not in names:
                bad.append(p.name)
        except Exception:
            bad.append(p.name)
    return bad


def workspace_root() -> Path | None:
    """沿 cwd 与 skill 路径向上找含完整第43期成果的真实工作区。"""
    seeds = [Path.cwd(), SKILL]
    for seed in seeds:
        for cand in (seed.resolve(), *list(seed.resolve().parents)[:6]):
            if (cand / '并购周报基础数据-2026年第43期-20261025至20261031').is_dir():
                return cand
    return None


@unittest.skipUnless(os.environ.get('MA_WEEKLY_E2E') == '1',
                     '设置 MA_WEEKLY_E2E=1 才执行完整离线回放')
class UnattendedE2ETest(unittest.TestCase):
    def test_offline_package_and_idempotent_rerun(self):
        root = workspace_root()
        self.assertIsNotNone(root, '未找到含第43期完整成果的工作区')
        bad = unreadable_template_files(root / PREV)
        if bad:
            # 现场模板被 OneDrive 回收为云端占位时，完整离线回放在物理上不可执行。
            # 不假装通过：先断言主链必须失败关闭、不得发布半成品，再如实跳过。
            with tempfile.TemporaryDirectory() as td:
                temp = Path(td)
                cp = subprocess.run(
                    [sys.executable, str(SCRIPTS / 'auto_run.py'),
                     '--as-of', AS_OF, '--root', str(root),
                     '--work-root', str(temp / 'runs'),
                     '--out-root', str(temp / 'out'), '--plan-only'],
                    cwd=str(root), capture_output=True, text=True, timeout=180)
                self.assertNotEqual(cp.returncode, 0, cp.stdout)
                summary = json.loads((temp / 'runs' / AS_OF / 'auto-run-summary.json')
                                     .read_text(encoding='utf-8'))
                self.assertEqual(summary['status'], 'BLOCKED')
                self.assertIn('TEMPLATE_UNREADABLE', summary.get('error', ''))
                self.assertFalse((temp / 'out').exists(), '阻断时不得发布半成品')
            self.skipTest(
                f'第43期模板 {len(bad)} 份文件为 OneDrive 云端占位，'
                '完整离线回放不可执行；已验证 TEMPLATE_UNREADABLE 失败关闭')
        with tempfile.TemporaryDirectory() as td:
            temp = Path(td)
            runs, out = temp / 'runs', temp / 'out'
            cmd = [sys.executable, str(SCRIPTS / 'auto_run.py'),
                   '--as-of', AS_OF, '--root', str(root),
                   '--work-root', str(runs), '--out-root', str(out),
                   '--offline', '--retries', '0', '--timeout', '900']

            first = subprocess.run(cmd, cwd=str(root), capture_output=True,
                                   text=True, timeout=1800)
            self.assertEqual(first.returncode, 0, first.stdout[-3000:] + first.stderr[-3000:])
            run_dir = runs / AS_OF
            summary = json.loads((run_dir / 'auto-run-summary.json').read_text(encoding='utf-8'))
            self.assertIn(summary['status'], ('COMPLETE', 'COMPLETE_WITH_EXCEPTIONS'))
            self.assertTrue(summary.get('published'), '成品未原子发布')
            final = Path(summary['final'])
            self.assertTrue(final.is_dir())
            natives = [p for p in final.iterdir()
                       if p.suffix.lower() in ('.xlsx', '.xlsm', '.docx', '.pptx')]
            pdfs = list(final.glob('*.pdf'))
            self.assertEqual(len(natives), 16, f'原生文件数={len(natives)}')
            self.assertEqual(len(pdfs), 3, f'PDF 数={len(pdfs)}')
            self.assertTrue((final / 'artifact-manifest.json').is_file())
            self.assertTrue((run_dir / '本期编制与待核说明.md').is_file())
            self.assertFalse((run_dir / 'computer-use-request.json').exists(),
                             '默认链路不得产生待接管请求')
            validation = json.loads((run_dir / 'validation-report.json').read_text(encoding='utf-8'))
            self.assertEqual(validation['hard_failures'], [], '硬门禁不应失败')

            second = subprocess.run(cmd, cwd=str(root), capture_output=True,
                                    text=True, timeout=900)
            self.assertEqual(second.returncode, 0, second.stdout[-2000:] + second.stderr[-2000:])
            state = json.loads((run_dir / 'run-state.json').read_text(encoding='utf-8'))
            for stage, info in state['stages'].items():
                self.assertEqual(info['status'], 'OK', f'{stage}: {info}')
                self.assertIn('skipped_at', info, f'{stage} 重跑未命中幂等跳过')
            self.assertEqual(len(list(out.glob(OUT_NAME + '*'))), 1,
                             '重跑产生了重复发布目录')


if __name__ == '__main__':
    unittest.main(verbosity=2)
