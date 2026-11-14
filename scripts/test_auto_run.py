#!/usr/bin/env python3
"""v4.0.0 无人值守链路快速回归测试（不联网、不渲染 PDF、不改真实期次目录）。

覆盖：期次/区间/文件名推导、跨 run OOXML 安全替换、状态幂等、
配置与工作区根自动探测、质量门禁辅助判定，以及 --plan-only 单日期入口。

    python scripts/test_auto_run.py
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

from lib import ooxml, period as P  # noqa: E402
from lib.state import RunState, digest_bytes  # noqa: E402
import quality_gate as QG  # noqa: E402

AS_OF = '2026-11-07'
PREV = '并购周报基础数据-2026年第43期-20261025至20261031'
EXPECTED_OUT = '并购周报基础数据-2026年第44期-20261101至20261107'
SPAN_43 = '20261025至20261031'
SPAN_44 = '20261101至20261107'


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '</Types>')


def write_minimal_office(path: Path) -> None:
    """写入最小但可读的 OOXML 包。

    v4.3.6 起 PLAN 阶段会做模板可读性预检，空文件（旧夹具用 touch()）
    与 OneDrive 云端占位同样会被判为 TEMPLATE_UNREADABLE。夹具必须是真包。
    """
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        if path.suffix.lower() in ('.xlsx', '.xlsm'):
            z.writestr('xl/workbook.xml',
                       '<?xml version="1.0"?>'
                       '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                       '<sheets/></workbook>')
        elif path.suffix.lower() == '.docx':
            z.writestr('word/document.xml',
                       '<?xml version="1.0"?>'
                       '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                       '<w:body/></w:document>')
        else:
            z.writestr('ppt/presentation.xml',
                       '<?xml version="1.0"?>'
                       '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')


def make_fixture(root: Path) -> Path:
    """造一个角色齐全、文件均为最小可读 OOXML 包的上期目录。"""
    d = root / PREV
    d.mkdir(parents=True)
    for i in [0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13]:
        write_minimal_office(d / f'{i}、底表{SPAN_43}.xlsx')
    write_minimal_office(d / f'5、2026年并购重组周报_{SPAN_43}.docx')
    write_minimal_office(d / f'A股市场并购重组动态-{SPAN_43}.docx')
    write_minimal_office(d / f'并购重组市场周报_{SPAN_43}.pptx')
    return d


class PeriodTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        make_fixture(self.root)

    def tearDown(self):
        self.tmp.cleanup()

    def test_single_date_derives_everything(self):
        p = P.resolve(self.root, AS_OF)
        self.assertEqual(p.start, '20261101')
        self.assertEqual(p.end, '20261107')
        self.assertEqual(p.issue_no, 44)
        self.assertEqual(p.span_token, SPAN_44)
        self.assertEqual(p.out_dirname, EXPECTED_OUT)
        self.assertEqual(Path(p.prev_dir).name, PREV)
        self.assertEqual(p.notes, [])

    def test_window_starts_day_after_prev_end_not_asof_minus_6(self):
        p = P.resolve(self.root, '2026-11-05')
        self.assertEqual(p.start, '20261101')
        self.assertEqual(p.end, '20261105')
        self.assertTrue(any('非现场惯例' in n for n in p.notes))

    def test_latest_incomplete_period_wins_over_older_complete_period(self):
        older = self.root / '并购周报基础数据-2026年第41期-20261011至20261017'
        latest = self.root / '并购周报基础数据-2026年第42期-20261018至20261024'
        older.mkdir();latest.mkdir()
        # only one native role is enough for an incomplete prior package
        (older / '0、底表_20261011至20261017.xlsx').touch()
        (latest / '0、底表_20261018至20261024.xlsx').touch()
        p = P.resolve(self.root, '2026-10-31')
        self.assertEqual(Path(p.prev_dir).name, latest.name)
        self.assertEqual(p.start, '20261025')
        self.assertEqual(p.issue_no, 43)
        self.assertTrue(any('缺角色' in n for n in p.notes))

    def test_historical_year_issue_numbers_bump_independently(self):
        p = P.resolve(self.root, AS_OF)
        self.assertEqual(
            P.map_filename('9、2025年第55期收购案例库_20261025至20261031.xlsx', p),
            '9、2025年第56期收购案例库_20261101至20261107.xlsx')
        self.assertEqual(
            P.map_filename('10、2024年第119期-20261025至20261031.xlsx', p),
            '10、2024年第120期-20261101至20261107.xlsx')

    def test_replacement_anchors_are_safe(self):
        p = P.resolve(self.root, AS_OF)
        r = P.text_replacements(p)
        self.assertEqual(r[SPAN_43], SPAN_44)
        self.assertEqual(r['2026年10月25日至2026年10月31日'],
                         '2026年11月1日至2026年11月7日')
        self.assertEqual(r['截至10月31日'], '截至11月7日')
        self.assertEqual(r['2026年第43期'], '2026年第44期')
        # 业务正文中的历史案例日期不得进入替换表
        self.assertNotIn('2026年10月30日', r)


class OoxmlTests(unittest.TestCase):
    def test_cross_run_replacement_has_no_stale_tail(self):
        # 「截至2026年10月31日」被拆在 3 个 <t> 节点里，替换后不得残留旧日期
        data = ('<w:p><w:r><w:t>截至2026年</w:t></w:r>'
                '<w:r><w:t>10月31日，在审项目共</w:t></w:r>'
                '<w:r><w:t>49单</w:t></w:r></w:p>').encode()
        out, hits = ooxml._patch_text_nodes(data, {'截至2026年10月31日': '截至2026年11月7日'})
        text = ''.join(m.group(2) for m in ooxml.TEXT_NODE_RE.finditer(out.decode()))
        self.assertEqual(hits['截至2026年10月31日'], 1)
        self.assertIn('截至2026年11月7日', text)
        self.assertNotIn('10月31日', text)
        self.assertIn('49单', text)

    def test_regex_key_and_unrelated_text_untouched(self):
        data = ('<w:p><w:t>并购重组周报</w:t><w:t>第43期</w:t>'
                '<w:t>2025年第55期</w:t></w:p>').encode('utf-8')
        out, hits = ooxml._patch_text_nodes(
            data, {'re:(?<!\\d{4}年)第43期': '第44期'})
        text = ''.join(m.group(2) for m in ooxml.TEXT_NODE_RE.finditer(out.decode()))
        self.assertIn('第44期', text)
        self.assertIn('2025年第55期', text)  # 历史年度期号不得被误改
        self.assertEqual(sum(hits.values()), 1)

    def test_no_replacements_returns_identical_bytes(self):
        data = b'<w:t>keep me</w:t>'
        out, hits = ooxml._patch_text_nodes(data, {})
        self.assertEqual(out, data)
        self.assertEqual(hits, {})


class StateTests(unittest.TestCase):
    def test_digest_accepts_str_and_bytes(self):
        self.assertEqual(digest_bytes('abc'), digest_bytes(b'abc'))
        self.assertEqual(len(digest_bytes('abc')), 64)

    def test_stage_skip_is_idempotent(self):
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / 'run-state.json'
            s = RunState(f, AS_OF, extra={'root': td})
            self.assertFalse(s.should_skip('PLAN', 'k1'))
            s.start('PLAN', 'k1')
            s.ok('PLAN', 'k1', out_dirname=EXPECTED_OUT)
            self.assertTrue(RunState(f, AS_OF).should_skip('PLAN', 'k1'))
            self.assertFalse(RunState(f, AS_OF).should_skip('PLAN', 'k2'))
            s.fail('BUILD', 'b1', 'boom')
            self.assertFalse(RunState(f, AS_OF).should_skip('BUILD', 'b1'))

    def test_detail_cannot_overwrite_reserved_status(self):
        """VALIDATE 阶段会回传门禁状态，不能把阶段自身的 OK 覆盖掉。"""
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / 'run-state.json'
            s = RunState(f, AS_OF)
            s.start('VALIDATE', 'k')
            s.ok('VALIDATE', 'k', validation_status='COMPLETE_WITH_EXCEPTIONS')
            info = RunState(f, AS_OF).stage('VALIDATE')
            self.assertEqual(info['status'], 'OK')
            self.assertEqual(info['validation_status'], 'COMPLETE_WITH_EXCEPTIONS')
            self.assertTrue(RunState(f, AS_OF).should_skip('VALIDATE', 'k'))


class QualityGateHelperTests(unittest.TestCase):
    def test_office_manifest_default_is_not_provided(self):
        with tempfile.TemporaryDirectory() as td:
            self.assertEqual(QG.office_manifest_status(Path(td))['status'], 'NOT_PROVIDED')
            d = Path(td) / 'office-pdf'
            d.mkdir()
            (d / 'office-export-manifest.json').write_text(
                json.dumps({'status': 'COMPLETE', 'exports': [1, 2, 3]}), encoding='utf-8')
            self.assertEqual(QG.office_manifest_status(Path(td))['status'], 'COMPLETE')

    def test_role_mapping(self):
        self.assertEqual(QG.role('5、2026年并购重组周报_20261101至20261107.docx'), 'weekly_docx')
        self.assertEqual(QG.role('A股市场并购重组动态-20261101至20261107.docx'), 'dynamic_docx')
        self.assertEqual(QG.role('并购重组市场周报_20261101至20261107.pptx'), 'market_ppt')
        self.assertEqual(QG.role('9、底表.xlsx'), 'xlsx_9')
        self.assertIsNone(QG.role('本期编制与待核说明.md'))


class EntryPointTests(unittest.TestCase):
    def test_plan_only_single_date(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            make_fixture(root)
            cp = subprocess.run(
                [sys.executable, str(SCRIPTS / 'auto_run.py'), '--as-of', AS_OF,
                 '--root', str(root), '--work-root', str(root / 'runs'),
                 '--out-root', str(root / 'out'), '--plan-only'],
                capture_output=True, text=True, timeout=90)
            self.assertEqual(cp.returncode, 0, cp.stdout + cp.stderr)
            summary = json.loads((root / 'runs' / AS_OF / 'auto-run-summary.json')
                                 .read_text(encoding='utf-8'))
            self.assertEqual(summary['status'], 'PLAN_ONLY')
            self.assertEqual(summary['plan']['period']['out_dirname'], EXPECTED_OUT)
            self.assertEqual(summary['plan']['timezone'], 'Asia/Shanghai')
            self.assertEqual((root / 'runs' / AS_OF / 'prev-fingerprints.json').exists(), True)

    def test_default_entrypoint_never_waits_for_computer_use(self):
        src = (SCRIPTS / 'auto_run.py').read_text(encoding='utf-8')
        self.assertNotIn('NEEDS_COMPUTER_USE', src)
        self.assertNotIn('computer-use-request.json', src)
        self.assertIn('--as-of', src)
        self.assertIn("'PUBLISH'", src)
        gate = (SCRIPTS / 'quality_gate.py').read_text(encoding='utf-8')
        self.assertIn('office_manifest_status', gate)
        self.assertIn("'BLOCKED' if hard_fail", gate)


if __name__ == '__main__':
    unittest.main(verbosity=2)
