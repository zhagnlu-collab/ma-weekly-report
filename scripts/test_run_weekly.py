#!/usr/bin/env python3
import tempfile, shutil, unittest, zipfile
from pathlib import Path
from run_weekly import new_run, next_stage, set_stage, load_state, discover_template
ROOT=Path(__file__).resolve().parents[2]
PREV='并购周报基础数据-2026年第43期-20261025至20261031'

def fixture_readable(tpl):
 """OneDrive 云端占位会让 copy2 超时；夹具不可读时明确跳过，不伪装成代码错误。"""
 if not tpl.exists():return False
 files=[p for p in tpl.iterdir() if p.is_file() and p.suffix.lower() in ('.xlsx','.docx','.pptx')]
 if not files:return False
 for p in files:
  try:
   with zipfile.ZipFile(p) as z:
    if '[Content_Types].xml' not in z.namelist():return False
  except Exception:return False
 return True

class RunTest(unittest.TestCase):
 def test_discovers_latest_complete_template(self):
  tpl=ROOT/PREV
  if not fixture_readable(tpl):self.skipTest('local fixture unavailable (OneDrive placeholder)')
  self.assertEqual(discover_template(ROOT),tpl)
 def setUp(self):
  self.tmp=Path(tempfile.mkdtemp());self.tpl=ROOT/PREV
  if not fixture_readable(self.tpl):self.skipTest('local fixture unavailable (OneDrive placeholder)')
  class A:pass
 def test_date_derivation_and_idempotence(self):
  a=self.A();a.as_of='2026-11-07';a.template_dir=self.tpl;a.work_root=self.tmp;a.profile='production';a.config_sha='x';a.resume=False;a.allow_config_change=False;a.allow_cross_year=False
  s,p=new_run(a);self.assertEqual((s['period_start'],s['period_end']),('2026-11-01','2026-11-07'));self.assertEqual(next_stage(s),'ACQUIRE');self.assertTrue((p.parent/'0_working').is_dir())
  with self.assertRaises(Exception):new_run(a)

 def test_stage_order(self):
  a=self.A();a.as_of='2026-11-07';a.template_dir=self.tpl;a.work_root=self.tmp;a.profile='production';a.config_sha='x';a.resume=False;a.allow_config_change=False;a.allow_cross_year=False
  state,path=new_run(a)
  x=self.A();x.state=path;x.stage='EXTRACT';x.status='PASS';x.detail='';x.artifact=None
  with self.assertRaises(ValueError):set_stage(x)
  x.stage='ACQUIRE';set_stage(x);self.assertEqual(next_stage(load_state(path)),'EXTRACT')
 def tearDown(self):shutil.rmtree(self.tmp,ignore_errors=True)
if __name__=='__main__':unittest.main(verbosity=2)
