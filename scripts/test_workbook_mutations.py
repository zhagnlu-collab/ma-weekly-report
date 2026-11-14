#!/usr/bin/env python3
import json,tempfile,unittest
from pathlib import Path
import openpyxl
from apply_workbook_mutations import mutate_file
from lib import ooxml
class M(unittest.TestCase):
 def setUp(self):
  self.td=tempfile.TemporaryDirectory();self.p=Path(self.td.name)/'0、test.xlsx';wb=openpyxl.Workbook();ws=wb.active;ws.title='Data';ws.append(['代码','简称','金额（万元）','公告日期']);ws.append(['000001','甲',1,'2026-01-01']);wb.save(self.p)
 def tearDown(self):self.td.cleanup()
 def test_update_and_append(self):
  a=[];mutate_file(self.p,[{'sheet':'Data','header_row':1,'key_columns':['代码'],'values':{'代码':'000001','简称':'乙'},'source_ids':['s']},{'sheet':'Data','header_row':1,'key_columns':['代码'],'values':{'代码':'000002','简称':'丙','金额（万元）':2,'公告日期':'2026-02-01'},'mark_entire_row':True,'source_ids':['s']}],a)
  ws=openpyxl.load_workbook(self.p).active;self.assertEqual(ws['B2'].value,'乙');self.assertEqual(ws['A3'].value,'000002');self.assertEqual(ws['D3'].number_format,'yyyy-mm-dd');self.assertEqual(len(a),6)
 def test_append_skips_styled_self_closing_rows(self):
  xml=('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
       '<dimension ref="A1:A62"/><sheetData>'
       '<row r="60"><c r="A60" t="inlineStr"><is><t>x</t></is></c></row>'
       '<row r="61" s="1"/><row r="62" s="1"/></sheetData></worksheet>').encode()
  out,n=ooxml.append_rows(xml,60,[{'A':'new'}])
  text=out.decode()
  self.assertEqual(n,1)
  self.assertIn('<row r="63"',text)
  self.assertEqual(text.count('r="62"'),1)

 def test_green_row_style_does_not_inherit_date_number_format(self):
  styles=(b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          b'<fills count="2"><fill/><fill/></fills>'
          b'<cellXfs count="1"><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/></cellXfs>'
          b'</styleSheet>')
  out,idx=ooxml.add_green_fill(styles)
  self.assertIsNotNone(idx)
  self.assertIn(b'numFmtId="0"',out)

 def test_unknown_header_rejected_without_save(self):
  original=self.p.read_bytes()
  with self.assertRaises(ValueError):mutate_file(self.p,[{'sheet':'Data','header_row':1,'key_columns':['代码'],'values':{'代码':'x','不存在':1}}],[])
  # mutate_file may save only at end; temp copy transaction is tested by caller
  self.assertEqual(self.p.read_bytes(),original)
if __name__=='__main__':unittest.main(verbosity=2)
