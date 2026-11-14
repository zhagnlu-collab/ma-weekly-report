#!/usr/bin/env python3
"""Structural and consistency gate for a complete weekly package. Never recalculates Office formulas."""
from __future__ import annotations
import argparse, hashlib, json, re, sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import openpyxl
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).resolve().parent))
from materialize_package import resolve_office_pdfs  # noqa: E402

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
REQUIRED={f'xlsx_{i}' for i in range(14) if i!=5}|{'weekly_docx','dynamic_docx','market_ppt'}

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def role(p):
 n=p.name
 if n.startswith('并购重组市场周报') and p.suffix.lower()=='.pptx':return 'market_ppt'
 if n.startswith('A股市场并购重组动态') and p.suffix.lower()=='.docx':return 'dynamic_docx'
 m=re.match(r'^(\d+)、',n)
 if m and p.suffix.lower()=='.xlsx' and m.group(1)!='5':return 'xlsx_'+m.group(1)
 if m and m.group(1)=='5' and p.suffix.lower()=='.docx':return 'weekly_docx'
 raise ValueError('无法识别角色 '+n)
def office_text(p:Path):
 text=[]
 with ZipFile(p) as z:
  if p.suffix.lower()=='.docx':
   root=ET.fromstring(z.read('word/document.xml'))
   text=[''.join(x.itertext()) for x in root.iter(W+'t')]
  elif p.suffix.lower()=='.pptx':
   slides=sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)],key=lambda n:int(re.search(r'\d+',Path(n).stem).group()))
   for n in slides:
    root=ET.fromstring(z.read(n));text.extend(x.text or '' for x in root.iter(A+'t'))
 return ''.join(text)
def main():
 ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('directory',type=Path);ap.add_argument('--period',required=True);ap.add_argument('--expected',type=Path);ap.add_argument('--allow-errors',type=Path,help='JSON exact allowlist: [{file,sheet,cell,reason}]');ap.add_argument('--out',required=True,type=Path);args=ap.parse_args()
 errors=[];warnings=[];files={};expected=json.loads(args.expected.read_text()) if args.expected else {};allowed=json.loads(args.allow_errors.read_text()) if args.allow_errors else []
 allow_keys={(x.get('file'),x.get('sheet'),x.get('cell')) for x in allowed}
 for x in allowed:
  if not x.get('reason'):errors.append('Excel error allowlist item lacks reason: '+str(x))
 for p in args.directory.iterdir():
  if not p.is_file() or p.name.startswith('~$'):continue
  try:r=role(p)
  except ValueError:continue
  if r in files:errors.append(f'角色重复: {r}')
  files[r]=p
 missing=sorted(REQUIRED-set(files));extra=sorted(set(files)-REQUIRED)
 if missing:errors.append('缺少角色: '+','.join(missing))
 if extra:errors.append('多余角色: '+','.join(extra))
 if not missing and not extra:
  try:
   resolve_office_pdfs(files,args.directory)
  except Exception as e:
   errors.append(f'Microsoft Office发布PDF门禁失败: {e}')
 range_expect=expected.get('filename_date_contains',args.period)
 details=[]
 for r,p in sorted(files.items()):
  try:
   with ZipFile(p) as z:
    bad=z.testzip()
    if bad:errors.append(f'{p.name}: ZIP损坏 {bad}')
   if range_expect and range_expect not in p.name:errors.append(f'{p.name}: 文件名缺少期间 {range_expect}')
   rec={'role':r,'file':p.name,'sha256':sha(p)}
   if r.startswith('xlsx_'):
    wb=openpyxl.load_workbook(p,data_only=False,read_only=False);cached=openpyxl.load_workbook(p,data_only=True,read_only=False)
    formulas=0;errs=[];missing_cache=[]
    for ws in wb:
     cs=ws._cells.values() if hasattr(ws,'_cells') else []
     for c in cs:
      if c.data_type=='f':
       formulas+=1
       if cached[ws.title][c.coordinate].value is None:missing_cache.append(c.coordinate)
    for ws in cached:
     cs=ws._cells.values() if hasattr(ws,'_cells') else []
     for c in cs:
      if c.data_type=='e':errs.append({'sheet':ws.title,'cell':c.coordinate,'value':c.value})
    rec.update(kind='excel',sheets=len(wb.sheetnames),formulas=formulas,missing_formula_cache=missing_cache,cached_errors=errs)
    if missing_cache:errors.append(f'{p.name}: {len(missing_cache)}个公式无缓存值，需用Excel重新计算并保存')
    unallowed=[e for e in errs if (p.name,e['sheet'],e['cell']) not in allow_keys]
    if unallowed:
     errors.extend(f'{p.name}/{e["sheet"]}/{e["cell"]}: 未解释的Excel错误 {e["value"]}' for e in unallowed)
    if errs and len(unallowed)<len(errs):warnings.append(f'{p.name}: {len(errs)-len(unallowed)}个Excel错误已按精确清单元数据豁免')
   else:
    txt=office_text(p);rec.update(kind=p.suffix.lower()[1:],text_chars=len(txt))
    if p.suffix.lower()=='.pptx':
     with ZipFile(p) as z:rec['slides']=len([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)])
   details.append(rec)
  except Exception as e:errors.append(f'{p.name}: {e}')
 # PDFs
 pdfs=list(args.directory.glob('*.pdf'));pdf_info=[]
 for p in pdfs:
  try:
   pages=len(PdfReader(p).pages);text=''.join((x.extract_text() or '') for x in PdfReader(p).pages);pdf_info.append({'file':p.name,'pages':pages,'text_chars':len(text)})
   if pages==0:errors.append(f'{p.name}: 0页')
  except Exception as e:errors.append(f'{p.name}: {e}')
 if len(pdfs)!=3:errors.append(f'PDF数量={len(pdfs)}，完整成果包必须为3份')
 # optional expected
 for r,checks in expected.get('roles',{}).items():
  p=files.get(r)
  if not p:continue
  txt=office_text(p) if p.suffix in ('.docx','.pptx') else ''
  for s in checks.get('must_contain',[]):
   if s not in txt:errors.append(f'{p.name}: 缺少必备文本 {s}')
  for s in checks.get('must_not_contain',[]):
   if s in txt:errors.append(f'{p.name}: 含禁止残留 {s}')
 # PPT PDF page equality where filename stem matches
 expected_pdf_pages=expected.get('pdf_pages',{})
 for pdf in pdf_info:
  stem=Path(pdf['file']).stem
  if stem in expected_pdf_pages and pdf['pages']!=int(expected_pdf_pages[stem]):errors.append(f'{pdf["file"]}: PDF页数{pdf["pages"]} != 预期{expected_pdf_pages[stem]}')
  native=args.directory/(stem+'.pptx')
  if native.exists():
   slides=next((x.get('slides') for x in details if x['file']==native.name),None)
   if slides is not None and slides!=pdf['pages']:errors.append(f'{pdf["file"]}: PDF页数{pdf["pages"]} != 幻灯片数{slides}')
 result={'status':'PASS' if not errors else 'FAIL','period':args.period,'required_native':len(REQUIRED),'native_found':len(files),'pdf_found':len(pdfs),'errors':errors,'warnings':warnings,'files':details,'pdfs':pdf_info}
 args.out.parent.mkdir(parents=True,exist_ok=True);args.out.write_text(json.dumps(result,ensure_ascii=False,indent=2));print(json.dumps({'status':result['status'],'native':len(files),'pdfs':len(pdfs),'errors':len(errors),'warnings':len(warnings),'out':str(args.out)},ensure_ascii=False));return 0 if not errors else 1
if __name__=='__main__':sys.exit(main())
