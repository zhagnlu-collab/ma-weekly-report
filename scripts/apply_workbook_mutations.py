#!/usr/bin/env python3
"""Apply evidence-backed, header-mapped mutations to a working copy of Excel templates."""
from __future__ import annotations
import argparse, copy, hashlib, json, re, shutil, sys, tempfile
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile
import openpyxl
from openpyxl.utils import get_column_letter

GREEN='FFE2EFDA'
ROLE_RE=re.compile(r'^(\\d+)、(.*)\\.xlsx$',re.I)

def sha(p:Path):return hashlib.sha256(p.read_bytes()).hexdigest()
def clean_header(v):return re.sub(r'\\s+','',str(v or '')).strip()
def files_by_role(d:Path):
 out={}
 for p in d.glob('*.xlsx'):
  if p.name.startswith('~$'):continue
  m=ROLE_RE.match(p.name)
  if m:out['xlsx_'+m.group(1)]=p
 return out

def detect_header(ws,explicit=None):
 if explicit:return int(explicit)
 best=(0,-1,0)
 for r in range(1,min(ws.max_row,15)+1):
  vals=[clean_header(ws.cell(r,c).value) for c in range(1,min(ws.max_column,80)+1)]
  named=sum(bool(x) for x in vals)
  unique=len(set(x for x in vals if x))
  score=(named,unique,len(''.join(vals)))
  if score>best:best=score
 if best[0]<2:raise ValueError(f'{ws.title}: 无法可靠识别表头')
 return best and next(r for r in range(1,min(ws.max_row,15)+1) if sum(bool(clean_header(ws.cell(r,c).value)) for c in range(1,min(ws.max_column,80)+1))==best[0])
def header_map(ws,row):
 result={}
 for c in range(1,ws.max_column+1):
  h=clean_header(ws.cell(row,c).value)
  if h:
   if h in result:raise ValueError(f'{ws.title}: 重复表头 {h}')
   result[h]=c
 return result
def normalize(v):
 if v is None:return ''
 if isinstance(v,(datetime,int,float)):return str(v)
 return re.sub(r'\\s+','',str(v)).upper()
def find_rows(ws,hm,keyvals,start):
 rows=[]
 for r in range(start,ws.max_row+1):
  ok=True
  for k,v in keyvals.items():
   if k not in hm:raise ValueError(f'{ws.title}: 无键字段 {k}')
   if normalize(ws.cell(r,hm[k]).value)!=normalize(v):ok=False;break
  if ok:rows.append(r)
 return rows
def mutate_file(path:Path,muts,audit):
 wb=openpyxl.load_workbook(path,data_only=False,keep_vba=path.suffix.lower()=='.xlsm')
 for m in muts:
  ws=wb[m['sheet']];hr=detect_header(ws,m.get('header_row'));hm=header_map(ws,hr);start=hr+1
  vals=m['values']
  for k in vals:
   if k not in hm:raise ValueError(f'{path.name}/{ws.title}: 未找到列 {k}')
  keycols=m.get('key_columns',[])
  if not keycols:raise ValueError('mutation requires key_columns')
  keyvals={k:vals[k] for k in keycols if k in vals}
  if len(keyvals)!=len(keycols):raise ValueError('key_columns 必须存在于 values')
  matches=find_rows(ws,hm,keyvals,start)
  if len(matches)>1:raise ValueError(f'{path.name}/{ws.title}: 键匹配多行 {matches}')
  if matches:row=matches[0];mode='update'
  elif m.get('mode','upsert')=='update':raise ValueError(f'{path.name}/{ws.title}: 键不存在')
  else:
   # conservative append: caller may supply insert_at; otherwise append after current used range.
   row=int(m.get('insert_at') or ws.max_row+1);mode='append'
   ref=int(m.get('style_from_row') or max(hr+1,row-1))
   if row<=ref and row<=ws.max_row:
    ws.insert_rows(row)
   for c in range(1,ws.max_column+1):
    src=ws.cell(ref,c);dst=ws.cell(row,c)
    if src.has_style:dst._style=copy.copy(src._style)
    if src.number_format:dst.number_format=src.number_format
  for k,v in vals.items():
   c=hm[k];cell=ws.cell(row,c)
   if cell.data_type=='f' and not m.get('allow_formula_override',False):raise ValueError(f'{path.name}/{ws.title}/{cell.coordinate}: 拒绝覆盖公式')
   if isinstance(v,str) and (('日期' in k or '时间' in k) and re.fullmatch(r'\\d{4}-\\d{2}-\\d{2}',v)):
    v=datetime.strptime(v,'%Y-%m-%d')
   old=cell.value;audit.append({'file':path.name,'sheet':ws.title,'cell':cell.coordinate,'field':k,'old':str(old) if old is not None else None,'new':str(v) if v is not None else None,'mode':mode,'source_ids':m.get('source_ids',[]),'reason':m.get('reason','')})
   cell.value=v
   if '日期' in k or '时间' in k:
    cell.number_format=m.get('number_format','yyyy-mm-dd')
  marked=m.get('mark_columns')
  if marked:
   for k in marked:
    if k not in hm:raise ValueError(f'{path.name}/{ws.title}: mark列不存在 {k}')
    ws.cell(row,hm[k]).fill=openpyxl.styles.PatternFill('solid',fgColor=GREEN)
  elif m.get('mark_entire_row'):
   for c in range(1,min(ws.max_column,80)+1):ws.cell(row,c).fill=openpyxl.styles.PatternFill('solid',fgColor=GREEN)
 wb.save(path)

def main():
 ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('workdir',type=Path);ap.add_argument('--mutations',required=True,type=Path);ap.add_argument('--out-audit',required=True,type=Path);args=ap.parse_args()
 data=json.loads(args.mutations.read_text(encoding='utf-8'));muts=data.get('mutations',[])
 if not muts:raise ValueError('mutations不能为空')
 byrole=files_by_role(args.workdir);groups={}
 for m in muts:
  r=m.get('file_role')
  if r not in byrole:raise ValueError(f'找不到角色 {r}')
  groups.setdefault(r,[]).append(m)
 stamp=datetime.now().strftime('%Y%m%d_%H%M%S');backup=args.workdir/'.mutations-backup'/stamp;backup.mkdir(parents=True,exist_ok=True)
 audit=[];before={}
 with tempfile.TemporaryDirectory(prefix='ma-weekly-mut-') as td:
  staging=Path(td)
  for r,items in groups.items():
   src=byrole[r];dst=staging/src.name;shutil.copy2(src,dst);before[str(src.resolve())]=sha(src);mutate_file(dst,items,audit)
   with ZipFile(dst) as z:
    bad=z.testzip()
    if bad:raise ValueError(f'{src.name}: zip损坏 {bad}')
   openpyxl.load_workbook(dst,read_only=True).close()
  for r in groups:
   src=byrole[r];shutil.copy2(src,backup/src.name);shutil.copy2(staging/src.name,src)
 result={'status':'PASS','mutations':len(muts),'files_changed':len(groups),'audit':audit,'backup':str(backup.resolve())}
 args.out_audit.parent.mkdir(parents=True,exist_ok=True);args.out_audit.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps({'status':'PASS','files_changed':len(groups),'cells_changed':len(audit),'audit':str(args.out_audit)},ensure_ascii=False));return 0
if __name__=='__main__':
 try:sys.exit(main())
 except Exception as e:print(json.dumps({'status':'FAIL','error':str(e)},ensure_ascii=False));sys.exit(1)
