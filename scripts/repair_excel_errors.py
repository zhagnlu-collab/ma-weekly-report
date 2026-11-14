#!/usr/bin/env python3
"""Repair explicitly selected Excel formula errors and replace verified sentinel errors."""
from __future__ import annotations
import argparse, copy, hashlib, json, re, shutil, sys, tempfile
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile
import openpyxl

ROLE_RE=re.compile(r'^(\d+)、(.*)\.xlsx$',re.I)
def files(d):
 out={}
 for p in d.glob('*.xlsx'):
  m=ROLE_RE.match(p.name)
  if m and not p.name.startswith('~$'):out['xlsx_'+m.group(1)]=p
 return out
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()

def main():
 ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('workdir',type=Path);ap.add_argument('--plan',required=True,type=Path);ap.add_argument('--out-audit',required=True,type=Path);args=ap.parse_args()
 plan=json.loads(args.plan.read_text(encoding='utf-8'));lookup=files(args.workdir)
 wraps=plan.get('formula_wraps',[]);repls=plan.get('replacements',[])
 if not wraps and not repls:raise ValueError('repair plan is empty')
 groups={}
 for kind,items in [('wrap',wraps),('replace',repls)]:
  for m in items:
   role=m.get('file_role')
   if role not in lookup:raise ValueError(f'unknown role {role}')
   groups.setdefault(role,[]).append((kind,m))
 stamp=datetime.now().strftime('%Y%m%d_%H%M%S');backup=args.workdir/'.excel-repair-backup'/stamp;backup.mkdir(parents=True,exist_ok=True)
 audit=[]
 with tempfile.TemporaryDirectory(prefix='ma-weekly-repair-') as td:
  staging=Path(td)
  for role,items in groups.items():
   src=lookup[role];dst=staging/src.name;shutil.copy2(src,dst)
   w=openpyxl.load_workbook(dst,data_only=False);v=openpyxl.load_workbook(dst,data_only=True)
   for kind,m in items:
    sheet=m['sheet'];s=w[sheet]
    cells=m.get('cells') or [m['cell']]
    for addr in cells:
     c=s[addr]
     oldraw=c.value;oldcached=v[sheet][addr].value
     if kind=='wrap':
      if not isinstance(oldraw,str) or not oldraw.startswith('='):raise ValueError(f'{src.name}/{sheet}/{addr}: wrap目标不是公式')
      if oldcached and not isinstance(oldcached,str):continue
      fallback=m.get('fallback','-');new='=IFERROR('+oldraw[1:]+','+json.dumps(fallback,ensure_ascii=False)+')'
     else:
      new=m['value']
     c.value=new;audit.append({'file':src.name,'sheet':sheet,'cell':addr,'old_raw':str(oldraw),'old_cached':str(oldcached),'new':str(new),'kind':kind,'source_ids':m.get('source_ids',[]),'reason':m.get('reason','')})
   w.save(dst)
   with ZipFile(dst) as z:
    if z.testzip():raise ValueError(f'{src.name}: damaged zip')
   openpyxl.load_workbook(dst,read_only=True).close()
   shutil.copy2(src,backup/src.name);shutil.copy2(dst,src)
 result={'status':'PASS','files_changed':len(groups),'changes':len(audit),'audit':audit,'backup':str(backup.resolve()),'note':'公式缓存需随后用Microsoft Excel重新计算并保存'}
 args.out_audit.parent.mkdir(parents=True,exist_ok=True);args.out_audit.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps({'status':'PASS','files_changed':len(groups),'changes':len(audit),'out':str(args.out_audit)},ensure_ascii=False));return 0
if __name__=='__main__':
 try:sys.exit(main())
 except Exception as e:print(json.dumps({'status':'FAIL','error':str(e)},ensure_ascii=False));sys.exit(1)
