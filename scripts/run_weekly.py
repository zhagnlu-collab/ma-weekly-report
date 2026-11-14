#!/usr/bin/env python3
"""Date-driven run state for ma-weekly-report. It never fabricates content."""
from __future__ import annotations
import argparse, hashlib, json, re, shutil, sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

STAGES=['PLAN','ACQUIRE','EXTRACT','CLASSIFY','LEDGER','METRICS','WORKBOOKS','WORD','PPT','PDF','VALIDATE','PUBLISH']
STATUSES={'PENDING','RUNNING','PASS','RETRY','FAILED','BLOCKED','NEEDS_COMPUTER_USE','NEEDS_SESSION','UI_SCHEMA_CHANGED','COMPLETE'}
NATIVE_ROLES={f'xlsx_{i}' for i in range(14) if i!=5}|{'weekly_docx','dynamic_docx','market_ppt'}
CN_TZ=timezone(timedelta(hours=8))

def sha(p:Path)->str:return hashlib.sha256(p.read_bytes()).hexdigest()
def new_name(name,old_start,old_end,start,end):
    n=name.replace(old_start,start.strftime('%Y%m%d')).replace(old_end,end.strftime('%Y%m%d'))
    n=re.sub(r'年第(\d+)期',lambda m:'年第'+str(int(m.group(1))+1)+'期',n)
    if old_start[:4]==old_end[:4]==start.strftime('%Y%m%d')[:4]:
        pass
    elif start.year==end.year:
        n=n.replace(old_start[:4]+'年',str(start.year)+'年')
    return n
def now():return datetime.now(CN_TZ).replace(microsecond=0).isoformat()
def role(p:Path)->str:
 n=p.name
 if n.startswith('并购重组市场周报') and p.suffix.lower()=='.pptx':return 'market_ppt'
 if n.startswith('A股市场并购重组动态') and p.suffix.lower()=='.docx':return 'dynamic_docx'
 m=re.match(r'^(\d+)、(.*)$',n)
 if m and p.suffix.lower()=='.xlsx':
  if m.group(1)=='5':return 'weekly_docx'
  return 'xlsx_'+m.group(1)
 if m and m.group(1)=='5' and p.suffix.lower()=='.docx':return 'weekly_docx'
 raise ValueError('无法识别角色: '+n)
def discover_template(root:Path):
 candidates=[]
 for d in root.iterdir():
  if not d.is_dir() or d.name.startswith('.') or '并购周报' not in d.name:continue
  try:complete_inventory(d)
  except Exception:continue
  m=re.search(r'[-_](\d{8})至(\d{8})',d.name)
  if m:candidates.append((m.group(2),m.group(1),d))
 if not candidates:raise ValueError(f'{root}下没有包含完整16个Office角色的上期目录')
 return sorted(candidates)[-1][2]

def complete_inventory(d:Path):
 result={}
 for p in d.iterdir():
  if not p.is_file() or p.name.startswith('~$'):continue
  try:r=role(p)
  except ValueError:continue
  if r in result:raise ValueError(f'角色重复: {r}')
  result[r]=p
 missing=sorted(NATIVE_ROLES-set(result));extra=sorted(set(result)-NATIVE_ROLES)
 if missing or extra:raise ValueError(f'完整成果模板不合法，missing={missing}, extra={extra}')
 return result
def load_state(p:Path):return json.loads(p.read_text(encoding='utf-8'))
def save_state(p:Path,s):
 p.parent.mkdir(parents=True,exist_ok=True);p.write_text(json.dumps(s,ensure_ascii=False,indent=2),encoding='utf-8')
def short(date_s, compact=True):d=date.fromisoformat(date_s);return d.strftime('%Y%m%d') if compact else d.strftime('%Y年%m月%d日')

def new_run(args):
 asof=date.fromisoformat(args.as_of);start=asof-timedelta(days=6)
 if start.year!=asof.year and not args.allow_cross_year:raise ValueError('默认周区间跨年；请显式 --allow-cross-year 并确认年度轮转')
 inv=complete_inventory(args.template_dir)
 base=inv['xlsx_0'].name
 rm=re.search(r'_(\d{8})至(\d{8})',base)
 if not rm:raise ValueError('当年基础数据文件名缺少起止日期')
 old_start,old_end=rm.groups()
 run_id=f'{start:%Y%m%d}_{asof:%Y%m%d}'
 work=args.work_root/run_id
 if work.exists() and not args.resume:raise ValueError(f'运行目录已存在: {work}')
 state_path=work/'run-state.json'
 if state_path.exists():
  state=load_state(state_path)
  if state['config_sha']!=args.config_sha and not args.allow_config_change:raise ValueError('配置已变化，需明确 --allow-config-change')
  return state,state_path
 if work.exists():
  # resume without state is unsafe; never infer completed stages from files.
  raise ValueError('目录存在但无run-state.json，拒绝自动续跑')
 work.mkdir(parents=True)
 stage_dir=work/'0_working'
 stage_dir.mkdir()
 copied=[]
 for r,p in sorted(inv.items()):
  dest=stage_dir/new_name(p.name,old_start,old_end,start,asof);shutil.copy2(p,dest);copied.append({'role':r,'source':str(p.resolve()),'staged':str(dest),'sha256':sha(dest)})
 stages={x:{'status':'PENDING','attempt':0,'history':[]} for x in STAGES}
 state={'run_id':run_id,'created_at':now(),'updated_at':now(),'as_of':asof.isoformat(),'period_start':start.isoformat(),'period_end':asof.isoformat(),'timezone':'Asia/Shanghai','profile':args.profile,'template_dir':str(args.template_dir.resolve()),'work_dir':str(work.resolve()),'state_path':str(state_path.resolve()),'config_sha':args.config_sha,'stages':stages,'template_files':copied,'outputs':{},'notes':[]}
 stages['PLAN']['status']='PASS';stages['PLAN']['history'].append({'at':now(),'status':'PASS','detail':'period and complete template validated'})
 save_state(state_path,state);return state,state_path

def set_stage(args):
 state=load_state(args.state);name=args.stage.upper();status=args.status.upper()
 if name not in STAGES:raise ValueError('未知阶段')
 if status not in STATUSES:raise ValueError('未知状态')
 if name!='PLAN' and status in {'PASS','COMPLETE'}:
  idx=STAGES.index(name);prev=STAGES[idx-1]
  if state['stages'][prev]['status']!='PASS':raise ValueError(f'{prev}尚未PASS，不能将{name}标记为{status}')
 if name!='PUBLISH' and status=='COMPLETE':raise ValueError('只有PUBLISH阶段可为COMPLETE')
 st=state['stages'][name];st['status']=status;st['attempt']=st.get('attempt',0)+(1 if status in {'RUNNING','RETRY'} else 0);st.setdefault('history',[]).append({'at':now(),'status':status,'detail':args.detail or '','artifact':args.artifact})
 if args.artifact:
  p=Path(args.artifact)
  if not p.is_absolute():p=Path(state['work_dir'])/p
  if not p.exists():raise ValueError(f'指定产物不存在: {p}')
  state['outputs'].setdefault(name,[]).append({'path':str(p.resolve()),'sha256':sha(p) if p.is_file() else None})
 state['updated_at']=now();save_state(args.state,state);print(json.dumps({'run_id':state['run_id'],'stage':name,'status':status,'next':next_stage(state)},ensure_ascii=False));return 0

def next_stage(state):
 for x in STAGES:
  if state['stages'][x]['status'] not in {'PASS','COMPLETE'}:return x
 return 'DONE'
def main():
 ap=argparse.ArgumentParser(description=__doc__);sub=ap.add_subparsers(dest='cmd',required=True)
 ini=sub.add_parser('init');ini.add_argument('--as-of',required=True);ini.add_argument('--template-dir',type=Path,help='省略时在当前目录自动选择最近完整上期');ini.add_argument('--work-root',type=Path,help='省略时使用当前目录/ma-weekly-runs');ini.add_argument('--profile',default='production');ini.add_argument('--config-sha',default='default');ini.add_argument('--resume',action='store_true');ini.add_argument('--allow-config-change',action='store_true');ini.add_argument('--allow-cross-year',action='store_true')
 st=sub.add_parser('set');st.add_argument('state',type=Path);st.add_argument('stage');st.add_argument('status');st.add_argument('--detail');st.add_argument('--artifact')
 sh=sub.add_parser('status');sh.add_argument('state',type=Path)
 args=ap.parse_args()
 try:
  if args.cmd=='init':
   if args.template_dir is None:args.template_dir=discover_template(Path.cwd())
   if args.work_root is None:args.work_root=Path.cwd()/'ma-weekly-runs'
   state,path=new_run(args);print(json.dumps({'state':str(path),'run_id':state['run_id'],'period_start':state['period_start'],'period_end':state['period_end'],'next':next_stage(state)},ensure_ascii=False));return 0
  if args.cmd=='set':return set_stage(args)
  state=load_state(args.state);print(json.dumps({'run_id':state['run_id'],'next':next_stage(state),'stages':state['stages']},ensure_ascii=False,indent=2));return 0
 except Exception as e:print(json.dumps({'status':'FAIL','error':str(e)},ensure_ascii=False));return 1
if __name__=='__main__':sys.exit(main())
