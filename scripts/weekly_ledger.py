#!/usr/bin/env python3
"""校验已人工核验的事件台账，生成不同时间口径的唯一交易数与顾问汇总。标准库。"""
import argparse
from collections import Counter
from datetime import date
from decimal import Decimal
import json
from pathlib import Path

BOARDS = {'沪主板','深主板','科创板','创业板','北交所'}
KINDS = {'ma_audit','ma_cash_purchase','ma_cash_sale','control','partial_transfer','second_shareholder','ipo','strategic_investment','professional_cooperation'}
TYPES = {'first_disclosed','first_draft','accepted','inquiry','replied','suspended','resumed','committee_scheduled','committee_passed','committee_deferred','committee_rejected','simplified_passed','registration_submitted','registered','terminated','control_completed','control_progress','completed','progress','lead_signal'}
REVIEW = {'accepted','inquiry','replied','suspended','resumed','committee_scheduled','committee_deferred'}
STATES = REVIEW | {'committee_passed','committee_rejected','simplified_passed','registration_submitted','registered','terminated'}
STATUS = {'CONFIRMED','CANDIDATE','UNVERIFIABLE'}

def day(s):
    return date.fromisoformat(s)

def decimal(v):
    return Decimal(str(v))

def validate(data):
    errors = []
    def check(ok, msg):
        if not ok: errors.append(msg)
    for k in ['period','sources','cases','events','coverage']:
        if k not in data: errors.append('缺少 '+k)
    if errors: return errors
    try:
        period = data['period']; start=day(period['start']);end=day(period['end']);cutoff=day(period['cutoff'])
        check(start<=cutoff<=end, '需要 start <= cutoff <= end')
    except (ValueError,KeyError,TypeError): return ['period日期需ISO格式，含start/end/cutoff']
    sources={};cases={};events={}
    for s in data['sources']:
        sid=s.get('source_id');check(bool(sid) and sid not in sources,'来源键为空或重复: '+str(sid));sources[sid]=s
        check(bool(s.get('locator')) and bool(s.get('excerpt')) and bool(s.get('file_or_url')), '来源缺URL/文件、定位或摘录: '+str(sid))
        try: day(s['published_on'])
        except (ValueError,KeyError,TypeError): errors.append('来源日期无效: '+str(sid))
    for c in data['cases']:
        cid=c.get('deal_id');check(bool(cid) and cid not in cases,'交易键为空或重复: '+str(cid));cases[cid]=c
        check(c.get('board') in BOARDS,'板块无效: '+str(cid))
        check(c.get('kind') in KINDS,'kind无效: '+str(cid))
        allowed_overlap={'partial_transfer':{'second_shareholder'},'second_shareholder':{'partial_transfer'},'ma_cash_purchase':{'ma_cash_sale'},'ma_cash_sale':{'ma_cash_purchase'}}
        check(set(c.get('also_kinds',[]))<=allowed_overlap.get(c.get('kind'),set()),'also_kinds组合无效: '+str(cid))
        check(bool(c.get('code')) and bool(c.get('name')), '交易缺代码/简称: '+str(cid))
        for a in c.get('advisors',[]):
            check(a.get('role') in {'independent','financial','acquirer'},'顾问角色无效: '+str(cid))
            check(bool(a.get('name')) and not any(x in a.get('name','') for x in ['未披露','待核','尚未','待确定']),'顾问名称为空或未公开线索: '+str(cid))
            try: day(a['disclosed_on'])
            except (ValueError,KeyError,TypeError):errors.append('顾问公开日期无效: '+str(cid))
            check(bool(a.get('source_ids')) and all(x in sources for x in a.get('source_ids',[])), '顾问缺来源: '+str(cid))
        for value in c.get('amounts',[]):
            check(value.get('basis') in {'initial_plan','latest_plan','actual_completion'},'金额阶段无效: '+str(cid))
            check(bool(value.get('currency')),'金额缺币种: '+str(cid))
            try:
                check(decimal(value['value_wan']).is_finite(), '金额不是有限数字: '+str(cid));day(value['known_on'])
            except (KeyError,ValueError,TypeError,ArithmeticError):errors.append('金额或已知日无效: '+str(cid))
            check(bool(value.get('source_ids')) and all(x in sources for x in value.get('source_ids',[])), '金额缺来源: '+str(cid))
    signatures=set(); first_milestones=set(); state_order=set()
    for ev in data['events']:
        eid=ev.get('event_id');check(bool(eid) and eid not in events, '事件键为空或重复: '+str(eid));events[eid]=ev
        check(ev.get('deal_id') in cases,'事件对应交易不存在: '+str(eid))
        check(ev.get('type') in TYPES,'事件类型无效: '+str(eid));check(ev.get('status') in STATUS,'事件状态无效: '+str(eid))
        try:
            ed=day(ev['date']);pub=day(ev['announced_on'])
            if ev.get('type') in {'first_disclosed','first_draft','registered'}:
                check(ed==pub,'该类型统计日采用公告日；实际事件日另存: '+str(eid))
            check(isinstance(ev.get('sequence',0),int) and ev.get('sequence',0)>=0,'sequence需非负整数: '+str(eid))
            if ev.get('type') in STATES:
                state_sig=(ev.get('deal_id'),ev.get('date'),ev.get('sequence',0))
                check(state_sig not in state_order,'同日状态先后不明，需sequence: '+str(eid));state_order.add(state_sig)
            sig=(ev.get('deal_id'),ev.get('type'),ev.get('date'),ev.get('sequence',0))
            check(sig not in signatures, '同事件重复，请合并来源: '+str(eid));signatures.add(sig)
            if ev.get('type') in {'first_disclosed','first_draft'}:
                milestone=(ev.get('deal_id'),ev.get('type'))
                check(milestone not in first_milestones,'首次里程碑重复: '+str(eid));first_milestones.add(milestone)
        except (KeyError,ValueError,TypeError):errors.append('事件日期/sequence无效: '+str(eid))
        check(bool(ev.get('source_ids')) and all(x in sources for x in ev.get('source_ids',[])), '事件缺来源: '+str(eid))
        if ev.get('type','').startswith('committee_') and ev.get('type')!='committee_scheduled':
            check(bool(ev.get('meeting_id')),'委员会结果缺meeting_id: '+str(eid))
    for ev in data['events']:
        refs=[sources[x] for x in ev.get('source_ids',[]) if x in sources]
        try:
            if refs: check(min(day(x['published_on']) for x in refs)<=day(ev['announced_on']), '事件声称已知日早于全部来源: '+str(ev.get('event_id')))
        except (KeyError,ValueError,TypeError): pass
    for c in data['cases']:
        for fact in c.get('advisors',[])+c.get('amounts',[]):
            refs=[sources[x] for x in fact.get('source_ids',[]) if x in sources]
            try:
                known=day(fact.get('disclosed_on') or fact.get('known_on'))
                if refs: check(min(day(x['published_on']) for x in refs)<=known,'字段声称已知日早于全部来源: '+str(c['deal_id']))
            except (KeyError,ValueError,TypeError): pass
    # coverage is a declaration, not validated evidence of actually searching.
    for q in data['coverage']:
        check(q.get('status') in {'COMPLETE','PARTIAL','LIMITED','NOT_RUN','NOT_APPLICABLE'},'覆盖状态无效')
        check(bool(q.get('group')) and bool(q.get('evidence')), '覆盖缺检索组或记录定位')
    check(bool(data['coverage']),'覆盖记录不能为空；草稿也应记录NOT_RUN')
    return errors

def compute(data):
    errors=validate(data)
    if errors: raise ValueError('\n'.join(errors))
    period=data['period'];start=day(period['start']);cutoff=day(period['cutoff']);yearstart=date(cutoff.year,1,1)
    cases={c['deal_id']:c for c in data['cases']}
    events=[e for e in data['events'] if e['status']=='CONFIRMED' and day(e['announced_on'])<=cutoff and day(e['date'])<=cutoff]
    ma={k:c for k,c in cases.items() if c['kind']=='ma_audit'}
    bycase={cid:[] for cid in cases}
    for e in events: bycase[e['deal_id']].append(e)
    def cohort(types,lo,hi,kinds=None):
        return {e['deal_id'] for e in events if e['type'] in types and lo<=day(e['date'])<=hi and (kinds is None or bool(({cases[e['deal_id']]['kind']} | set(cases[e['deal_id']].get('also_kinds',[]))) & kinds))}
    def advisors(cid):
        # 顾问是截止日已公开的当前独立财顾；不含“未披露”线索及一般财顾。
        return {a['name'] for a in cases[cid].get('advisors',[]) if a['role']=='independent' and day(a['disclosed_on'])<=cutoff}
    def amount(cid, actual=False):
        # 中间汇总采用当前最新计划，后续实际对价独立保留；不得自动用实际替代计划。
        values=[v for v in cases[cid].get('amounts',[]) if v['basis'] in ({'actual_completion'} if actual else {'initial_plan','latest_plan'}) and v['currency']=='CNY' and day(v['known_on'])<=cutoff]
        if not values: return None
        value=sorted(enumerate(values),key=lambda p:(day(p[1]['known_on']),p[0]))[-1][1]
        return decimal(value['value_wan'])
    def summary(ids, actual=False):
        ids=sorted(ids);rows={};total=Decimal(0);unknown=[];boards=Counter()
        for cid in ids:
            boards[cases[cid]['board']]+=1;value=amount(cid, actual)
            if value is None:unknown.append(cid)
            else:total+=value
            for name in advisors(cid):
                row=rows.setdefault(name,{'projects':0,'known_amount_wan':Decimal(0),'unknown_amount_projects':0})
                row['projects']+=1
                if value is None:row['unknown_amount_projects']+=1
                else:row['known_amount_wan']+=value
        for row in rows.values():row['known_amount_wan']=str(row['known_amount_wan'])
        return {'projects':len(ids),'deal_ids':ids,'by_board':dict(boards),'projects_with_disclosed_independent_advisor':sum(bool(advisors(cid)) for cid in ids),'advisor_instances':sum(r['projects'] for r in rows.values()),'advisors':dict(sorted(rows.items(),key=lambda p:(-p[1]['projects'],p[0]))),'known_cny_amount_wan':str(total),'unknown_cny_amount_deals':unknown,'amount_basis':'实际完成；不回退计划上限' if actual else '截止日最新计划/原始计划；不是实际完成总价'}
    result={'period':period,'limitations':['事件及覆盖状态来自输入声明，脚本未联网核实','金额汇总仅CNY计划，已知小计不代表全市场完整金额','当前顾问按截止日公开的独立角色；如表注需其他角色须单列复算'],'weekly':{},'annual':{}}
    for group,lo in [('weekly',start),('annual',yearstart)]:
        for name,types in {'first_disclosed':{'first_disclosed'},'first_draft':{'first_draft'},'accepted':{'accepted'},'committee_passed':{'committee_passed'},'simplified_passed':{'simplified_passed'},'approved_including_simplified':{'committee_passed','simplified_passed'},'registered':{'registered'},'terminated':{'terminated'}}.items():
            result[group][name]=summary(cohort(types,lo,cutoff,{'ma_audit'}))
        for kind in KINDS-{'ma_audit'}:
            result[group][kind+'_first_disclosed']=summary(cohort({'first_disclosed'},lo,cutoff,{kind}))
        result[group]['control_completed']=summary(cohort({'control_completed'},lo,cutoff,{'control'}),actual=True)
        ce=[e for e in events if e['type'] in {'committee_passed','committee_deferred','committee_rejected'} and lo<=day(e['date'])<=cutoff and e['deal_id'] in ma]
        result[group]['committee_meetings']=len({e['meeting_id'] for e in ce})
        result[group]['committee_appearances']=len(ce)
        result[group]['committee_unique_reviewed']=len({e['deal_id'] for e in ce})
    inreview=set()
    for cid in ma:
        relevant=sorted([e for e in bycase[cid] if e['type'] in STATES],key=lambda e:(day(e['date']),e.get('sequence',0)))
        if relevant and relevant[-1]['type'] in REVIEW:inreview.add(cid)
    result['as_of_in_review']=summary(inreview)
    result['as_of_in_review_accepted_this_year']=summary(inreview & cohort({'accepted'},yearstart,cutoff,{'ma_audit'}))
    # The report calls plan/draft events a single "initial project" cohort. A
    # combined cash purchase/sale keeps two worksheet-facing cases but must
    # count once in the market total, so dedupe this cohort by security code.
    initial_raw = cohort({'first_disclosed','first_draft'}, start, cutoff,
                         {'ma_audit','ma_cash_purchase','ma_cash_sale'})
    initial_by_code = {}
    for cid in initial_raw:
        initial_by_code[cases[cid].get('code') or cid] = cid
    result['weekly']['initial_projects'] = summary(set(initial_by_code.values()))
    for event_type, field in (('first_disclosed', 'first_disclosed'), ('first_draft', 'first_draft')):
        ids = cohort({event_type}, start, cutoff, {'ma_audit','ma_cash_purchase','ma_cash_sale'})
        by_code = {}
        for cid in ids:
            by_code[cases[cid].get('code') or cid] = cid
        result['weekly'][field] = summary(set(by_code.values()))
    term_week = cohort({'terminated'}, start, cutoff, {'ma_audit'})
    inreview_terms = {cid for cid in term_week if '在审' in str(cases[cid].get('facts', {}).get('termination_stage', ''))}
    plan_terms = {cid for cid in term_week if '预案' in str(cases[cid].get('facts', {}).get('termination_stage', ''))}
    result['weekly']['terminated_in_review'] = summary(inreview_terms)
    result['weekly']['terminated_plan_stage'] = summary(plan_terms)
    # 不把查询覆盖complete当事实全部确认；同时报告候选和因未来时点排除的事件。
    result['unconfirmed_events']=[e['event_id'] for e in data['events'] if e['status']!='CONFIRMED']
    result['coverage']=data['coverage']
    result['coverage_declaration_complete']=all(q['status'] in {'COMPLETE','NOT_APPLICABLE'} for q in data['coverage'])
    return result

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('ledger',type=Path);ap.add_argument('--out',required=True,type=Path);args=ap.parse_args()
    if args.ledger.resolve()==args.out.resolve():ap.error('不得覆盖输入台账')
    try: result=compute(json.loads(args.ledger.read_text(encoding='utf-8')))
    except (ValueError,TypeError,KeyError) as ex:
        print(json.dumps({'validation':'FAIL','error':str(ex)},ensure_ascii=False));raise SystemExit(1)
    args.out.parent.mkdir(parents=True,exist_ok=True);args.out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'validation':'PASS','unconfirmed_events':len(result['unconfirmed_events']),'coverage_declaration_complete':result['coverage_declaration_complete'],'out':str(args.out)},ensure_ascii=False))

if __name__=='__main__':main()
