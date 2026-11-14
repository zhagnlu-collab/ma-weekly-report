#!/usr/bin/env python3
"""Apply explicit, traceable review decisions to an extracted event ledger.

The extractor is deliberately conservative. This step lets an analyst record
include/exclude/promote/facts decisions after reading the PDF, without editing
the generated ledger by hand. Every decision should carry a reason and source.
"""
from __future__ import annotations
import argparse,json,sys
from pathlib import Path


def load(p:Path):
    return json.loads(p.read_text(encoding='utf-8'))


def dump(p:Path,d):
    p.parent.mkdir(parents=True,exist_ok=True)
    p.write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf-8')


def normalize(ledger:dict, decisions:dict)->dict:
    d=json.loads(json.dumps(ledger,ensure_ascii=False))
    cases={c['deal_id']:c for c in d.get('cases',[])}
    events=d.setdefault('events',[])
    removed_case_ids=set()
    for item in decisions.get('sources',[]):
        sid=item.get('source_id')
        if not sid:
            raise ValueError('source decision requires source_id')
        idx=next((i for i,x in enumerate(d.get('sources',[])) if x.get('source_id')==sid),None)
        if idx is None:
            d.setdefault('sources',[]).append({k:v for k,v in item.items() if k!='action'})
        else:
            d['sources'][idx].update({k:v for k,v in item.items() if k!='action'})
    for item in decisions.get('events',[]):
        match=[e for e in events if e.get('event_id')==item.get('event_id')]
        action=item.get('action')
        if action in {'exclude','remove'}:
            events[:]=[e for e in events if e.get('event_id')!=item.get('event_id')]
        elif action in {'include','confirm','promote'}:
            if not match:
                raise ValueError(f"event not found: {item.get('event_id')}")
            match[0]['status']='CONFIRMED'
            match[0]['review_decision']={k:item.get(k) for k in ('reason','source','page','excerpt') if item.get(k)}
        elif action=='update':
            if not match:
                raise ValueError(f"event not found: {item.get('event_id')}")
            match[0].update(item.get('values',{}))
            match[0]['review_decision']={k:item.get(k) for k in ('reason','source','page','excerpt') if item.get(k)}
        else:
            raise ValueError(f"unsupported event action: {action}")
    for item in decisions.get('cases',[]):
        did=item.get('deal_id'); action=item.get('action')
        if did not in cases:
            if action=='remove':
                continue
            raise ValueError(f"case not found: {did}")
        case=cases[did]
        if action in {'remove','exclude'}:
            removed_case_ids.add(did)
            events[:]=[e for e in events if e.get('deal_id')!=did]
        elif action=='update':
            for k,v in item.get('values',{}).items():
                if k=='also_kinds':
                    case[k]=v
                elif k in {'facts','fact_evidence'}:
                    case.setdefault(k,{}).update(v)
                elif k=='advisors':
                    case[k]=v
                elif k=='amounts':
                    case[k]=v
                else:
                    case[k]=v
            case['review_decision']={k:item.get(k) for k in ('reason','source','page','excerpt') if item.get(k)}
    used={e['deal_id'] for e in events}
    d['cases']=[c for c in d.get('cases',[]) if c['deal_id'] in used]
    used_sources={sid for e in events for sid in e.get('source_ids',[])}
    d['sources']=[s for s in d.get('sources',[]) if s.get('source_id') in used_sources]
    d.setdefault('limitations',[]).append('已应用显式复核决定；决定及原文锚点见 curation-decisions.json')
    return d


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('ledger',type=Path);ap.add_argument('decisions',type=Path)
    ap.add_argument('--out',required=True,type=Path);ap.add_argument('--report',type=Path)
    args=ap.parse_args()
    d=normalize(load(args.ledger),load(args.decisions));dump(args.out,d)
    report={'cases':len(d.get('cases',[])),'events':len(d.get('events',[])),
            'confirmed':sum(e.get('status')=='CONFIRMED' for e in d.get('events',[])),
            'candidate':sum(e.get('status')=='CANDIDATE' for e in d.get('events',[])),'out':str(args.out)}
    if args.report:dump(args.report,report)
    print(json.dumps(report,ensure_ascii=False))
    return 0
if __name__=='__main__':sys.exit(main())
