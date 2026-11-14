#!/usr/bin/env python3
"""合成数据行为测试；不访问网络、不修改真实周报。"""
from copy import deepcopy
import unittest
from weekly_ledger import compute, validate

class LedgerTest(unittest.TestCase):
    def setUp(self):
        self.d={'period':{'start':'2026-11-01','end':'2026-11-07','cutoff':'2026-11-07'},'sources':[],'cases':[],'events':[],'coverage':[{'group':'synthetic-only','status':'PARTIAL','evidence':'unittest fixture;不是真实检索'}]}
    def source(self,date):
        sid='s-'+date
        if not any(x['source_id']==sid for x in self.d['sources']):
            self.d['sources'].append({'source_id':sid,'published_on':date,'file_or_url':'fixture.pdf','locator':'正文交易方案','excerpt':'合成测试证据，不代表真实公告'})
        return sid
    def case(self,cid,kind='ma_audit',board='深主板'):
        c={'deal_id':cid,'code':'920000.BJ' if board=='北交所' else '000001.SZ','name':'示例公司','board':board,'kind':kind,'advisors':[],'amounts':[]};self.d['cases'].append(c);return c
    def event(self,cid,typ,date,**kw):
        announced=kw.pop('announced_on',date)
        e={'event_id':'e'+str(len(self.d['events'])),'deal_id':cid,'type':typ,'date':date,'announced_on':announced,'status':'CONFIRMED','source_ids':[self.source(announced)]};e.update(kw);self.d['events'].append(e);return e
    def advisor(self,c,name,role='independent',when='2026-11-01'):
        c['advisors'].append({'name':name,'role':role,'disclosed_on':when,'source_ids':[self.source(when)]})
    def amount(self,c,value,basis='initial_plan',currency='CNY',when='2026-11-01'):
        c['amounts'].append({'value_wan':str(value),'basis':basis,'currency':currency,'known_on':when,'source_ids':[self.source(when)]})
    def test_sunday_and_bse_are_included(self):
        self.case('a',board='北交所');self.event('a','first_disclosed','2026-11-01')
        r=compute(self.d);self.assertEqual(r['weekly']['first_disclosed']['by_board'],{'北交所':1});self.assertFalse(r['coverage_declaration_complete'])
    def test_multiple_advisors_dont_duplicate_market_amount(self):
        c=self.case('a');self.amount(c,'100');self.advisor(c,'甲');self.advisor(c,'乙');self.advisor(c,'甲');self.advisor(c,'丙','financial');self.advisor(c,'丁',when='2026-11-08');self.event('a','first_disclosed','2026-11-02')
        r=compute(self.d)['weekly']['first_disclosed'];self.assertEqual(r['projects'],1);self.assertEqual(r['advisor_instances'],2);self.assertEqual(r['known_cny_amount_wan'],'100')
    def test_cross_year_and_pass_exit_review(self):
        self.case('old');self.event('old','first_disclosed','2025-01-01');self.event('old','accepted','2025-12-30');self.event('old','committee_passed','2026-11-04',meeting_id='SZ-2026-10')
        r=compute(self.d);self.assertEqual(r['annual']['first_disclosed']['projects'],0);self.assertEqual(r['annual']['committee_passed']['projects'],1);self.assertEqual(r['as_of_in_review']['projects'],0)
    def test_deferred_then_passed_and_simplified(self):
        self.case('a');self.event('a','committee_deferred','2026-11-02',meeting_id='SH-2026-1');self.event('a','committee_passed','2026-11-06',meeting_id='SH-2026-2');self.case('b');self.event('b','simplified_passed','2026-11-07')
        r=compute(self.d)['weekly'];self.assertEqual(r['committee_meetings'],2);self.assertEqual(r['committee_appearances'],2);self.assertEqual(r['committee_unique_reviewed'],1);self.assertEqual(r['approved_including_simplified']['projects'],2)
    def test_future_knowledge_and_unconfirmed_excluded(self):
        self.case('a');self.event('a','accepted','2026-11-03');self.event('a','committee_passed','2026-11-06',announced_on='2026-11-08',meeting_id='SZ-2026-1');self.case('b');self.event('b','first_disclosed','2026-11-03',status='CANDIDATE')
        r=compute(self.d);self.assertEqual(r['as_of_in_review']['projects'],1);self.assertEqual(r['weekly']['committee_passed']['projects'],0);self.assertEqual(len(r['unconfirmed_events']),1)
    def test_control_completion_week_uses_announcement_date(self):
        c=self.case('control',kind='control')
        e=self.event('control','control_completed','2026-11-06')
        e['announced_on']='2026-11-06';e['effective_on']='2026-10-29'
        r=compute(self.d)['weekly']['control_completed']
        self.assertEqual(r['projects'],1)

    def test_control_plan_and_actual_amount_are_separate(self):
        c=self.case('control',kind='control');self.amount(c,'140719.41');self.event('control','first_disclosed','2026-11-01');self.event('control','control_completed','2026-11-06')
        r=compute(self.d)['weekly']['control_completed'];self.assertEqual(r['unknown_cny_amount_deals'],['control']);self.assertEqual(r['known_cny_amount_wan'],'0')
        self.amount(c,'50000',basis='actual_completion',when='2026-11-06');self.assertEqual(compute(self.d)['weekly']['control_completed']['known_cny_amount_wan'],'50000')
    def test_foreign_currency_is_not_cny_zero(self):
        c=self.case('a');self.amount(c,'0.0001',currency='HKD');self.event('a','first_disclosed','2026-11-01');r=compute(self.d)['weekly']['first_disclosed'];self.assertEqual(r['unknown_cny_amount_deals'],['a'])
    def test_shared_transaction_can_enter_both_transfer_tables(self):
        c=self.case('t',kind='partial_transfer');c['also_kinds']=['second_shareholder'];self.event('t','first_disclosed','2026-11-02');r=compute(self.d)['weekly'];self.assertEqual(r['partial_transfer_first_disclosed']['deal_ids'],['t']);self.assertEqual(r['second_shareholder_first_disclosed']['deal_ids'],['t'])
    def test_duplicate_first_milestone_is_rejected(self):
        self.case('a');self.event('a','first_disclosed','2026-11-01');self.event('a','first_disclosed','2026-11-02');self.assertTrue(any('首次里程碑重复' in x for x in validate(self.d)))
    def test_ambiguous_same_day_state_requires_sequence(self):
        self.case('a');self.event('a','accepted','2026-11-01');self.event('a','inquiry','2026-11-01');self.assertTrue(validate(self.d));self.d['events'][-1]['sequence']=1;self.assertEqual(validate(self.d),[])
    def test_announcement_date_not_signing_day_for_registration(self):
        self.case('a');self.event('a','registered','2026-11-01',announced_on='2026-11-02');self.assertTrue(any('公告日' in x for x in validate(self.d)))
    def test_missing_evidence_and_undisclosed_advisor_are_rejected(self):
        c=self.case('a');self.advisor(c,'甲（未披露）');e=self.event('a','first_disclosed','2026-11-01');e['source_ids']=[];self.assertGreaterEqual(len(validate(self.d)),2)

if __name__=='__main__':unittest.main(verbosity=2)
