#!/usr/bin/env python3
import unittest
from extract_filings import classify_title, _extract_control_facts, canonical_advisor


class TitleClassificationTests(unittest.TestCase):
    def events(self, title):
        return {(event, kind) for event, kind, _confidence in classify_title(title)}

    def test_plan_and_draft_are_distinct_milestones(self):
        self.assertIn(('first_disclosed', 'ma_audit'), self.events(
            '发行股份及支付现金购买资产并募集配套资金暨关联交易预案'))
        self.assertIn(('first_draft', 'ma_audit'), self.events(
            '发行股份及支付现金购买资产并募集配套资金暨关联交易报告书（草案）'))

    def test_suspension_notice_is_pipeline_lead_only(self):
        events = self.events('关于筹划发行股份购买资产并募集配套资金事项的停牌公告')
        self.assertIn(('lead_signal', 'ma_audit'), events)
        self.assertNotIn(('first_disclosed', 'ma_audit'), events)

    def test_share_plus_cash_suspension_is_not_cash_ma(self):
        events = self.events('关于筹划发行股份及支付现金方式购买资产并募集配套资金事项的停牌公告')
        self.assertIn(('lead_signal', 'ma_audit'), events)
        self.assertNotIn(('lead_signal', 'ma_cash_purchase'), events)

    def test_control_planning_suspension_is_not_first_disclosure(self):
        events = self.events('关于筹划公司控制权变更事项的停牌公告')
        self.assertNotIn(('first_disclosed', 'control'), events)
        self.assertNotIn(('control_progress', 'control'), events)

    def test_parent_level_strategic_investor_is_not_subsidiary_investment(self):
        events = self.events('农业银行关于引入战略投资者并签署《附条件生效的战略合作协议》的公告')
        self.assertNotIn(('first_disclosed', 'strategic_investment'), events)

    def test_only_surviving_share_swap_issuer_gets_registration(self):
        self.assertNotIn(('registered', 'ma_audit'), self.events(
            '信达证券股份有限公司关于中国国际金融股份有限公司换股吸收合并东兴证券股份有限公司、信达证券股份有限公司等事项获得中国证券监督管理委员会同意注册及核准批复的公告'))
        self.assertIn(('registered', 'ma_audit'), self.events(
            '中金公司关于公司换股吸收合并东兴证券股份有限公司、信达证券股份有限公司等事项获得中国证券监督管理委员会同意注册及核准批复的公告'))

    def test_control_change_completed_title_is_explicit(self):
        self.assertIn(('control_completed', 'control'), self.events('关于公司控制权变更完成的公告'))

    def test_transfer_intent_is_not_first_disclosure(self):
        events = self.events('关于控股股东签署《股份转让意向协议》暨控制权拟发生变更的提示性公告')
        self.assertNotIn(('first_disclosed', 'control'), events)
        self.assertNotIn(('control_progress', 'control'), events)

    def test_completed_share_transfer_recognition(self):
        self.assertIn(('control_completed', 'control'), self.events(
            '股份协议转让事项完成股份过户登记手续暨控制权发生变更的公告'))

    def test_upper_rights_change_with_new_license_is_completion(self):
        self.assertIn(('control_completed', 'control'), self.events(
            '关于实际控制人变更暨控股股东的上层权益变动并换发营业执照的公告'))

    def test_investment_fund_establishment_is_professional_cooperation(self):
        self.assertIn(('first_disclosed', 'professional_cooperation'), self.events(
            '关于参与设立河南省并购母基金的公告'))

    def test_public_tender_for_controlling_shares_is_control_disclosure(self):
        self.assertIn(('first_disclosed', 'control'), self.events(
            '关于控股股东拟协议转让公司股份公开征集受让方的公告'))

    def test_fund_share_transfer_out_is_not_new_cooperation(self):
        self.assertNotIn(('first_disclosed', 'professional_cooperation'), self.events(
            '关于转让部分私募基金认购份额暨关联交易的公告'))

    def test_terminated_cash_share_acquisition_gets_cash_kind(self):
        self.assertIn(('terminated', 'ma_cash_purchase'), self.events(
            '关于终止筹划股权收购事项的提示性公告'))

    def test_question_reply_is_not_review_state(self):
        events = self.events('关于回复上海证券交易所《关于对公司重大资产重组草案的信息披露问询函》的公告')
        self.assertNotIn(('inquiry', 'ma_audit'), events)
        self.assertNotIn(('replied', 'ma_audit'), events)

    def test_terminated_planning_is_not_formal_termination(self):
        events = self.events('关于终止筹划发行股份及支付现金购买资产事项暨公司股票复牌的公告')
        self.assertNotIn(('terminated', 'ma_audit'), events)
        self.assertNotIn(('lead_signal', 'ma_audit'), events)

    def test_company_name_change_is_not_control_completion(self):
        events = self.events('关于控股股东、实际控制人之一致行动人变更名称及经营范围暨完成工商变更登记的公告')
        self.assertNotIn(('control_completed', 'control'), events)

    def test_consistent_action_administrator_change_is_not_control_completion(self):
        events = self.events('关于实际控制人的一致行动人完成工商变更登记的公告')
        self.assertNotIn(('control_completed', 'control'), events)

    def test_named_investment_company_capital_increase_is_not_fund_cooperation(self):
        events = self.events('关于对安徽省新能创业投资有限责任公司增资暨关联交易的公告')
        self.assertNotIn(('first_disclosed', 'professional_cooperation'), events)

    def test_consistent_action_expiry_is_substantive_control_disclosure(self):
        events = self.events('关于一致行动协议到期终止及实际控制人变更暨权益变动的提示性公告')
        self.assertIn(('first_disclosed', 'control'), events)

    def test_combined_cash_purchase_and_sale_create_two_role_events(self):
        events = self.events('重大资产出售及重大资产购买暨关联交易报告书（草案）')
        self.assertIn(('first_draft', 'ma_cash_purchase'), events)
        self.assertIn(('first_draft', 'ma_cash_sale'), events)
        self.assertNotIn(('first_draft', 'ma_audit'), events)

    def test_refinancing_registration_is_not_ma_registration(self):
        self.assertNotIn(('registered', 'ma_audit'), self.events(
            '关于向特定对象发行股票申请获得中国证监会同意注册批复的公告'))
        self.assertIn(('registered', 'ma_audit'), self.events(
            '关于发行股份及支付现金购买资产并募集配套资金事项获得中国证监会同意注册批复的公告'))

    def test_interim_takeover_report_is_not_new_control_event(self):
        self.assertNotIn(('first_disclosed', 'control'), self.events(
            '董事会关于某公司要约收购事宜致全体股东的报告书'))

    def test_external_strategic_investment_is_in_scope(self):
        self.assertIn(('first_disclosed', 'strategic_investment'), self.events(
            '关于全资孙公司增资扩股引入外部投资者的公告'))
        self.assertNotIn(('first_disclosed', 'strategic_investment'), self.events(
            '关于使用募集资金向全资子公司增资以实施募投项目的公告'))

    def test_professional_cooperation_excludes_lifecycle_notices(self):
        self.assertIn(('first_disclosed', 'professional_cooperation'), self.events(
            '关于与专业投资机构共同投资的公告'))
        self.assertNotIn(('first_disclosed', 'professional_cooperation'), self.events(
            '关于私募基金清算注销暨关联交易的公告'))


class ControlAndAdvisorTests(unittest.TestCase):
    def test_negative_actual_controller_statement_wins(self):
        facts = _extract_control_facts(
            '本次权益变动完成后，控股股东变更为新集团，上市公司实际控制人仍为某市国资委。')
        self.assertEqual(facts['control_result'], '未导致实际控制人变化')

    def test_same_day_negative_statement_does_not_hide_actual_transition(self):
        facts = _extract_control_facts(
            '徐晰人不再为公司实际控制人，上市公司的实际控制人变更为程栋。')
        self.assertEqual(facts['control_result'], '实际控制人发生变化')
        self.assertEqual(facts['old_controller'], '徐晰人')
        self.assertEqual(facts['new_controller'], '程栋')
        self.assertTrue(facts['completion_confirmed'])

    def test_partial_completion_is_progress_not_full_control_completion(self):
        facts = _extract_control_facts(
            '本次协议转让已完成过户登记。本次控制权变更的交易方案由协议转让、间接股权转让共同构成，'
            '目前协议转让事项已经办理完毕，间接股权转让尚需办理工商登记，孙凯先生将成为实际控制人。')
        self.assertEqual(facts['control_result'], '控制权拟发生变更')
        self.assertFalse(facts.get('completion_confirmed', False))

    def test_completed_transfer_uses_actual_effective_date(self):
        facts = _extract_control_facts(
            '股份转让事宜已完成过户登记。第一期标的股份过户日期为2026年10月15日。'
            '上市公司控股股东变更为极宁科技，实际控制人变更为潘涛先生。')
        self.assertEqual(facts['control_result'], '实际控制人发生变化')
        self.assertEqual(facts['new_controller'], '潘涛')
        self.assertEqual(facts['completion_date'], '2026-10-15')

    def test_advisor_aliases_are_canonicalized(self):
        self.assertEqual(canonical_advisor('浙商证券股份有限公司'), '浙商证券')
        self.assertEqual(canonical_advisor('中国银河证券股份有限公司'), '银河证券')


if __name__ == '__main__':
    unittest.main(verbosity=2)
