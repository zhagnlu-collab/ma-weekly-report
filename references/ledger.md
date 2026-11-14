# 事件台账接口与脚本边界

中间JSON是计算底稿，不替代现有Excel交付。先人工/工具从完整公告核验字段，再录入。数组顺序不代表时序，日期ISO `YYYY-MM-DD`，金额采用万元小数（建议十进制字符串），比例0—1数值，币种CNY/HKD/USD等明确标识。

## 顶层结构

```json
{
  "period": {"start":"2026-11-01","end":"2026-11-07","cutoff":"2026-11-07"},
  "sources": [{
    "source_id":"source-a", "published_on":"2026-11-02",
    "file_or_url":"已保存官方公告.pdf", "locator":"交易方案章节/实际页码",
    "excerpt":"实际公告摘录；示例不代表真实披露"
  }],
  "cases": [{
    "deal_id":"issuer-a-target-a-2026", "code":"000001.SZ", "name":"示例公司",
    "board":"深主板", "kind":"ma_audit",
    "advisors":[{"name":"示例券商","role":"independent","disclosed_on":"2026-11-02","source_ids":["source-a"]}],
    "amounts":[{"basis":"initial_plan","value_wan":"12500.50","currency":"CNY","known_on":"2026-11-02","source_ids":["source-a"]}]
  }],
  "events": [{
    "event_id":"event-a", "deal_id":"issuer-a-target-a-2026", "type":"first_disclosed",
    "date":"2026-11-02", "announced_on":"2026-11-02", "sequence":0,
    "status":"CONFIRMED", "source_ids":["source-a"]
  }],
  "coverage":[{"group":"审核类首次发现/三所","status":"PARTIAL","evidence":"检索记录文件＋实际查询范围"}]
}
```

示例为结构展示，股票代码不代表该公司实际交易，不能原样作为真实事件使用。

## 枚举及含义

`kind`：ma_audit、ma_cash_purchase、ma_cash_sale、control、partial_transfer、second_shareholder、ipo、strategic_investment、professional_cooperation。

同一协议转让既满足12又满足13时，不复制deal：主kind为partial_transfer，`also_kinds:["second_shareholder"]`，反向也可；互为条件现金购买/出售可按同一组合deal分类型标识或保留独立leg_id并在汇总经济交易时去重。脚本允许cash_purchase与cash_sale的also_kinds，不能把分组计数直接加总当市场唯一数。其他多类型关系另行明确映射，不任意用标签扩大范围。

`type`：
- first_disclosed、first_draft：独立里程碑；每个交易原则上只有一次，各公告/更正附在同一来源组。
- accepted、inquiry、replied、suspended、resumed：在审状态变化。
- committee_scheduled：计划，**不是审议结果**。
- committee_passed、committee_deferred、committee_rejected：实际结果；同时填写`meeting_id`（交易所＋年度＋届次）。会议date采用会议日，announced_on采用实际结果公告日。
- simplified_passed、registration_submitted、registered、terminated。
- control_completed：控制权已变更正式完成公告日；事件实际生效日另加`effective_on`，不会自动代替统计日期。
- control_progress：完成部分过户、仅收到过户确认或其他间接交易尚未完成、控制权仍为拟变更的阶段性进展；只更新最新公告日/进展，不进入完成表，也不作为新的首次披露。
- completed、progress：其他交易完成和进展。

`status`：CONFIRMED、CANDIDATE、UNVERIFIABLE；后两类不进正式统计，输出事件ID待核。

同一天多状态需要`sequence`明确先后，不能凭输入数组顺序决定；同一结果多份文件合并source_ids，不建立重复event。首次披露与首次草案同日可以是两个不同类型的事件，但统计首次披露交易数只计一单。

顾问角色：independent（脚本主排名采用）、financial、acquirer。最后两类保留台账供模板特殊角色另行归集，脚本不会自动计入独立财顾。未公开顾问不造disclosed_on；对公开状态不确定的顾问不填入advisors正式数组，可用pending_advisors辅助字段。

金额：initial_plan、latest_plan、actual_completion 分阶段、分currency、分known_on存。通用并购汇总用截止日最新计划；**control_completed汇总仅用实际完成金额，不回退计划上限**。缺人民币金额不以0代替，输出known_cny_amount_wan及unknown_cny_amount_deals；前者可能为0且后者非空，表示只有已知小计为0、不能宣称总额0。多币种实际金额需先分笔核验并统一来源，不在脚本中隐式折汇。

## 调用

```bash
python scripts/weekly_ledger.py ledger.json --out work/metrics.json
python scripts/inspect_templates.py 上期目录 --out work/template-inventory.json
python scripts/test_weekly_ledger.py
```

脚本不联网、不提取公告、不判断法规、不自动写入Excel，不替代Office渲染。`validation=PASS`仅证明输入结构和预设算术可执行；`coverage_declaration_complete`仅复述用户填入的覆盖状态，并非真实完整性认证。把metrics的可追溯deal_ids和原始source_ids映射到实际单元格/段落，再进行业务复核。

脚本暂不自动计算财务PE/PB、实际控制链、全部表注的特殊财顾家次、IPO细分阶段或外币折算。这些按参考口径逐项核验，不能因为脚本没有相应字段就省略成果。
