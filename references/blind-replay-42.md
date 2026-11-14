# 第42期严格盲测回放（2026-11-08）

> 结论先行：**结构达标，核心累计口径达标，明细台账与叙述层未达标。**
> 相比第43期盲测（`PASS 4 / FAIL 18`），v4.1.0 的年度累计桥接已修复；
> 但 Word/PPT「从零重写」与公告业务口径闸门仍是决定性缺口，产出只能称**内部待复核稿**。

## 1. 盲测纪律

- **目标**：只用第41期模板 + 官方公告，在隔离目录重建第42期（统计区间 `20261018至20261024`）。
- **封存**：`_blind42_20261108_1450/skill/` 为 14:50 冻结的 v4.1.0 快照；`workroot/` 只含第41期完整成果包（16 原生 + 3 PDF）。
- **封存校验**：生成前扫描无第42期成果、无 `GOLDEN_42` 常量；命中 `第42期` 的两处仅为第43期回放的旧期次检测字面量与测试夹具，不进入主生成链。
- **既有知识披露（诚实原则）**：第42期曾在上一轮第43期盲测中**作为模板**被读取过（文件清单、缺 `dynamic_docx` 等结构性事实）。本轮无法"忘记"这些上下文，但生成链不使用第42期任何文件或数值，且生成达到终态前未读取现有第42期成果。
- **拆封条件**：`auto-run-summary.json` 出现终态且已原子发布。首次读取现场第42期的时间为 **15:08:13**，晚于生成终态 **15:07:29**，中间先冻结了生成稿 21 个文件的 SHA-256（`generated-sha256.txt`）。

### 时间线

| 时间 | 事件 |
|---|---|
| 14:50 | 冻结 skill 快照、复制第41期模板、写入 `BLIND-DISCIPLINE.md` |
| 14:51 | `plan-only` 验证：`prev=第41期`、`span=20261018至20261024`、`out=第42期` |
| 15:00:06 | 首次在线运行（`python3` = WorkBuddy Python 3.13） |
| 15:03:31 | **BLOCKED**：`build_reports.py` 导入 `pptx` 失败（`PIL._imaging` ABI 不兼容） |
| 15:05:47 | 归档失败运行，改用原生 3.12 解释器**全新重跑**（不改任何生成逻辑） |
| 15:07:29 | 终态 `COMPLETE_WITH_EXCEPTIONS`，`published=true`，21 文件 |
| 15:07:50 | 冻结生成稿哈希 |
| 15:08:13 | 首次读取现场第42期，开始盲后对比 |

> 环境阻断属于运行基础设施故障，不是生成链的内容能力问题；归档失败运行留证后重跑，避免让系统"逃过"评测。详见 R1。

## 2. 生成结果

九阶段全部 `OK`，硬门禁 0 失败；软门禁例外 4 项（来源覆盖、待核结论字段、PDF 分页、Office 原生 PDF）。

| 阶段 | 关键输出 |
|---|---|
| ACQUIRE | 在线 6 源；`SSE_MOMENTUM/ SZSE_MOMENTUM/ CNINFO_FULLTEXT = COMPLETE`，`SSE_COMMITTEE/ CSRC_REGISTRATION = PARTIAL`，`BSE_MOMENTUM = LIMITED` |
| EXTRACT | 43 源 / 37 笔交易 / 41 条事件 / 209 项待核 |
| BUILD | 16 原生文件；5 个 Excel 追加 23 行；46 处单元格更新；3 处累计桥接 |
| RENDER | 3 PDF / 15 张逐页 PNG / 0 错误 |
| PUBLISH | 16 原生 + 3 PDF + `artifact-manifest.json` + 《本期编制与待核说明》= 21 文件 |

历史年度期号推导正确：第41期 `2025年第53期 / 2024年第117期 / 2023年第149期` → 第42期 `54 / 118 / 150`，与现场一致。

## 3. 盲后对比（`compare42.json` / `compare42.md`）

### 3.1 文件角色

| 项 | 生成稿 | 现场第42期 |
|---|---|---|
| 原生 Office | **16** | 15 |
| PDF | 3 | 1（仅周报 PDF） |
| 其它 | manifest + 编制说明 | 无 |
| 缺失角色 | **无** | — |
| 多出角色 | `dynamic_docx`（《A股市场并购重组动态》Word） | — |

现场当期**未编制**动态 Word 及其 PDF、PPT PDF；生成稿因第41期模板完整而继承生成。属角色取舍差异（R8），不是内容错误，但需要可配置开关。

### 3.2 内容体量

| 角色 | 生成稿 | 现场 | 相似度 |
|---|---|---|---|
| 周报 Word | 61 行（原始段落 115） | **1597 行** | 0.0209 |
| 周报 PDF | 4 页 | **31 页** | — |
| 市场周报 PPT | 48 行（原始形状 35） | **159 行** | 0.0315 |

### 3.3 Excel 行级差异

| 表 | 生成 / 现场行数 | 新增 / 缺失行 |
|---|---|---|
| `0` 基础数据（42 sheet） | 41/42 sheet 同名 | 229 / 231（24 个 sheet 有差异） |
| `1/2/3/8/9` 历史年度表 | 完全一致 | **0 / 0** |
| `4` 审核情况统计 | 2 sheet | 9 / 9 |
| `6` 华泰联合案例 | 4 sheet | 177 / 179 |
| `7` 控制权变更 | 5 sheet | 142 / 144 |
| `10` IPO 终止 | 6 sheet | 49 / 52 |
| `11` 引战与合作 | 3 sheet | 28 / 23 |
| `12/13` 协议转让 | 1 sheet each | 22 / 20 |

**大量"行差异"并非交易漏记，而是字段规范化问题**：例如 `11` 表本期行，现场为 `2026-10-20 00:00:00 | 中旗新材 | … | 结构化基金字段`，生成稿为 `2026-10-20 | 中旗新材 | … | 公告正文原文`；`12`/`7` 表前后 220 字符几乎相同，差异来自日期格式与 `_x000B_` 大小写。详见 R4。

## 4. 已知答案门禁（`test_known_answer_42.py`）

现场稿口径在拆封后固化，`GOLDEN` 未进入生成链。

```
PASS 22 / WARN 2 / FAIL 24  -> KNOWN_ANSWER_FAIL
```

**全部通过的核心事实（相较第43期盲测的决定性进步）**：

- 本期事件 7 项：`first_disclosed=2`（三峡新材、新亚制程）、`first_draft=1`（天津港）、`accepted=0`、`committee_passed=1`（光库科技）、`registered=0`、`terminated=0`、`simplified_passed=0`；
- 年度累计 5 项：`first_disclosed=74`、`passed=36`、`registered=32`、`terminated=66`、`monthly_reviews=37`；
- `annual != weekly`：第43期盲测中 `metrics.annual == metrics.weekly` 的架构缺陷**已修复**；
- 追加行序号为数值、无第41期残留、16 个原生角色齐备。

**失败集中在两类**：

- 明细台账勾稽 10 项：`cash_purchase 25/26`、`control_first 145/146`、`control_completed 121/122`、`strategic 223/222`、`professional 486/482`、`partial_transfer 18/16`、`current_strategic 2/1`、`current_professional 11/7`、`current_partial_transfer 5/3`、`in_review 51/50`；
- 叙述层：18 项现场正文口径仅 5 项命中（且多为证据索引中的公司名），覆盖率 **5/18**；`WARN` 2 项为 Word/PPT 体量不足。

> `in_review` 50/51 是**现场稿自身内部差异**：现场正文分板块写 17+16+3+15+0=51 单，工作簿"是否在审=是"勾稽为 50 单。生成稿取 51。第43期也存在同类 49/48 差异。须由业务确认后固定单一口径，不能静默取一。

## 5. 差异根因

| # | 根因 | 证据 |
|---|---|---|
| **R1** | `build_reports.ensure_libs()` 无条件把 **Python 3.12 site-packages 注入 3.13 进程**，`python-pptx → PIL.Image → _imaging` C 扩展 ABI 不兼容 | 首次在线运行 `BLOCKED` 的 `FATAL.txt` 完整栈 |
| **R2** | `build_reports.py` 设计上 `"replaces stale template prose wholesale"`，**从零重写 Word/PPT**，不使用上期模板版式、表格、图表 | Word 115/1597 行、PPT 35/159 行、PDF 4/31 页；18 项正文口径仅 5 项命中 |
| **R3** | 公告**业务口径闸门缺失**：宽口径关键词召回即建模，未按工作手册区分"认购私募基金份额 / 与专业机构共同投资""协议转让 / 控制权变更""筹划停牌 / 预案首次披露" | `current_professional 11 vs 7`、`current_partial_transfer 5 vs 3`、`strategic 223 vs 222`；同时控制权与现金购买**漏记** 1 单 |
| **R4** | 台账**字段未规范化**：把公告正文原文直接写入单元格，日期写成字符串而非 datetime，公告标题/交易结构未按表头拆分 | `11/12/7` 表行级样例；生成稿单元格含"证券代码：001212 证券简称：…"整段正文 |
| **R5** | 追加明细行后**汇总区公式缓存未重算** | `0` 表沪主板统计：生成 `24 单 / 11 单已披露顾问 / 12 单未披露`，现场 `25 / 12 / 13`；而按代码去重行数正确（74） |
| **R6** | 在审口径存在正文数与工作簿勾稽数的**内部差异**，生成链未显式登记 | 现场 51/50；生成稿 51 |
| **R7** | Word 叙述把 `2 单首次披露 + 1 单首次披露草案（天津港）` 合并写成"3 单"；现场把天津港单列于"八、本期新披露案例简介"且不计入首次披露单数 | 事件台账 `first_disclosed=2 / first_draft=1` 分层正确，仅表述合并 |
| **R8** | 文件角色按模板全量继承，缺按当期取舍的开关 | 生成稿多出 `dynamic_docx` 及 2 份 PDF |

## 6. 解决对策

### P0（阻断级 / 本轮已实施）

1. **禁止跨解释器 ABI 混用**：`ensure_libs()` 改为只接受与当前 `sys.version_info[:2]` 匹配的 site-packages，否则抛清晰错误。
2. **编排器环境自举**：`auto_run.py` 入口新增 `ensure_runtime()`，主链依赖（`openpyxl/docx/pptx/pypdf/lxml`）不齐时，在同机探测合格解释器并 `os.execve` 重新执行自身，使所有子脚本继承同一解释器；找不到则 `BLOCKED` 并给出安装提示。支持 `MA_WEEKLY_PYTHON` 显式指定。
3. **固化第42期已知答案门禁**：新增 `scripts/test_known_answer_42.py`，与第43期门禁并列，作为连续三期回放的第二个锚点。

### P1（内容正确性，下一阶段实施）

| 对策 | 交付物 | 验收标准 |
|---|---|---|
| 叙述层改为「上期模板保真 + 局部注入」 | `narrative_preserve.py`，复用 `lib/ooxml.py::_patch_text_nodes()` 与 `build_package.py` 的 OOXML 局部补丁 | 周报 Word ≥ 现场 90% 段落数、PDF 页数差 ≤ 2 页、18 项叙述口径覆盖率 ≥ 90% |
| 公告业务口径闸门 | `business_gate.py`，规则取自《20231027并购周报统计工作手册 v1.0》并配置化 | `current_professional / current_partial_transfer / current_strategic / control_*` 等 10 项勾稽全通过 |
| 台账字段规范化 | `field_norm.py`，按表头映射输出日期(datetime)、标题、结构、金额、顾问 | 新增行不含公告正文整段；`12` 表日期为 datetime；行级差异降为纯事实差异 |
| 汇总区重算 | `summary_recalc.py` 或固定 Office 重算通道 | 汇总单元格值与明细勾稽一致；`formula_fingerprint` 升级为"汇总=明细"一致性检查 |
| 门禁接入主链 | `quality_gate.py` / `VALIDATE` 阶段自动跑已知答案门禁，结果写入 `validation-report.json` | 门禁未全绿时强制 `COMPLETE_WITH_EXCEPTIONS` 并在说明首页标"事实未核验完毕" |

### P2（工程与发布）

- 角色开关配置（是否出动态 Word、PPT PDF），按现场当期取舍装配。
- 在审口径双口径登记：正文数 vs 工作簿数差异写 `pending-review.json`，不静默取一。
- Word 首次披露表述分列："本期首次披露 2 单，另有 1 单首次披露草案（天津港）"。
- 接入 Microsoft Office 固定发布环境 / 原生 PDF 通道，解决 LibreOffice 分页差异。
- **连续三期（41/42/43）已知答案回放全绿**，方可把产出称为可发布稿。

## 7. 本次 skill 变更（v4.1.0 → v4.1.1）

| 文件 | 变更 |
|---|---|
| `scripts/build_reports.py` | `ensure_libs()` 同版本探测 + 清晰错误，禁止跨 minor 版本注入 |
| `scripts/auto_run.py` | 新增 `RUNTIME_MODULES` / `_runtime_ok()` / `ensure_runtime()`，入口自动选择合格解释器 |
| `scripts/test_known_answer_42.py` | 新增，第42期已知答案门禁 |
| `references/blind-replay-42.md` | 新增本报告 |
| `SKILL.md` | 版本 4.1.1；第 8 节增加第42期盲测结论与"核心累计达标 / 明细与叙述未达标"分层表述 |
| `references/unattended.md` | 验收层增加第42期门禁与运行环境自检说明 |

> 注：主 skill 在 15:02–15:03 另有并发改动（`max_pdf` 300→400、BUILD 阶段纳入 `build_reports.py` 幂等键、第43期门禁新增 1900 年日期样式检查），本次全部保留。本轮盲测基于 **14:50 冻结的 v4.1.0 快照**，上述并发改动不影响盲测生成内容。

## 8. 仍未完成

- 上述 P1 五项均未实现；`narrative_preserve.py` 是内容达标的前提。
- 已知答案门禁仍是**旁路工具**，尚未接入 `VALIDATE` 主链。
- 第41期回放、连续三期无人工修改尚未执行。
- 深交所正确 catalog、北交所/证监会注册源仍为 `LIMITED/PARTIAL`。
- Office 原生 PDF 发布环境、Excel 公式在最终 Office 环境重算仍缺失。

## 9. 复现命令

```bash
cd "<工作区根>/_blind42_20261108_1450"

# 在线重建（v4.1.1 起会自动选择合格解释器；冻结快照需显式指定 3.12）
python3 ../../ma-weekly-report/scripts/auto_run.py \
  --as-of 2026-10-24 --root workroot --work-root runs --out-root out --max-pdf 20

# 盲后对比（仅终态后执行）
python3 compare42.py --generated "out/并购周报基础数据-2026年第42期-20261018至20261024" \
  --reference "../../并购周报基础数据-2026年第42期-20261018至20261024" \
  --out compare42.json

# 已知答案门禁
python3 ../../ma-weekly-report/scripts/test_known_answer_42.py \
  --run-dir runs/2026-10-24 \
  --package "out/并购周报基础数据-2026年第42期-20261018至20261024" \
  --report known-answer-42.json
```
