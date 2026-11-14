# v4.0.0 无人值守运行手册（日期进、成品出）

## 1. 唯一输入

```bash
python scripts/auto_run.py --as-of 2026-11-07
```

`--as-of` 是唯一必需参数。程序不提问、不等人、不发消息、不点击。其余参数都是可选的诊断/隔离开关：

| 参数 | 用途 |
|---|---|
| `--root` | 工作区根目录；默认自动探测（skill 父目录 → cwd 及其上级中首个含期次目录的根） |
| `--work-root` | 运行目录根，默认 `<root>/.weekly-runs` |
| `--out-root` | 成品发布根目录，默认 `<root>` |
| `--prev` | 显式指定上期目录（仅诊断用，正常不需要） |
| `--config` | JSON 配置，见 `config/unattended.example.json` |
| `--soffice` / `--fontconfig` | 固定渲染器与字体配置 |
| `--pdf-source-dir` | 已有 Office 原生导出的 3 份 PDF 时直接装配 |
| `--max-pdf` | 本期最多下载的公告原件数，默认 400，按公告角色优先 |
| `--offline` | 结构回放，不联网（覆盖矩阵全为 NOT_RUN） |
| `--fresh` | 另起带时间戳的运行目录 |
| `--retries` | 单阶段失败重试次数，默认 1 |
| `--publish-policy` | `with-exceptions`（默认）/ `complete-only` |
| `--plan-only` | 只做期次与路径推导 |

环境变量 `MA_WEEKLY_SOFFICE`、`MA_WEEKLY_FONTCONFIG` 可替代同名参数；渲染器按「显式参数 → 配置 → 已知 bundled 路径 → PATH → LibreOffice.app」顺序自动发现。

## 2. 日期语义与推导

- 时区 `Asia/Shanghai`，截止时点当日 `23:59:59`，闭区间。
- 本期起点 = **最新完整上期结束日 + 1 天**，终点 = `as-of`。不是 `as-of - 6`，也不是周一至周六。
- 模板默认取结束日早于 `as-of`、且 16 个角色齐全的最新一期；不完整上期会出件但在 `plan.json` 标注缺口。
- 当年期号 +1；历史年度库文件自身期号也 +1（第 43 期的 2025 年第 55 期 → 第 44 期 2025 年第 56 期）。
- 文件名中的 `YYYYMMDD至YYYYMMDD` 与期号按真实文件名映射；输出目录名 `并购周报基础数据-{年}年第{N}期-{起}至{止}`。
- 跨年时当年期号重置为第 1 期并写入 `notes`，不静默假定新年度初始化口径；`as-of` 非周六也写入 `notes`。

## 3. 运行目录布局

```text
<work-root>/<as-of>/
  run-state.json              阶段状态、输入指纹、最终状态
  plan.json / period.json     计划与期次快照
  prev-fingerprints.json      上期 16 个原生文件 SHA-256
  sources/                    官方源原始响应、PDF、网页快照
  source-manifest.json        来源 URL / 方法 / 访问时间 / 哈希 / 页数 / 覆盖矩阵
  event-ledger.json           Deal / Event / Evidence / Coverage
  ledger.sqlite               同台账的可查询关系表
  pending-review.json         结论性待核字段与未归类候选
  metrics.json                周度 / 年度 / 截止日指标快照（含 validation）
  content/                    16 个原生 Office 文件（生成中）
  build-report.json           文本替换命中、累计数字桥接、追加行审计
  building/                   候选成果包（16 原生 + 3 PDF）
  render-report.json / visual/ PDF 与逐页 PNG
  validation-report.json      硬/软门禁明细
  本期编制与待核说明.md/.docx
  publish-report.json / auto-run-summary.json / blocked-report.json
  stage-logs/                 每条子命令与异常堆栈
```

成品目录额外包含 `artifact-manifest.json`（逐文件 SHA-256、状态、修订号、例外清单）。

## 4. 阶段契约与断点续跑

| 阶段 | 输入 | 输出 | 幂等键 |
|---|---|---|---|
| PLAN | as-of、上期目录 | plan/period/指纹 | 日期 + 上期 + period.py 哈希 |
| ACQUIRE | 起止日、max_pdf | sources/、覆盖矩阵 | 起止日 + 参数 + 缓存有无 + 脚本哈希 |
| EXTRACT | sources/、manifest | 事件台账、SQLite、待核清单 | 来源树哈希 + manifest + 脚本哈希 |
| METRICS | 事件台账 | metrics.json | 台账哈希 + weekly_ledger.py 哈希 |
| BUILD | 上期 16 文件、台账、期次 | content/、build-report | 台账 + period + 指纹 + 生成脚本哈希 |
| RENDER | content/、渲染器 | building/、逐页 PNG、render-report | content 哈希 + 渲染器/字体/PDF 源 |
| VALIDATE | building/ + 各报告 | validation-report | 包哈希 + 上游报告 + 门禁脚本哈希 |
| NOTES | 门禁、待核、来源 | 编制与待核说明 | 状态 + 报告哈希 |
| PUBLISH | building/ + 说明 | 正式成果目录 + manifest | 状态 + 包哈希 + 说明 + 输出根目录 |

阶段成功后写 `status=OK` 与 `input_key`；重跑命中同一键即跳过，不会重复插入交易、重复下载或重复发布。失败只重跑失败阶段；任一阶段彻底失败写 `blocked-report.json` 并把整体状态置 `BLOCKED`。

## 5. 数据契约与来源

`acquire_sources.py` 纯标准库直连官方入口，保存原始响应、URL、访问时间、SHA-256、页数与记录数：

- 上交所并购重组项目动态 `commonExcelKcb.do?sqlId=GP_BGCZ_XMLB`（结构化 `.xls`）
- 上交所并购重组委会议公告与结果
- 深交所 `www.szse.cn/api/report/ShowReport/data`
- 巨潮 `www.cninfo.com.cn/new/hisAnnouncement/query`（多关键词逐页取全，下载公告原件）
- 北交所、证监会注册信息（best-effort 网页快照）

覆盖状态 `COMPLETE / PARTIAL / LIMITED / NOT_RUN / NOT_APPLICABLE` 逐组记录。**列表页只产生候选；零结果不得解释为无案例。**单一来源失败不阻断整条流水线，但会在软门禁留下例外。

`extract_filings.py` 把候选归为 `Deal / Event / Evidence / Coverage`：

- 同一交易多份公告共享 `deal_id`；同一公司不同标的按公告标识区分，不按公司名粗暴去重。
- 事件带 `date / announced_on / type / status / sequence / source_ids / meeting_id`。
- 公告全文用 `pymupdf` 抽取金额/比例；抽不到不填、不猜。
- 是否构成重大资产重组、控制权结论、独立财顾角色、金额口径等结论性字段进入 `pending-review.json`。
- SQLite 表：`deals`、`events`、`event_sources`、`evidence`、`coverage`。

`weekly_ledger.py` 从 **已核验事件**生成指标快照；脚本成功不等于外部事实核验完成。若台账不满足口径契约，`metrics.json` 写 `validation=FAIL` 并成为软门禁例外，绝不伪造指标。

## 6. Office 生成与保真边界

- `lib/ooxml.py` 用 ZIP/XML 字节级局部补丁：只改目标 `<t>` 文本节点、指定单元格、追加行与 `styles.xml` 绿色填充，其余部件（批注、drawing、chart、externalLink、公式、合并区、修订、母版）原样复制。现场 13 个 Excel 含批注与外链，openpyxl 整本保存会破坏这些对象，故主链路不使用。
- 跨 run 的替换先拼接同部件所有 `<t>` 节点、求不重叠替换区间，再按每个节点原始字符范围重建：替换文本写进区间起点所在节点，被覆盖的其余字符全部删除，避免「截至2026年11月7日2026年10月31日」这类残留。
- 替换表只含明确的期次、统计区间、封面口径与「截至」句式锚点；业务正文中的历史案例日期、金额、批文号一律不动。独立「第N期」用负向后顾正则，避免误改历史年度期号。
- 年度累计按 **上期 Word/PPT 正文基数 + 本期已核实增量** 桥接，不用本期台账重算覆盖；上期正文找不到该统计句式时跳过并记录。
- 本期事件按真实表头列名（股票简称/代码/首次披露时间/板块等）追加到对应编号 Excel，新增行用 `RGB(226,239,218)` / `FFE2EFDA` 标记。只写可回源字段，不覆盖人工补充列。
- **公式不重算**：openpyxl 不计算公式，缓存值沿用上期；涉及公式的汇总须在最终 Office 环境重算核对，说明中必须写明。

## 7. PDF、视觉回归与质量门禁

默认用固定 bundled LibreOffice + `FONTCONFIG_FILE` 导出 3 份 PDF，再由 `pymupdf` 逐页渲染 PNG，统计页数、中文字符、空白页与中文缺字页，并与上期 PDF 页数比对。LibreOffice 与 Microsoft Word 发布版分页可能不同，脚本如实记录 `limitation`，**不把 PDF 可打开当版式通过**。若已有 Office 原生导出 PDF，用 `--pdf-source-dir` 装配并同样逐页检查。

硬门禁（失败即 `BLOCKED`，禁止发布）：16 个原生角色齐全、3 份 PDF、Office ZIP 完整、Excel 可打开、公式缓存可扫描、**本期未新增** Excel 公式错误值、文件名期次正确、上期原件哈希未变、无临时文件、事件台账无重复、PDF 可打开且逐页渲染。当 `render-report.json` 声明 `pdf_source=office-native` 时，额外硬校验 `office-pdf/office-export-manifest.json` 三份登记齐备。

公式检查按「文件角色/工作表名/单元格坐标」归一化后与上期集合做差：上期人工稿里普遍存在 `#N/A` 等 VLOOKUP 无匹配的缓存值，属于既有状态，不能拿来阻断本期；只有本期新增的错误值或无缓存公式才分别进入硬/软门禁。这是为了避免「每期都因同一批历史缓存值被卡死」的伪阻断。

软门禁（不足则 `COMPLETE_WITH_EXCEPTIONS`）：来源覆盖完整、无待核结论性字段、PDF 分页与上期基准一致、PDF 无空白/中文缺字页、台账指标显式 `validation=PASS`、Excel 公式缓存已在最终 Office 重算、发布级 Office 原生 PDF 已装配。

默认 LibreOffice 链路因此通常落在 `COMPLETE_WITH_EXCEPTIONS`：这是如实标注，不是失败。要把状态推进到 `COMPLETE`，需补齐官方源覆盖、提供 Office 原生导出的三份 PDF 与导出清单、并在 Microsoft Excel 完成公式重算。

`COMPLETE_WITH_EXCEPTIONS` 会出件，但《本期编制与待核说明》首页强制标注「待复核稿，不得命名为最终发布版、不得对外发送」；`artifact-manifest.json` 同时记录例外清单。

## 8. 原子发布

全部文件先在成品根目录同文件系统的 `.publishing-<pid>-<ts>/` 拼齐，做角色/哈希/临时文件预检，写 `artifact-manifest.json` 并二次校验，再 `os.replace` 一次性改名为正式目录；发布后逐文件复核哈希。任何失败只留 staging，不产生半份成果包。正式目录已存在时自动改用 `_r2/_r3` 修订目录，不覆盖已完成版本。

## 9. 电脑操控边界

用户已授权 computer use，但**默认链路一次调用即到终态，不产生任何待接管请求**。增强通道是事后可选项，只用于：

- 官网改版、只能经网页下载、或已有登录态才能访问的补充采集；
- Excel 在 Microsoft Excel 中「公式 → 立即计算」并核对缓存错误；
- Word/PowerPoint 以「最适合打印」导出发布级 PDF，消除版式软例外。

增强产物通过 `--pdf-source-dir`、`office-pdf/office-export-manifest.json` 和写回 `sources/` 的证据回流，再以同一日期重跑主入口；幂等状态只刷新受影响阶段。登录失效、许可证过期、验证码、付费权限或 AX 控件变化时，增强任务停止并记录 `NEEDS_SESSION` / `UI_SCHEMA_CHANGED`，**保留默认成果包**，不得猜坐标无限重试、不得绕过验证、不得把屏幕文字直接当结构化事实。详细规程见 [unattended-computer-use.md](unattended-computer-use.md)。

## 10. 已知限制（不得对外声称已解决）

1. 结论性字段的全文判定、跨年口径、券商角色与金额口径仍需人工或第二条解析链复核，脚本只列 `pending-review.json`。
2. 深交所并购重组专用目录号、北交所与证监会注册源的结构化采集仍是 best-effort，可能记 `PARTIAL/LIMITED`。
3. LibreOffice 导出的 PDF 分页与 Microsoft Word 发布版不一致；发布级版式需固定 Office 环境。
4. Excel 公式缓存值沿用上期，未在脚本内重算。
5. PPT 原生图表数据、内嵌工作簿与案例解析文字尚未由指标快照全量重建，本期只做期次/口径锚点替换与结构保真。
6. 首次部署需确认渲染器、字体配置与官方接口可用性；`--offline` 仅用于回归测试，不能作为正式出件依据。

## 11. 验收与回归

```bash
# 单元自测（期次推导、替换引擎、状态幂等、门禁）
python scripts/test_auto_run.py

# 离线回放：复用已固化的 sources，验证 16 原生 + 3 PDF + 说明 + 原子发布
python scripts/auto_run.py --as-of 2026-11-07 --offline \
  --out-root <测试输出目录> --work-root <测试运行目录>

# 同一命令重跑：所有阶段应 skipped，不产生重复记录/重复发布
python scripts/auto_run.py --as-of 2026-11-07 --offline \
  --out-root <测试输出目录> --work-root <测试运行目录>

# 正式运行（真实官方源）
python scripts/auto_run.py --as-of 2026-11-07
```

验收分两层，**结构通过不等于内容通过**：
- **结构层**：已知答案回放能重建上期实际存在的全部原生角色与 PDF；文件名期次与统计区间正确；上期原件哈希不变；重复运行不重复插入交易；来源缺失时自动标例外或阻断；PDF 逐页 PNG 可查。
- **内容层**：`test_known_answer_43.py` 与 `test_known_answer_42.py` 的事实检查全通过；文档体量检查可以是有说明的 WARN，但不能把模板不等价冒称为发布版。还要求 0/4/7/10号底表当期事件代码集一致、核心金额可复算及连续三期无人工修改。当前第43期核心事实通过（`PASS 27/WARN 3/FAIL 0`）、模板保真仍有 WARN；**第42期为 `PASS 22/WARN 2/FAIL 24`，未通过**，失败集中在明细台账口径与叙述层，详见 [blind-replay-42.md](blind-replay-42.md)。任意一期门禁有 FAIL 时，产出不得称可发布稿。

### 2026-11-08 本机回归结果（供复用，不替代逐期内容复核）

| 场景 | 结果 |
|---|---|
| 基础单元/入口自测 | `test_auto_run.py` 15项、`test_extract_filings.py` 10项、`test_weekly_ledger.py` 12项、`test_workbook_mutations.py` 4项、`test_materialize_package.py` 4项，全通过 |
| 第43期真实源采集 | 上交所动态109条；深交所官方并购重组接口145条；北交所项目页3条；巨潮38组关键词去重359条，按角色优先下载359/359份 |
| 第43期全链回放 | 9阶段全OK；16原生+3PDF+manifest+说明，共21文件；`COMPLETE_WITH_EXCEPTIONS`，硬门禁0失败 |
| 第43期事实门禁 | `PASS 27 / WARN 3 / FAIL 0`；0/4/7/10号Excel当期事件代码集合与现场稿一致 |
| 来源覆盖 | SSE_MOMENTUM / SZSE_MOMENTUM / CNINFO_FULLTEXT / BSE_MOMENTUM = `COMPLETE`；SSE_COMMITTEE / CSRC_REGISTRATION = `PARTIAL` |
| 发布控制 | 原子发布21文件；重复运行按阶段指纹跳过；`--publish-policy complete-only` 不出软例外稿 |
| 渲染器失效与恢复 | 失败写 `blocked-report.json`，退出码2；增加临时用户配置和重试后，三份PDF及逐页PNG一次通过 |
| `--pdf-source-dir` | 仍需固定 Microsoft Word/PowerPoint 导出清单才能消除“发布级原生PDF”软例外 |

### 2026-11-08 第43期盲测差异（已执行）

隔离运行路径 `_independent_43_release2_20261108/`。源数据先独立采集，现场第43期仅在独立成品发布后用于对照。完整分析见 [blind-replay-43.md](blind-replay-43.md)。

| 维度 | 独立稿 | 现场第43期 | 判定 |
|---|---|---|---|
| 文件角色 | 16原生+3PDF+说明 | 16原生+2PDF | 骨架通过；PDF多出PPT预览 |
| 核心事件 | 预案2、现金草案1、受理1、过会3、注册1、控制权3/完成1 | 同左 | 一致 |
| 0/4/7/10号工作簿事件代码集 | 一致 | 基准 | 通过 |
| 在审数 | 工作簿勾稽48 | 正文49，基础表48 | 现场稿内部不一致；独立稿如实列48 |
| 12号协议转让 | 当期4行 | 当期0行 | 独立稿多识别4单，均有原文 |
| 11号专业机构合作 | 当期23行 | 当期7行 | 口径和覆盖仍需双向复核 |
| 周报/动态/PPT文本行 | 154 / 96 / 35 | 1620 / 141 / 187 | WARN：模板保真不足 |
| PDF页数 | 周报5、动态4、PPT8 | 周报32、动态5、无PPT PDF | WARN：版式不等价 |

### 2026-11-08 第42期严格盲测（已执行，未通过）

隔离运行路径 `_blind42_20261108_1450/`；只用第41期完整模板 + 官方公告，生成终态并冻结哈希后才拆封现场第42期。完整分析见 [blind-replay-42.md](blind-replay-42.md)。

| 维度 | 独立稿 | 现场第42期 | 判定 |
|---|---|---|---|
| 文件角色 | 16原生+3PDF+说明 | 15原生+1PDF | 骨架通过；多出动态Word及2份PDF（模板继承） |
| 本期事件口径 | 首次披露2、草案1、过会1、受理0、注册0、终止0 | 同左 | **通过** |
| 年度累计 | 74 / 36 / 32 / 66 / 37 | 同左 | **通过**，第43期 `annual==weekly` 缺陷已修复 |
| 明细台账勾稽 | 10项不符（专业合作11/7、协议转让5/3等） | 基准 | 未通过：业务口径闸门缺失 |
| 叙述层 | Word 115行、PDF 4页；18项口径命中5项 | Word 1597行、PDF 31页 | 未通过：从零重写 |
| 汇总区公式缓存 | 沪主板 24/11/12 | 25/12/13 | 未通过：未重算 |
| 已知答案门禁 | `PASS 22 / WARN 2 / FAIL 24` | — | `KNOWN_ANSWER_FAIL` |

**运行环境自检（v4.1.1 新增）**：首次在线运行曾在 WorkBuddy Python 3.13 下 `BLOCKED`——`build_reports.ensure_libs()` 把 3.12 的 site-packages 注入 3.13 进程，`PIL._imaging` C 扩展 ABI 不兼容。现在 `ensure_libs()` 只接受同 minor 版本，`auto_run.py` 入口 `ensure_runtime()` 会在主链依赖（`openpyxl/docx/pptx/pypdf/lxml`）不齐时自动切换到同机合格解释器并重新执行自身；找不到则清晰 `BLOCKED`。可用 `MA_WEEKLY_PYTHON` 显式指定，**严禁跨 Python minor 版本混用 site-packages**。

**仍未完成的验收项**：Word/PPT模板保真重建、公告业务口径闸门、台账字段规范化、Excel全字段回填和分榜公式重算、已知答案门禁接入 `VALIDATE` 主链、连续三期无人工修改、Office原生PDF发布链。当前默认产出是**事实已核验范围明确的内部待复核稿**，不得直接对外发送。

## 12. 定时执行（可选）

稳定后可设每周六 22:00 运行，仅生成成果包；邮件、微信或群发送需单独授权与配置，不包含在「生成成品」内。
