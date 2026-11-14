# 可选增强：电脑操控补查与发布级版式

**本文件描述的不是默认链路。** 默认「日期进、成品出」由 `auto_run.py` 纯脚本完成，不依赖任何 GUI 应用、登录态或人工接管：

```bash
python scripts/auto_run.py --as-of 2026-11-07
```

程序从工作区最近完整期次自动推导起止日、期号、历史年度库编号、模板、输出位置和验收阈值；工作先进入运行目录的 `content/` 与 `building/`，全部门禁通过后原子发布，一次调用即结束，不会停在 `NEEDS_COMPUTER_USE` 等待下一次触发。默认链路如实把「LibreOffice 分页不等价于 Microsoft Office 发布版」「公式缓存未在 Office 重算」记为软例外，产出 `COMPLETE_WITH_EXCEPTIONS` 的完整成果包并随附《编制与待核说明》。

只有在下列三种情形才启用本文件的电脑操控增强，且它始终是**事后可选**，不阻断默认出件：

1. 官方接口改版、返回 `LIMITED/PARTIAL`，或入口只能靠已有登录态访问，需要 Chrome 补查；
2. 需要 Microsoft Excel「公式→立即计算」重算并复核全部缓存值；
3. 需要 Microsoft Word/PowerPoint「最适合打印」导出的发布级 PDF，以消除版式软例外、把状态推进到 `COMPLETE`。

增强结果通过 `--pdf-source-dir`、`office-pdf/office-export-manifest.json` 与写回 `sources/` 的采集证据回流到主链路，再以**同一日期重跑** `auto_run.py`；幂等状态会跳过已完成阶段，只刷新受影响阶段。禁止把「接口零返回」解释为无案例，也禁止绕过验证码或付费权限。

## 1. 增强通道自举（仅启用增强时需要）

首次启用增强前跑一次 `bootstrap`：Chrome 打开六个官方入口并确认持久登录态；Microsoft Excel 重算一个带公式样本；Word 和 PowerPoint 各导出一个固定测试文档；Wind 完成一次股票、行情、行业读取。记录应用版本、AX 控件指纹、下载目录和黄金 PDF 元数据。默认链路无需这一步。

登录态失效、许可证过期或应用更新导致控件消失时，增强任务写 `NEEDS_SESSION` 或 `UI_SCHEMA_CHANGED` 并**放弃增强、保留默认成果包**；不得猜坐标无限重试，也不能自动绕过验证码或付费权限。

## 2. 阶段状态机

默认主链路（`auto_run.py` 实际执行）：

```text
PLAN → ACQUIRE → EXTRACT → METRICS → BUILD → RENDER → VALIDATE → NOTES → PUBLISH
```

状态：`PENDING / RUNNING / OK / SKIPPED / FAILED / COMPLETE / COMPLETE_WITH_EXCEPTIONS / BLOCKED`。增强通道另用自己的 `NEEDS_SESSION / UI_SCHEMA_CHANGED`，只影响增强任务，不改写主链路状态。

每个阶段以输入、配置和程序版本哈希做幂等键，写入 `run-state.json`。相同输入已通过则跳过；失败只重跑失败阶段。重复执行同一日期形成修订序号，不覆盖已完成成果。

## 2A. 已实现的可执行组件

```bash
# 默认链路：一次调用直接出件，不产生任何待接管请求
python scripts/auto_run.py --as-of YYYY-MM-DD

# 可选增强：先由电脑操控导出三份发布级 PDF 并登记，再以同一日期重跑
python scripts/record_office_export.py --pdf-dir RUN/office-pdf --file-role weekly_docx --source CONTENT/5、...docx --pdf RUN/office-pdf/5、...pdf --native-pages 32
python scripts/auto_run.py --as-of YYYY-MM-DD --pdf-source-dir RUN/office-pdf

# 诊断/辅助工具（非主链路依赖）
python scripts/apply_workbook_mutations.py 0_working --mutations mutations.json --out-audit mutation-audit.json
python scripts/weekly_ledger.py ledger.json --out metrics.json
python scripts/materialize_package.py 已核实内容目录 --out 最终装配目录 --pdf-source-dir 本期Office导出目录
python scripts/validate_package.py 最终成果目录 --period YYYYMMDD至YYYYMMDD --expected expected.json --out validation.json
python scripts/test_auto_run.py
```

`record_office_export.py` 每登记一份 PDF 就更新 `RUN/office-pdf/office-export-manifest.json`，三份齐全后状态变 `COMPLETE`；以同一日期加 `--pdf-source-dir` 重跑时，`render_verify.py` 装配 Office 原生 PDF，`quality_gate.py` 把「Microsoft Office 发布 PDF 清单」作为硬门禁核验。`run_weekly.py` 保留用于期次推导和阶段状态单元验证。

`materialize_package.py` 只接受精确匹配的三份 Office PDF 及完整导出清单；LibreOffice 只能通过 `--preview-soffice` 写入 `preview/`，返回 `PREVIEW_ONLY`，不能进入该工具的正式装配。`validate_package.py` 同时检查 Excel 公式缓存、缓存错误、文件期间、跨文件文本和三份 PDF，可作独立复核。

## 2B. 增强执行闭环（可选，默认不执行）

1. 先完整跑一次 `auto_run.py --as-of YYYY-MM-DD`，得到结构完整的默认成果包与 `source-manifest.json` 覆盖矩阵。
2. 若覆盖矩阵含 `LIMITED/PARTIAL/NOT_RUN`，用 Chrome 按缺口补查官方入口，把原始响应、公告 PDF、页数/记录数和覆盖证据写回 `sources/` 与 `source-manifest.json`；再以同一日期重跑主入口刷新事件台账。不得以接口零返回宣告无案例。
3. 若需消除内容层例外，用 `apply_workbook_mutations.py` 等辅助工具把已确认事件写入工作副本；Word 新增/修改使用 HTLH 修订；不得保留与本期台账矛盾的旧数字、旧案例、旧状态或无证据推断。
4. 按第 5 节逐个打开 13 个 Excel，使用「公式→立即计算」，保存并关闭。重算后公式不得无缓存值或保留 `#REF!/#VALUE!/#DIV/0!/#NAME?/#N/A`；确无业务的零分母用精确单元格修复，不整表豁免。
5. 按第 6、7 节打开周报 Word、动态 Word 和 PPT，先读状态栏原生页数/幻灯片数，再导出「最适合打印」PDF；两份 Word 各登记一次，PPT 登记一次。
6. 如 PDF 页数与原生页数不符，删除该次 PDF 并重导出；不得用 LibreOffice 结果替代。三份登记齐全后用 `--pdf-source-dir` 重跑主入口。
7. 只有 `validation-report.json` 为 `COMPLETE`，或明确允许软例外且输出状态为 `COMPLETE_WITH_EXCEPTIONS`，才可发布。`BLOCKED` 只保留诊断材料。

## 3. Chrome电脑操控

已验证入口：

- 上交所并购审核动态：https://www.sse.com.cn/listing/renewal/ma/
- 上交所会议公告与结果：https://www.sse.com.cn/listing/announcement/commiteema/
- 深交所并购审核动态：https://www.szse.cn/listing/projectdynamic/reproperty/index.html
- 北交所并购重组：https://www.bse.cn/audit/t_project_news_r.html。旧 `/projectDynamic/ma.html` 已失效；页面有 Cookie/JS 挑战时通过 Chrome 无障碍树读取全表，并保存为 `sources/bse_momentum.json`。
- 巨潮资讯：https://www.cninfo.com.cn/
- 证监会：https://www.csrc.gov.cn/ 后进入上市公司监管、并购重组注册。

执行要求：稳定官方接口优先；接口返回 `LIMITED/PARTIAL`、页面结构变化或需读动态表格时，使用Chrome原生应用的电脑操控补查。导航前点击地址栏，Command+A全选后粘贴URL。优先读无障碍树，表格、状态、总数和分页控件直接读取；AX缺失才截图并用OCR辅助，关键事实仍回公告PDF。

每次点击筛选、翻页、详情或下载后重读当前状态，不复用旧控件索引。对每类查询记录关键词、日期、板块筛选、结果数、总页数、实际读取页数、去重前后数量及排除理由。

下载闭环：定位浏览器下载项和新文件；检查文件头、MIME、SHA-256；同哈希复用，同名不同哈希保留两版并核定正式版本；保存原URL、公告标题、下载时间和页面位置。标题和摘要只产生候选，金额、比例、主体、控制权、日期等关键字段必须读全文。

## 4. Wind补充

Wind只补公告外字段：证券基础信息、行情、市值、Wind行业、财务数据。通过AX操作证券查询、字段选择、导出，再用CSV/Excel字段名映射进模板，不从屏幕文字直接拼数据。记录Wind版本、查询日期、复权、币种、单位和导出文件哈希；终端未登录则为 `NEEDS_SESSION`。

有稳定API时优先API。API与Wind冲突时保留双方和来源，不静默择一。

## 5. Excel重算

1. 每次只打开一个工作副本；超过一个文件时关闭前一文件。
2. 进入“公式”功能区，执行“立即计算”；不要使用会和“插入”命令冲突的快捷键。
3. 保存并关闭。打开下一文件前重新读取当前窗口状态，不复用上一条AX控件索引。
4. 以程序复核 `data_only=True` 的公式缓存和错误值。工作簿若有损坏提示，优先从原件副本恢复、重算并另存；不要接受删除公式或隐藏页的“修复”。

## 6. Word发布PDF

2026-11-08实机验证：

1. Word通过电脑操控打开周报工作副本；状态栏可读总页数。
2. 文件→另存为。
3. 文件格式选择PDF。
4. 选择“最适合打印”。
5. 输出到 `RUN/office-pdf/`，文件名必须与 DOCX 完全一致。
6. PDF页数必须等于Word状态栏总页数。
7. 比较正文文本并逐页渲染检查。
8. 用 `record_office_export.py` 登记源文件哈希、PDF哈希、Word页数、PDF页数、应用和导出方式。
9. 关闭文档，不保存任何未授权改动。

测试中的A股动态在Word状态栏为5页，经该方法导出PDF也是5页；LibreOffice同一文件曾导成4或8页，因此LibreOffice只可作为预览，不能作为发布版判定器。

## 7. PowerPoint发布PDF

2026-11-08实机验证：

1. PowerPoint打开演示文稿副本；状态栏可读幻灯片总数。
2. 检查原生图表、内嵌工作簿、目录和页码。
3. 另存为PDF。
4. 选择“最适合打印”。
5. 输出到同一 `RUN/office-pdf/`，文件名必须与 PPTX 完全一致。
6. PDF页数必须等于幻灯片数。
7. 用 `record_office_export.py` 登记源文件哈希、PDF哈希、幻灯片数、PDF页数、应用和导出方式。
8. 逐页渲染，检查文字、图表、图片和母版漂移。

测试PPT显示20张幻灯片，电脑操控导出PDF为20页。

## 8. 自动校验

- 13个Excel：ZIP完整、可打开、绿色周标记、隐藏页、榜单、汇总、公式缓存及错误值。
- 周度、年度、截止日状态及跨年桥接。
- 分板块合计、财顾项目数与家次、我司清单、委员会场次与审议项目。
- Word、动态、PPT的关键数量、金额、主体和期间一致。
- 三份PDF必须与Office导出清单的源/PDF哈希、原生页数/幻灯片数逐一匹配；用 `pdftoppm` 逐页生成PNG后检查文本和视觉。
- 文件名、期号、起止日、目录、页眉页脚、数据来源注。
- 源文件未变化、输出原子发布、重复运行幂等。

## 9. 失败与重试

- 网络超时：3次指数退避，仍失败标 `PARTIAL`。
- 页面总数未知：拆小日期或板块查询，不虚构总页数。
- AX结构改变：刷新并尝试一次视觉辅助；仍失败标 `UI_SCHEMA_CHANGED`。
- 登录/许可证失效：保存已完成部分，标 `NEEDS_SESSION`。
- Office崩溃：打开未锁定工作副本，最多重试2次。
- Excel损坏提示：使用原件和备份恢复，禁止通过删除公式或工作表规避。
- Office导出清单缺项、页数不符或源文件在导出后变化：删除该次PDF并重新导出，不能使用预览PDF顶替。
- PDF页数差异/版式回归超限：阻断发布，保留Office源文件和诊断图。
- 数据冲突：并列保留，阻断最终发布。
- 下载重复：按哈希去重；同名异哈希不得覆盖。

## 10. 正常生产的人工边界

默认生产完全由脚本完成，用户只提交截止日；人工只在首次配置、官网栏目变化、统计口径变化、应用升级和重大数据冲突修复时介入。启用增强通道时，电脑操控任务清单由同一自动化任务继续执行，仍不需要用户中途确认。修复后把脱敏页面结构、控件路径、黄金样本和版本纳入回归集，下一期恢复无人运行。

若要由定时任务触发，自动化提示词只写：`使用 $ma-weekly-report，以今天之前最近的周六为截止日，按技能的无人在环流程执行直到发布或明确阻断；不得中途停下等待确认。` 主入口一次调用即到 `COMPLETE`/`COMPLETE_WITH_EXCEPTIONS`/`BLOCKED` 终态，不会停在中途等待第二次触发；只有选择启用增强通道并遇到会话失效时，增强任务才写 `NEEDS_SESSION` 并保留默认成果包。

电脑操控授权用于本任务的信息浏览和文件生成。若执行中出现登录其他账户、上传文件、发送消息、修改云端数据、付款等外部行为，仍需按 Computer Use 现行确认规则处理。
