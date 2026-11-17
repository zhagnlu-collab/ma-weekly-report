# ma-weekly-report

面向 A 股并购重组周报编制的 Codex skill，版本 **v4.5.16**。

该 skill 以“输入一个统计截止日，自动生成一套可复核的周报成果”为目标，覆盖期次推导、公告与交易所数据采集、13 个 Excel 底表更新、指标复核、Word/PDF/PPT 生成、逐页渲染检查和质量门禁。详细工作流、业务口径与限制以 [`SKILL.md`](SKILL.md) 为准。

> 本项目是个人工作流与研发成果，不代表任何机构的官方口径或法律意见。自动生成的成果均应按 `SKILL.md` 中的质量门禁复核，不得未经复核直接对外发布。

## 适用场景

- 编制或更新 A 股并购重组周报。
- 滚动维护重组底表、控制权变更、协议转让、IPO 终止、引战合作等 Excel。
- 根据上期成果和本期公告生成 Word、PDF、PPT 完整成果包。
- 复核已填数据与公告原文、交易所项目动态、证监会注册信息是否一致。
- 在既有模板上保留格式、公式、图表和跨表勾稽关系。

不适用于泛行业周报、单一并购案例研究或纯申报文件复核。

## 主要能力

- 按截止日自动推导期号、统计区间和上期模板。
- 采集上交所、深交所、北交所、巨潮资讯、证监会等公开来源。
- 建立公告检索台账、全文核实台账和事件状态记录。
- 更新 **13 个 Excel 底表**，执行公式重算、跨表勾稽和标绿范围校验。
- 生成 **WORD/PDF 周报、WORD/PDF 市场动态、PPT/PDF 市场周报**。
- 通过质量门禁区分 `COMPLETE`、`COMPLETE_WITH_EXCEPTIONS` 和 `BLOCKED`。
- 保留断点与运行状态，重复运行同一输入时可跳过已完成阶段。

## 交付物

一次完整运行通常生成：

- 13 个原生 Excel 工作簿；
- 周报 Word、PDF；
- 市场动态 Word、PDF；
- 市场周报 PPT、PDF；
- 检索台账、全文核实台账、指标快照和运行状态；
- `validation-report.json` 等质量门禁文件；
- 《本期编制与待核说明》（存在软门禁例外时）。

实际文件数量及目录以当期模板和运行结果为准。`COMPLETE_WITH_EXCEPTIONS` 表示结构已生成但仍有待核事项，不等同于内容已全部核实；`BLOCKED` 不视为可交付成果。

## MCP 与数据能力

本 skill 的完整数据处理链需要金融数据 MCP。**不是所有 MCP 都必须安装**，至少按下表配置一套核心数据源。Codex 中可分别通过 `wind-mcp-skill`、`ifind-finance-data` 和 `jianwei-agent-search` 接入 Wind、同花顺 iFinD 与见微能力；这些外部组件不包含在本仓库中。

| 优先级 | MCP / 数据组件 | 必要性 | 主要用途 |
|---|---|---|---|
| 1 | **Wind MCP** | 核心，二选一 | 行情、收盘价、市值、Wind/中上协/证监会行业、注册资本、主营业务、财务与公司资料 |
| 1 | **同花顺 iFinD MCP** | 核心，二选一 | 行情、市值、行业、注册资本、主营业务、公司资料与公告资讯 |
| 2 | **见微公告检索** | 推荐，可选 | 公告全文、权益变动报告书、股权结构、注册资本和主营业务补充核验 |
| 3 | **交易所项目查询 skills** | 可选 | 复用上交所、深交所、北交所项目动态的公开查询能力；没有时可通过官网或通用网页访问完成 |
| 4 | **Computer Use / 浏览器控制** | 可选增强 | 网页改版、登录态受限页面补查，以及 Microsoft Office 原生 PDF 导出 |

### 1. Wind MCP

适合需要 Wind 分类体系、证券基本资料、行情和历史数据的场景。建议配置以下能力：

- 证券搜索与基本资料；
- 股票行情、收盘价、市值；
- Wind 一级至四级行业；
- 中上协、证监会行业分类；
- 注册资本、主营业务、财务和股东信息。

Wind MCP 的账号、授权、额度和调用方式以 [Wind 官方 MCP 服务](https://aifinmarket.wind.com.cn)说明为准。使用前先在当前 Codex 环境中确认 MCP 已连接，并实际调用一次证券查询验证权限。

### 2. 同花顺 iFinD MCP

可作为 Wind MCP 的替代或交叉验证渠道，主要使用股票资料、行情、行业和公告资讯能力。常用字段包括：

- `get_stock_info`：注册资本、主营业务等；
- `get_stock_summary`：公司概况；
- 行情、市值和行业分类数据；
- 公司公告与新闻检索。

首次使用通常需要在 [iFinD MCP 官网](https://mcp.51ifind.com/#/profile)的“个人中心－密钥”页面获取 MCP 密钥，并按对应 skill 的配置要求保存。密钥属于个人凭证，不应提交到本仓库，也不应写入周报或日志。

### 3. 见微公告检索（可选）

当 Wind / iFinD 没有返回所需字段时，可使用见微数据检索公告全文、简式权益变动报告书、股权结构、注册资本和主营业务。该能力在常见 Codex 环境中以 `jianwei-agent-search` skill/API 的形式接入，需要用户自己的 [见微数据](https://www.jianweidata.com) ApiKey。

如果只使用公开交易所网站和巨潮公告，也可以不安装见微，但部分字段可能需要更多人工核验，并会在质量门禁中标记为待核或覆盖不足。

### 不需要额外 MCP 的公开来源

以下来源由 skill 自带脚本直接访问或通过浏览器核验，不要求安装专用 MCP：

- [上交所并购重组项目动态及会议公告](https://www.sse.com.cn/listing/renewal/ma/)；
- [深交所重组项目动态](https://www.szse.cn/listing/projectdynamic/reproperty/index.html)；
- [北交所重大资产重组项目动态](https://www.bse.cn/audit/t_project_news_r.html)；
- [巨潮资讯公告](https://www.cninfo.com.cn/)；
- [证监会并购重组注册信息](https://www.csrc.gov.cn/)。

搜索引擎、微信公众号和新闻页面只能用于发现线索，关键金额、比例、主体、控制权及日期必须回到公告或官方页面核验。

## 安装

### 1. 安装 skill

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/zhagnlu-collab/ma-weekly-report.git \
  ~/.codex/skills/ma-weekly-report
```

也可以在 Codex 中直接打开本仓库目录使用。

### 2. 安装 Python 依赖

建议使用 Python 3.12 左右版本：

```bash
python3 -m pip install \
  openpyxl python-docx python-pptx pypdf pdfplumber PyMuPDF lxml
```

### 3. 配置文档与 PDF 工具

- 正式 Office 导出优先使用 Microsoft Word、Excel、PowerPoint。
- LibreOffice 可用于格式转换和预览，不作为正式发布版的唯一判定工具。
- PDF 逐页渲染检查需要 Poppler 的 `pdftoppm`，并应确保其位于 `PATH` 中。

### 4. 配置 MCP

在 Codex 的 MCP / 连接器设置中完成以下至少一套配置：

1. Wind MCP；或
2. 同花顺 iFinD MCP。

如需提高公告检索覆盖率，再配置见微公告检索。配置完成后，建议先运行一次只读股票查询，确认认证、权限和额度正常。

## 快速开始

先做计划推导，不执行采集和写文件：

```bash
cd ~/.codex/skills/ma-weekly-report
python3 scripts/auto_run.py --as-of YYYY-MM-DD --plan-only
```

确认期次、区间、模板和输出目录后执行完整流程：

```bash
python3 scripts/auto_run.py --as-of YYYY-MM-DD
```

如果是离线结构回放或已有官方 PDF，可使用：

```bash
python3 scripts/auto_run.py --as-of YYYY-MM-DD --offline
python3 scripts/auto_run.py --as-of YYYY-MM-DD --pdf-source-dir /path/to/office-pdf
```

运行目录和输出目录默认位于 skill 所在目录及其父目录附近，可通过 `--root`、`--work-root`、`--out-root` 或 `--config` 调整。完整参数说明见：

```bash
python3 scripts/auto_run.py --help
```

## 运行边界

- 生成结果默认是待复核稿，不代表已完成外部核验。
- 任何公告、法规、监管结论和统计数字都应回到原文核验。
- 股票行情、市值和行业字段必须记录查询日期、复权、币种和单位。
- MCP 结果与公告冲突时应并列保留来源，不得静默选择。
- 不得绕过验证码、权限或付费限制。
- 不得上传未经授权的内部材料、客户信息或个人信息。

## 目录结构

```text
.
├── SKILL.md          # 主技能说明与业务口径
├── README.md         # 安装、MCP 与使用说明
├── agents/           # Codex 界面入口配置
├── config/           # 无人值守配置示例
├── references/       # 检索、工作簿、质量门禁和演示规则
└── scripts/          # 采集、生成、复核与质量检查脚本
```

## 公开发布说明

- 仅发布 `ma-weekly-report/` 的正式文件，不包含备份、缓存、周报原始 Excel、录制、项目记忆或本机日志。
- 已移除本机绝对路径和微信文件路径；`pdftoppm` 改为从 `PATH` 查找。
- 本仓库不包含任何 MCP 密钥、API Key、账号密码或客户内部材料。

## 参与贡献

欢迎通过 Issue、Discussion 和 Pull Request 改进本 skill。外部贡献者请在 Fork 中修改并发起 PR，维护者审核通过后合并；`main` 分支不接受未经审核的直接推送。

- 贡献流程与本地验证要求见 [`CONTRIBUTING.md`](CONTRIBUTING.md)。
- 第一次使用 GitHub 或需要完整的 Fork、分支、PR 教程，请阅读 [`docs/github-contribution-tutorial.md`](docs/github-contribution-tutorial.md)。
- 提交前请运行 `python3 scripts/validate_skill.py`。
- 安全问题、密钥泄露或敏感数据风险请按 [`SECURITY.md`](SECURITY.md) 私下报告。
- 参与讨论即表示同意遵守 [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md)。
- 不得在 Issue、PR、日志或测试文件中提交 MCP 密钥、客户信息、个人信息或未公开交易信息。

## 许可证

本项目采用 [MIT License](LICENSE)。在保留版权声明和许可证声明的前提下，可以自由使用、复制、修改、合并、发布、分发、再许可和销售本项目的副本。
