# ma-weekly-report

面向 A 股并购重组周报编制的 Codex skill，版本 **v4.5.16**。

它用于按上期成果和原始公告更新周报底表、控制权变更、协议转让、IPO 终止、引战合作等数据，并生成 Word/PDF/PPT 成果包。详细工作流、口径和限制以 [`SKILL.md`](SKILL.md) 为准。

## 安装

```bash
mkdir -p ~/.codex/skills
git clone <本仓库 URL> ~/.codex/skills/ma-weekly-report
```

也可以在 Codex 中直接打开仓库目录使用。

## 使用

典型入口：

```bash
python scripts/auto_run.py --as-of YYYY-MM-DD
```

脚本支持按截止日推导期次、采集公开来源、生成和校验成果。正式使用前应阅读 [`SKILL.md`](SKILL.md) 及 `references/` 中的来源、工作簿和交付规则。

## 环境

- Python 3.12 左右，并具备 `openpyxl`、`python-docx`、`python-pptx`、`pypdf`、`lxml` 等依赖。
- 正式 Office 导出与版式验收需要 Microsoft Office；LibreOffice/Poppler 可用于预览和渲染检查。
- 部分外部检索可能需要用户在环境变量中自行配置第三方数据服务凭证。仓库不包含任何密钥。

## 公开发布说明

- 仅发布 `ma-weekly-report/` 的正式文件，不包含备份、缓存、周报原始 Excel、录制、项目记忆或本机日志。
- 已移除本机绝对路径和微信文件路径；`pdftoppm` 改为从 `PATH` 查找。
- 本仓库是个人工作流与研发材料，不代表任何机构的官方口径或法律意见，生成结果均应按 `SKILL.md` 的待核与质量门禁复核。

## 许可证

本仓库未添加 `LICENSE` 文件，公开可见不构成开源许可或再分发授权。如需使用许可，请联系仓库所有者。
