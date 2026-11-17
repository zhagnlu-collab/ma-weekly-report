# 参与贡献

感谢参与改进 `ma-weekly-report`。本仓库采用“公开提交、维护者审核、通过 Pull Request 合并”的协作方式。

## 可以贡献什么

- 修复脚本错误、兼容性问题或文档错误。
- 补充交易所、证监会、巨潮等公开来源的采集与核验方法。
- 改进 Wind MCP、同花顺 iFinD MCP、见微检索等数据源的适配。
- 增加可复现的测试、黄金样本和回归案例。
- 完善周报口径、质量门禁、错误提示和使用文档。
- 提出新的功能建议或数据源变化线索。

## 协作流程

1. 先在 Issues 或 Discussions 中搜索是否已有相同问题。
2. Fork 本仓库，并从 `main` 创建独立分支：
   - 新功能：`feat/short-description`
   - 缺陷修复：`fix/short-description`
   - 文档修改：`docs/short-description`
3. 完成修改并补充必要文档或测试。
4. 在本地运行验证。
5. 向 `main` 发起 Pull Request，填写变更范围、验证方式和风险。
6. 维护者审核、讨论并合并。

普通贡献者不能直接向主仓库 `main` 推送。维护者也将尽量通过独立分支和 Pull Request 完成变更。

## 本地开发

```bash
git clone https://github.com/<your-account>/ma-weekly-report.git
cd ma-weekly-report
git remote add upstream https://github.com/zhagnlu-collab/ma-weekly-report.git
git switch -c feat/short-description
```

安装常用依赖：

```bash
python3 -m pip install \
  openpyxl python-docx python-pptx pypdf pdfplumber PyMuPDF lxml
```

运行仓库结构校验：

```bash
python3 scripts/validate_skill.py
```

如修改了业务规则，还应在具备相应 MCP 和官方数据访问权限的环境中做最小可复现验证。

## MCP 与外部数据源

本仓库不包含 Wind MCP、同花顺 iFinD MCP、见微 API 或其他第三方数据服务凭证。

- 不得提交 MCP ApiKey、账号密码、Cookie、Token、私钥或 `.env` 文件。
- 不得在 Issue、PR、日志或测试夹具中写入客户信息、未公开交易信息或个人敏感信息。
- 涉及 MCP 的缺陷应说明产品名称、工具名、脱敏后的请求参数、错误码和可复现步骤。
- 外部来源变化应附官方页面链接、核验日期和必要的脱敏截图或文本摘录。

## Pull Request 要求

- 一个 PR 尽量只解决一个问题。
- 说明修改前后的行为、影响文件和数据口径。
- 业务规则变化必须同步更新 `SKILL.md` 或 `references/`。
- 用户可见变化必须同步更新 `README.md`。
- 不得降低质量门禁、删除既有核验步骤或把待核内容写成已核实事实。
- 不得加入本机绝对路径、缓存、日志、备份或未经授权的二进制文件。
- 提交前运行 `python3 scripts/validate_skill.py`。

## 版本与发布

维护者负责确认版本号、发布说明和主分支合并。贡献者可以在 PR 中建议版本级别，但不应自行发布、覆盖历史提交或强制推送 `main`。

## 安全与敏感问题

发现密钥泄露、未授权数据上传、路径穿越、命令注入或其他安全问题，请按 [SECURITY.md](SECURITY.md) 处理，不要先公开披露敏感细节。

## 许可证说明

本项目采用 [MIT License](LICENSE)。除非另有明确书面约定，提交 Pull Request 即表示同意该贡献按 MIT License 授权并随项目分发。
