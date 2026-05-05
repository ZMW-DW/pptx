# Contributing Guide

感谢对 thesis-defense-pptx-skill 的兴趣！这是一个面向毕业论文答辩 PPT
场景的 AI workflow skill，**不是通用 PPT 生成框架**。下面说明哪些方向欢迎
PR、本地开发环境、提交前的自检要求和安全相关注意事项。

## 项目定位

- ✅ 欢迎：让 inspect / clone / fill / quality-gate 流程在更多模板下更稳的
  修复；新增轻量级 helper、preset 或 example；补本地 / Codex / Claude Code
  / Cursor 等 agent 的接入说明；修脚本 bug 与跨平台问题；补文档与质量门
  规则。
- ❌ 不太欢迎：把脚本改造成"内置一套通用 AI PPT 设计模板"的方向。这个仓库
  的核心定位是 **保留用户提供的 PowerPoint 模板**，不是生成新风格。

如果你的需求是"我希望 AI 直接根据论文生成一套全新风格的 PPT"，建议参考
[`ppt-master`](https://github.com/hugohe3/ppt-master) 这种生成式工具，它和
本仓库是互补关系。

## 提 Issue

提 issue 时请尽量包含：

1. 操作系统 + Python 版本 + Microsoft PowerPoint 版本（如果走 COM 步骤）。
2. 复现命令的完整调用方式。
3. 期望行为 vs 实际行为。
4. 如果是模板适配问题，附 `dump_pptx_content.py` 输出的 `dump.md`（删除
   敏感学校信息后），不要直接附整份 `.pptx` 模板。

如果是安全相关问题（例如脚本调用 PowerPoint COM 的安全使用方式），请走私下
渠道（GitHub Security Advisory），不要公开 issue。

## 本地开发环境

最小依赖：

```bash
pip install python-pptx Pillow PyMuPDF pypdf
```

如果要走 PowerPoint COM 阶段（导出 PNG / 检查文字溢出 / 复制母版页），需要：

- Windows
- 已安装的 Microsoft PowerPoint
- `pip install pywin32`

跑示例验证你的改动没破坏主流程：

```bash
python examples/minimal_markdown/run_example.py
```

非 Windows 环境下示例会自动跳过 COM 步骤，但 `python-pptx` 部分、dump、scan
和 contact sheet 必须仍能跑通。

## 提 PR 前的自检

提 PR 前请确认：

- [ ] 修改的 Python 脚本能通过 `python -m py_compile <file>`。
- [ ] `examples/minimal_markdown/run_example.py` 能跑完，并且
      `scan.json["bad_hits"]` 为空（没有意外引入旧模板词）。
- [ ] 没有把 `*.pptx` / `*.pdf` / `rendered_slides/` / `*.png` 等运行
      产物 commit 进 git；`.gitignore` 已覆盖示例目录的常见产物。
- [ ] 不要在 PR 里附带含个人姓名、学号、学校签章、校徽位图等隐私 / 版权敏感
      内容的真实模板和论文。
- [ ] commit message 写明 **why**，而不只是 **what**。

## Coding Style

- Python 用 4 空格缩进，类型注解视情况而定。
- 脚本对外的 CLI 通过 `argparse`，不要引入新的 CLI 框架。
- Helper（`pptx_template_tools.py`）保持函数式 + 不可变输入；不要在内部
  缓存 slide / shape 引用。
- 不要在脚本里发起网络请求。
- PowerShell 脚本统一加 UTF-8 codepage 兜底，Python 脚本统一在入口加
  `sys.stdout.reconfigure(encoding="utf-8", errors="replace")`，避免
  cp936 控制台破坏中文 / 特殊字符 (`−`, `Δ`, `✓`)。

## 安全相关贡献

请优先关注：

- PowerPoint COM 自动化场景下不要让脚本默默改用户已经打开的 deck；
- 所有路径参数都应基于用户显式输入，不应猜测 / glob 用户磁盘；
- 不要在脚本里调用网络 API；
- 不要把第三方校徽 / 模板 / 字体等版权资产 commit 进仓库。

## 版权与协议

本项目使用 [Apache License 2.0](LICENSE)。提交 PR 即表示你同意你的贡献以
同样的协议被接受、分发和再许可。
