# Thesis Defense PPTX Skill

一个用于生成本科/研究生毕业论文答辩 PPT 的 Codex/Agent Skill。它面向需要严格复用本地 PowerPoint 模板的场景，能够从本地论文 PDF/LaTeX 项目和指定 `.pptx` 模板出发，生成可编辑的正式答辩 PPTX，并执行逐页导出、版式检查和文字溢出检查。

[English README](README.md)

## 功能

- 从本地论文 PDF/LaTeX 项目中提取论文内容和候选图表。
- 尽量保留已有 PowerPoint 模板的封面、字体、字号、配色、导航、卡片样式和页面比例。
- 输出真实可编辑的 `.pptx` 文件，而不是图片型幻灯片。
- 使用 PowerPoint COM 导出逐页 PNG，便于视觉检查。
- 生成整套 PPT 的总览图，快速发现版式问题。
- 使用真实 PowerPoint 渲染结果检查文字框溢出风险。
- 检查旧模板文字、占位词、TODO 等残留内容。

## 仓库结构

```text
skills/
└── thesis-defense-pptx/
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── references/
    │   └── pptx_quality_gate.md
    └── scripts/
        ├── clone_template_deck.ps1
        ├── export_pptx_png.ps1
        ├── extract_thesis_context.py
        ├── inspect_pptx_overflow.ps1
        ├── make_contact_sheet.py
        ├── pptx_template_tools.py
        └── scan_pptx_text.py
```

## 环境要求

- 推荐在 Windows 上使用。
- 如果需要完整质量检查，需要安装 Microsoft PowerPoint。
- Python 3.10 或更高版本。
- Python 依赖：
  - `python-pptx`
  - `Pillow`
  - `PyMuPDF` 或 `pypdf`

安装依赖：

```powershell
python -m pip install python-pptx Pillow PyMuPDF pypdf
```

## 本地安装

将 Skill 复制到 Codex skills 目录：

```powershell
Copy-Item -Recurse -Force `
  .\skills\thesis-defense-pptx `
  "$env:USERPROFILE\.codex\skills\thesis-defense-pptx"
```

然后新开一个 Codex 会话，直接提出生成答辩 PPT 的需求，或明确说明：

```text
使用 thesis-defense-pptx skill。
```

## 推荐流程

1. 从论文 PDF/LaTeX/Word 项目中提取研究背景、方法、实验和结论。
2. 分析用户提供的 PowerPoint 模板。
3. 使用 PowerPoint COM 复制模板原生页面，生成 PPT 骨架。
4. 将论文内容转写为答辩口径，填入可编辑文本框、表格和图表。
5. 导出每页 PNG。
6. 生成总览图并进行视觉检查。
7. 检查文字溢出和旧模板词残留。
8. 根据检查结果反复修复，直到通过质量门槛。

## 常用命令

提取论文上下文：

```powershell
python .\skills\thesis-defense-pptx\scripts\extract_thesis_context.py `
  --input "D:\path\to\thesis-project" `
  --output "D:\path\to\thesis_context.md"
```

导出 PPTX 为逐页 PNG：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\thesis-defense-pptx\scripts\export_pptx_png.ps1 `
  -Pptx "D:\path\to\deck.pptx" `
  -OutDir "D:\path\to\visual_check" `
  -Width 1600 -Height 900
```

检查文字溢出：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\thesis-defense-pptx\scripts\inspect_pptx_overflow.ps1 `
  -Pptx "D:\path\to\deck.pptx" `
  -Tolerance 40
```

生成总览图：

```powershell
python .\skills\thesis-defense-pptx\scripts\make_contact_sheet.py `
  --input "D:\path\to\visual_check" `
  --output "D:\path\to\contact_sheet.png"
```

## 说明

本 Skill 不内置固定 PPT 模板。它的目标是尽量复用用户提供的模板，而不是强行套用通用设计风格。

## 许可证

MIT
