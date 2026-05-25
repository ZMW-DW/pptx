# CLAUDE.md — Thesis Defense PPTX 毕业答辩 PPT 生成

## 项目概述

从论文 PDF 生成毕业答辩 PPT。核心原则：**严格套用现成 PPT 模板，把论文内容转写成答辩口径**，保留模板的封面、校徽、配色、导航、字体字号，只替换内容。

## 核心命令

```bash
# 安装依赖
uv sync

# 从论文 PDF 提取文本和候选图表（默认提取 120 页，截断 60000 字符）
uv run skills/thesis-defense-pptx/scripts/extract_thesis_context.py \
  --input <论文PDF或项目目录> \
  --output data/thesis_context.md

# 如果论文超过 120 页或提取不完整，直接用 PyMuPDF 补读关键章节
uv run python3 -c "
import pymupdf as fitz
doc = fitz.open('论文路径')
for i in range(起始页, 结束页):
    text = doc[i].get_text('text')
    if text.strip():
        print(f'=== Page {i+1} ===')
        print(' '.join(text.split())[:2000])
        print()
doc.close()
"

# 导出 PPTX 每页的全部形状/文本/表格/图片清单（替换内容前必须先跑）
uv run skills/thesis-defense-pptx/scripts/dump_pptx_content.py \
  --pptx <input.pptx> \
  --output data/dump.md

# 导出指定页
uv run skills/thesis-defense-pptx/scripts/dump_pptx_content.py \
  --pptx <input.pptx> --output data/dump.md --slide 1,4,8,9

# 扫描 PPTX 文本和违禁词残留（每次填充后必跑）
uv run skills/thesis-defense-pptx/scripts/scan_pptx_text.py \
  --pptx <output.pptx> \
  --bad "中期检查,项目简介,TODO,占位,[placeholder],[论文标题],[学院名称],[导师姓名],[学生姓名],[在此填入],[模型描述],[模型架构],[实验结果表格],[实验结果数据],[研究总结与未来展望]"

# 生成总览图（需先导出 PNG，仅 Windows+PowerPoint）
uv run skills/thesis-defense-pptx/scripts/make_contact_sheet.py \
  --input <png_dir> --output data/contact_sheet.png
```

## 内置模板

哈尔滨商业大学 (HRBCU) 通用答辩模板：
- **路径:** `skills/thesis-defense-pptx/templates/HRBCU-defense-template.pptx`
- **规格:** 22 页, 16:9, 主色红 `#830604` (可设 `THESIS_PPTX_BRAND_RED` 环境变量覆盖)
- **结构:** 封面 → 目录 → 4 章节（背景/现状/研究内容/总结）→ 致谢
- **逐页说明:** `skills/thesis-defense-pptx/references/hrbcu_template.md`
- 模板已清理个人内容，全部正文替换为 `[placeholder]` 类标签

HRBCU 同校用户可直接使用内置模板；其他学校用户需自行提供 `.pptx` 模板。

## 工作流程

### 1. 收集输入
确认：论文 PDF 路径、学校（是否 HRBCU）、答辩层次（本/硕/博）、用户自备模板（如有）

### 2. 提取论文上下文

**Step 2a:** 运行 `extract_thesis_context.py` 提取全文到 `data/thesis_context.md`。

**Step 2b: 发现提取不完整时的补救方法。** 若论文公式/表格/图片较多，PyMuPDF 文本提取可能不完整（常见于理工科硕士论文的实验章节）。此时直接用上述 PyMuPDF 命令按页码范围提取关键内容：
- 摘要/结论：前 10 页 + 后 10 页
- 实验数据：从目录定位"实验"所在章节的页码范围
- 分章节提取，每次 `for i in range(start_page, end_page)` 循环

重点记录：
- 论文标题、作者、导师、学院、专业、学位
- 数据集名称、样本量、类别分布
- 模型对比数据：Acc/Prec/Recall/F1 具体数值
- 消融实验数据：各模块移除后的性能变化
- 核心创新点（一句话概括）、局限性、未来工作方向

### 3. 分析模板
- HRBCU 模板：读 `references/hrbcu_template.md` 获取 22 页 slide 清单
- 用户自供模板：先 `dump_pptx_content.py`，识别封面/目录/章节页/正文页/致谢页
- 记下：主色、字体、字号、各 slide 的 shape 精确名称和位置

### 4. 搭建 PPT 骨架
```bash
cp skills/thesis-defense-pptx/templates/HRBCU-defense-template.pptx work/output.pptx
```
- 规划 slide 顺序：硕士 15-17 页，本科 12-14 页
- 保留需要的 slide，删除多余的（用 `prs.slides._sldIdLst` 和 `prs.part.drop_rel`）
- 典型结构：封面→目录→研究背景→研究现状→研究内容总览→方法1→实验1→方法2→架构图→实验2→消融实验→总结→展望→致谢

### 5. 填充内容 — 关键步骤

**必须先 `dump_pptx_content.py` 再写替换脚本！** 模板字符串肉眼不可靠（em-dash、全半角引号、空格差异）。

```
dump → read dump → write fill script → run fill → scan → fix → repeat
```

**替换策略选择（重要）：**

| 场景 | 函数 | 原因 |
|------|------|------|
| 导航标签、章节名（短文本，唯一） | `replace_exact_text` | 精确匹配，不会误伤 |
| 正文段落（长文本） | `replace_partial_text` + `min_len` | 容忍模板标点差异 |
| **一个 slide 有多个同 keyword 的 shape** | 用 `find_shape_by_text` 逐个定位，`set_text_preserve_style` 逐个写入 | `replace_partial_text` 会同时命中所有匹配 shape，导致内容重复 |

**多 textbox 幻灯片处理陷阱：**
模板中 slide 10/11/12/13/16 等有多个 `[实验结果表格或分析]` 或 `[模型描述与创新点]` textbox。`replace_partial_text` 会一次性全部替换为相同文字，造成严重的内容重复。

正确做法：
```python
# 逐个找到 shape，按位置赋予不同内容
shapes = [s for s in iter_shapes(slide.shapes) if s.has_text_frame and "keyword" in s.text]
shapes.sort(key=lambda s: s.left)  # 按左位置排序
set_text_preserve_style(shapes[0], "左栏内容")
set_text_preserve_style(shapes[1], "右栏内容")
for s in shapes[2:]:
    set_text_preserve_style(s, "")  # 清空多余 shape
```

### 6. 内容转写原则
- 论文长段落 → 答辩 bullet points（`•` 开头，每条 1-2 行）
- 每页一个核心信息，不在同一页堆砌多个主题
- 优先用模板已有 textbox，不随意新增 shape
- 实验结果页保留关键数字：数据集名、Acc/F1 最优值、对比的 2-3 个 baseline
- 表格数据转为简短列表或卡片式对比
- 保留模板字体和字号层级；内容溢出时先缩短文字，再考虑缩小字号

### 7. 质量检查（每轮填充后必做）

```
scan_pptx_text.py → 检查 bad_hits 是否为空
dump_pptx_content.py → 目视检查关键 slide 是否有重复内容或空 textbox
```

在 macOS 上最后手动打开 `.pptx` 逐页检查视觉（字体、溢出、对齐）。

### 8. 交付
给出：输出路径、总页数、完成的检查项、需人工复查的内容。

## Python 工具库 (`pptx_template_tools.py`)

```python
from skills.thesis_defense_pptx.scripts.pptx_template_tools import *

# --- 保留原样式的文字替换 ---
set_text_preserve_style(shape, "新文字")
replace_exact_text(slide, {"旧文字": "新文字"})
replace_partial_text(slide, {"关键词": "新全文"}, min_len=20)

# --- 新增元素 ---
add_text(slide, "文字", x, y, w, h, size=18, bold=False, color=BLACK)
add_para(slide, ["行1", "行2"], x, y, w, h, size=14)
rect(slide, x, y, w, h, line=RED)        # 卡片框
tag(slide, "标签", x, y, w, h, fill=RED)  # 红底白字标签
add_pic_fit(slide, "img.png", x, y, w, h) # 图片自适应缩放

# --- 查找 ---
shape = find_shape_by_text(slide, "关键词")
pic = find_picture(slide, name_substr="架构")
table = find_table(slide)

# --- 替换图片/表格 ---
replace_picture(slide, old_pic, "new.png")
write_table(table, [["cell1", "cell2"], ["cell3", "cell4"]])

# --- 遍历所有 shape（含 group 内嵌套） ---
for shape in iter_shapes(slide.shapes):
    if shape.has_text_frame and "目标文字" in shape.text:
        ...
```

主色常量 `RED = RGBColor(0x83, 0x06, 0x04)` 可通过环境变量 `THESIS_PPTX_BRAND_RED` 覆盖。

## 删除 PPT 幻灯片

```python
# 从末尾向前删除（索引基于 0）
for idx in sorted([17, 16, 14, 12, 10], reverse=True):
    slide_id = prs.slides._sldIdLst[idx]
    rId = slide_id.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    if rId is None:
        rId = slide_id.rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[idx]
```

## 平台差异

| 功能 | macOS/Linux | Windows + PowerPoint |
|------|------------|---------------------|
| python-pptx 读/写/编辑 | ✓ | ✓ |
| extract_thesis_context | ✓ | ✓ |
| dump_pptx_content | ✓ | ✓ |
| scan_pptx_text | ✓ | ✓ |
| make_contact_sheet | ✓ | ✓ |
| COM 克隆模板 slide | ✗ | ✓ |
| COM 导出 PNG | ✗ | ✓ |
| COM 溢出检查 | ✗ | ✓ |

macOS 上跳过 COM 步骤，依赖 python-pptx + `dump_pptx_content` + `scan_pptx_text` + 手动打开文件检查。

## 常见问题和解决

| 问题 | 原因 | 解决 |
|------|------|------|
| PDF 提取不完整 | 论文长或公式多 | 用 PyMuPDF 直接按页码范围提取 |
| `replace_partial_text` 导致内容重复 | 多个 shape 匹配同一 keyword | 改用 `find_shape_by_text` 逐个定位写入 |
| `scan_pptx_text` 仍有 bad hits | 某些 textbox 不在 keyword 匹配范围 | 用 `find_shape_by_text` 搜索残留文本并逐个清除 |
| sidebar 和 body 文字重复 | 两处都是 `[模型描述与创新点]` | sidebar 写短标签（如 "HB-HAGT\n双路径\n特征融合"），body 写详细内容 |
| 文字溢出 textbox | 中文内容比 placeholder 长 | 缩短 bullet 文字、手动加 `\n` 换行，最后才调字号 |
| 封面信息不完整 | 模板使用半角空格分隔 | 从 dump 文件复制原始格式字符串 |
| 模板 slide placeholder 类型不一致 | 同一模板类别的不同 slide 可能使用不同 placeholder（如 T10 用 `[模型描述与创新点]` 但 T11/T12 用 `[实验结果表格或分析]`） | dump 后逐个检查 slide 的实际 placeholder 文本，不要假设同类 slide 用相同 keyword |
| 构建后 lxml 操作 `_sldIdLst` 插入/重排 slide 导致 ZIP 损坏 | python-pptx 内部维护的 rels 索引与 lxml 直接操作不一致，产生 "Duplicate name: slideXX.xml" | 所有 slide 增删和顺序调整必须在构建时通过 `remove_slides()` + KEEP set 一次性完成，禁止构建后操作 `_sldIdLst` |
| `find_shape_by_text` 搜不到表格内文字 | `shape.has_text_frame` 不覆盖 `shape.table` 的 cell 文本 | 搜索表格文字时额外遍历 `shape.table.rows[].cells[]` 的 `cell.text` |
| clear_content_shapes 清不掉 placeholder | 模板 placeholder 文本不在 keyword 过滤列表中 | 改用激进清理：删除 header 以下所有 TEXT_BOX (type 17) shape，再从头添加新 textbox |
