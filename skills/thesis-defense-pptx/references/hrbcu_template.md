# HRBCU Defense Template

Built-in reusable template for Harbin University of Commerce (哈尔滨商业大学) thesis defense presentations.

**File:** `templates/HRBCU-defense-template.pptx`

**Source:** Adapted from a real defense PPTX (author: 韩蓉, supervisor: 郑徳权教授). All personal and thesis-specific content has been replaced with `[placeholder]` labels. Structural and decorative elements are preserved.

## Template Specs

- **Slides:** 22
- **Aspect ratio:** 16:9 widescreen (13.33in × 7.50in)
- **Primary color:** Red `#830604` (configurable via `THESIS_PPTX_BRAND_RED` env var)
- **Fonts:** Default to template originals; CJK text uses system-installed fonts

## Slide Inventory

| Slide | Type | Layout | Placeholder / Notes |
|-------|------|--------|---------------------|
| 1 | Cover | 标题幻灯片 | `[论文标题]`, `[学院名称]`, `[专业名称]`, `[导师姓名]`, `[学生姓名]` |
| 2 | TOC | 空白 | 4 items with numbered circles: 研究背景与意义, 国内外研究现状, 研究内容与主要工作, 总结与展望 |
| 3 | Section divider | 空白 | Section "1" — 研究意义 |
| 4 | Content | 空白 | `[在此填入论文相关内容]` — single body text area |
| 5 | Section divider | 空白 | Section "2" — 研究现状 |
| 6 | Content | 空白 | `[在此填入论文相关内容]` — single body text area |
| 7 | Section divider | 空白 | Section "3" — 研究内容与主要工作 |
| 8 | Content | 空白 | `[模型描述与创新点]` — sidebar + body text. For first model/contribution description |
| 9 | Content | 空白 | `[模型架构图]` — architecture diagram placeholder. Left-side label area for model name |
| 10 | Content | 空白 | `[实验结果表格或分析]` — table/chart area + analysis text |
| 11 | Content | 空白 | `[实验结果表格或分析]` — double table + analysis |
| 12 | Content | 空白 | `[实验结果表格或分析]` — ablation experiment tables (left + right) |
| 13 | Content | 空白 | `[实验结果表格或分析]` — additional ablation tables |
| 14 | Content | 空白 | `[模型描述与创新点]` — sidebar + body text. For second model/contribution description |
| 15 | Content | 空白 | `[模型架构图]` — architecture diagram placeholder |
| 16 | Content | 空白 | `[实验结果表格或分析]` — experiment results with comparison table |
| 17 | Content | 空白 | `[实验结果表格或分析]` — ablation with chart area |
| 18 | Content | 空白 | `[实验结果表格或分析]` — additional ablation table |
| 19 | Section divider | 空白 | Section "4" — 总结 |
| 20 | Content | 空白 | `[研究总结与未来展望]` summary body text |
| 21 | Content | 空白 | `[研究总结与未来展望]` second summary page |
| 22 | Closing | 标题幻灯片 | "谢谢观看！敬请各位老师批评指正！" — generic closing |

## Common Decorative Shapes (per slide)

- **图片 1** (top-right, most slides): School logo badge
- **图片 3** (slide 1, 22): Cover/closing logo
- **矩形 28** (content slides): Red header accent bar at top
- **椭圆** (section dividers): Numbered circle indicators
- **L 形** (slides 8, 14): Sidebar accent shapes
- **箭头** (slides 9, 15): Arrow connector for architecture diagrams

## Content Replacement Strategy

When filling this template with a new thesis:

1. **Cover (slide 1):** Replace `[论文标题]` with the actual thesis title. Update student/supervisor/school/major info. Use `replace_exact_text()` or `set_text_preserve_style()`.

2. **TOC (slide 2):** Keep the 4-section structure. TOC labels are already generic. Optionally rename to match the specific thesis structure.

3. **Content slides:** Two approaches:
   - **Preserve and replace:** Use `replace_partial_text()` / `replace_exact_text()` to swap placeholder text, keeping existing shapes and positions.
   - **Clear and rebuild:** Use `clean_content_slide()` to remove all shapes except structural ones (logo, header bar), then build new content with `add_text()`, `add_para()`, `rect()`, `write_table()`, etc.

4. **Tables:** Use `write_table()` to fill experiment results while preserving cell formatting.

5. **Architecture slides (9, 15):** Replace `[模型架构图]` with an actual model diagram image or a text-based component flow.

6. **Quality gate:** After filling, run `dump_pptx_content.py`, `scan_pptx_text.py` (with bad-terms from the previous thesis author), and visual inspection.
