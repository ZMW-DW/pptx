"""Build a small editable defense deck from thesis.md and sample_template.pptx.

This is intentionally a compact example, not a universal PPT generator. It
shows how an agent can preserve a template's visual language and use the helper
functions shipped in skills/thesis-defense-pptx/scripts/pptx_template_tools.py.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent.parent
SCRIPT_DIR = REPO_ROOT / "skills" / "thesis-defense-pptx" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from pptx_template_tools import add_para, add_text, rect, tag, write_table, RED, BLACK, WHITE  # noqa: E402

TEMPLATE = HERE / "sample_template.pptx"
THESIS = HERE / "thesis.md"
OUT = HERE / "final.pptx"


def section(text: str, title: str) -> str:
    pattern = re.compile(rf"^##\s+{re.escape(title)}\s*$", re.M)
    match = pattern.search(text)
    if not match:
        return ""
    next_match = re.search(r"^##\s+", text[match.end():], re.M)
    end = match.end() + next_match.start() if next_match else len(text)
    return text[match.end():end].strip()


def compact_lines(text: str, limit: int = 3) -> list[str]:
    sentences = [s.strip() for s in re.split(r"[。.!?]\s*", text) if s.strip()]
    lines = []
    for sentence in sentences[:limit]:
        if len(sentence) > 42:
            sentence = sentence[:39] + "..."
        lines.append("• " + sentence)
    return lines or ["• 内容待补充"]


def clear_slide(slide) -> None:
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)


def build_cover(slide, title: str) -> None:
    clear_slide(slide)
    band = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
    band.fill.solid()
    band.fill.fore_color.rgb = RED
    band.line.color.rgb = RED
    add_text(slide, "Minimal Markdown 示例", 0.65, 0.28, 4.5, 0.45, 16, bold=True, color=WHITE)
    add_text(slide, "毕业论文答辩", 0.9, 2.3, 5.2, 0.75, 34, bold=True, color=RED)
    add_text(slide, title, 0.95, 3.28, 9.5, 0.65, 20, bold=True, color=BLACK)
    add_text(slide, "可编辑 PPTX / 模板保持 / 质量检查", 0.95, 4.1, 7.6, 0.45, 15)
    add_text(slide, "Author: thesis-defense-pptx-skill", 0.95, 5.25, 5.8, 0.35, 12)


def build_background(slide, lines: list[str]) -> None:
    clear_slide(slide)
    tag(slide, "背景", 0.7, 0.25, 1.1, 0.38)
    add_text(slide, "研究背景", 0.75, 0.92, 3.2, 0.45, 24, bold=True, color=RED)
    rect(slide, 0.75, 1.55, 5.7, 4.7)
    rect(slide, 6.85, 1.55, 5.7, 4.7)
    add_text(slide, "问题场景", 1.05, 1.85, 4.8, 0.35, 18, bold=True, color=RED)
    add_para(slide, lines, 1.05, 2.35, 5.0, 3.25, 15)
    add_text(slide, "模板约束", 7.15, 1.85, 4.8, 0.35, 18, bold=True, color=RED)
    add_para(slide, ["• 不重新设计学校模板", "• 保留封面、导航、配色和卡片样式", "• 生成可编辑 PPTX，而不是图片页"], 7.15, 2.35, 5.0, 3.25, 15)


def build_method(slide, lines: list[str]) -> None:
    clear_slide(slide)
    tag(slide, "方法", 0.7, 0.25, 1.1, 0.38)
    add_text(slide, "方法设计", 0.75, 0.92, 3.2, 0.45, 24, bold=True, color=RED)
    rect(slide, 0.75, 1.55, 11.8, 4.7)
    add_para(slide, lines, 1.1, 1.95, 5.5, 3.6, 15)
    table_shape = slide.shapes.add_table(4, 2, Inches(7.0), Inches(2.0), Inches(4.7), Inches(2.6))
    write_table(
        table_shape.table,
        [
            ["阶段", "产物"],
            ["读取论文", "thesis_context"],
            ["套用模板", "editable pptx"],
            ["质量检查", "contact sheet"],
        ],
        font_size=12,
    )


def build_result(slide, lines: list[str]) -> None:
    clear_slide(slide)
    tag(slide, "结果", 0.7, 0.25, 1.1, 0.38)
    add_text(slide, "结果与总结", 0.75, 0.92, 3.2, 0.45, 24, bold=True, color=RED)
    rect(slide, 0.75, 1.55, 5.7, 4.7)
    rect(slide, 6.85, 1.55, 5.7, 4.7)
    add_text(slide, "最小示例结果", 1.05, 1.85, 4.8, 0.35, 18, bold=True, color=RED)
    add_para(slide, lines, 1.05, 2.35, 5.0, 3.25, 15)
    add_text(slide, "检查项", 7.15, 1.85, 4.8, 0.35, 18, bold=True, color=RED)
    add_para(slide, ["✓ dump shape/text", "✓ scan stale words", "✓ render preview PNG", "✓ contact sheet"], 7.15, 2.35, 5.0, 3.25, 16)


def main() -> int:
    if not TEMPLATE.exists():
        raise SystemExit(f"template not found: {TEMPLATE}")
    text = THESIS.read_text(encoding="utf-8")
    prs = Presentation(str(TEMPLATE))
    while len(prs.slides) < 4:
        prs.slides.add_slide(prs.slide_layouts[6])

    title = text.splitlines()[0].lstrip("# ").strip()
    build_cover(prs.slides[0], title)
    build_background(prs.slides[1], compact_lines(section(text, "研究背景"), 3))
    build_method(prs.slides[2], compact_lines(section(text, "方法设计"), 4))
    result_text = section(text, "实验结果") + "\n" + section(text, "结论")
    build_result(prs.slides[3], compact_lines(result_text, 4))

    prs.save(OUT)
    print(OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
