"""Render a lightweight visual preview from final.pptx.

python-pptx cannot render PowerPoint exactly like Microsoft PowerPoint. This
script deliberately creates an approximate PNG preview from slide text so the
minimal example can run cross-platform. Real delivery should still use
PowerPoint export_pptx_png.ps1.
"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

HERE = Path(__file__).resolve().parent
PPTX = HERE / "final.pptx"
OUT_DIR = HERE / "rendered_slides"

W, H = 1280, 720
RED = (0x83, 0x06, 0x04)
DARK = (38, 38, 38)
LIGHT = (247, 243, 241)


def font(size: int, bold: bool = False):
    candidates = [
        "C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def extract_texts(slide) -> list[str]:
    texts = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = shape.text.strip()
            if text:
                texts.append(text)
        if getattr(shape, "has_table", False):
            rows = []
            for row in shape.table.rows:
                rows.append(" | ".join(cell.text.strip() for cell in row.cells))
            if rows:
                texts.append("\n".join(rows))
    return texts


def render_slide(idx: int, texts: list[str]) -> Image.Image:
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)
    title_font = font(34, True)
    body_font = font(22)
    small_font = font(18)

    draw.rectangle((0, 0, W, 82), fill=RED)
    draw.text((48, 24), "thesis-defense-pptx-skill minimal example", fill="white", font=small_font)

    title = texts[0] if texts else f"Slide {idx}"
    draw.text((70, 130), title[:46], fill=RED, font=title_font)

    card1 = (70, 215, 610, 610)
    card2 = (670, 215, 1210, 610)
    for box in (card1, card2):
        draw.rounded_rectangle(box, radius=18, fill=LIGHT, outline=RED, width=3)

    body_items = texts[1:] or ["(empty slide)"]
    left = body_items[: max(1, len(body_items) // 2)]
    right = body_items[max(1, len(body_items) // 2):]

    y = 245
    for item in left[:5]:
        draw.multiline_text((100, y), item[:180], fill=DARK, font=body_font, spacing=8)
        y += 92

    y = 245
    for item in right[:5]:
        draw.multiline_text((700, y), item[:180], fill=DARK, font=body_font, spacing=8)
        y += 92

    draw.text((70, 650), f"Slide {idx}", fill=DARK, font=small_font)
    return img


def main() -> int:
    prs = Presentation(str(PPTX))
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for old in OUT_DIR.glob("*.png"):
        old.unlink()
    for idx, slide in enumerate(prs.slides, 1):
        img = render_slide(idx, extract_texts(slide))
        out = OUT_DIR / f"slide_{idx:02d}.png"
        img.save(out)
        print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
