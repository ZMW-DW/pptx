"""Run the minimal_markdown thesis-defense-pptx example end-to-end.

When Windows + Microsoft PowerPoint are available, the example uses the
official `export_pptx_png.ps1` to produce a faithful visual preview. On
environments without PowerPoint (Linux / macOS / minimal Windows), it falls
back to `render_preview.py`, which now renders a deliberately spartan
placeholder so the contact sheet does NOT misrepresent the generated PPTX.

Optional: pass --template <path-to-real.pptx> to run the example against a
real institutional PowerPoint template. In that mode build_deck.py overlays
the generated thesis.md example content onto native template pages, producing
an editable final.pptx that is genuinely generated from the example input.
The script then runs dump / scan / PowerPoint PNG export / contact-sheet
quality gates against that generated deck.
Real templates are never committed; only run artifacts under
examples/minimal_markdown/ are produced.
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent.parent
SCRIPTS = REPO_ROOT / "skills" / "thesis-defense-pptx" / "scripts"

EXPORT_PS1 = SCRIPTS / "export_pptx_png.ps1"
SAMPLE_TEMPLATE = HERE / "sample_template.pptx"
FINAL_PPTX = HERE / "final.pptx"
RENDERED_DIR = HERE / "rendered_slides"
CONTACT_SHEET = HERE / "contact_sheet.png"
EXPECTED_DIR = HERE / "expected"

def slide_number(path: Path) -> int:
    """Extract a 1-based slide number from PowerPoint-exported PNG names."""
    match = re.search(r"(\d+)", path.stem)
    return int(match.group(1)) if match else 0


def rendered_pngs() -> list[Path]:
    """Return exported slide PNGs once, sorted by slide number.

    On Windows, ``Path.glob("*.png")`` also matches ``.PNG`` files, so
    combining two glob patterns can silently duplicate every slide.
    """
    seen: dict[str, Path] = {}
    for path in RENDERED_DIR.glob("*"):
        if path.is_file() and path.suffix.lower() == ".png":
            seen[str(path.resolve()).lower()] = path
    return sorted(seen.values(), key=slide_number)


def run(cmd: list[str]) -> None:
    print(">>", " ".join(str(c) for c in cmd))
    subprocess.run(cmd, check=True)


def can_try_powerpoint() -> bool:
    """Cheap pre-check: only attempt the COM export when on Windows + powershell."""
    return sys.platform == "win32" and shutil.which("powershell") is not None


def export_with_powerpoint() -> bool:
    """Use export_pptx_png.ps1 to render real PNGs via PowerPoint COM.

    We don't probe COM availability ahead of time (the probe is slow and
    flaky on first launch); we simply try the export and fall back to PIL
    if the PowerShell process fails or produces no PNGs.
    """
    if RENDERED_DIR.exists():
        for old in RENDERED_DIR.glob("*.png"):
            try:
                old.unlink()
            except OSError:
                pass
        for old in RENDERED_DIR.glob("*.PNG"):
            try:
                old.unlink()
            except OSError:
                pass
    RENDERED_DIR.mkdir(parents=True, exist_ok=True)
    try:
        # Give PowerPoint plenty of time to launch and export the deck.
        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(EXPORT_PS1),
            "-Pptx",
            str(FINAL_PPTX),
            "-OutDir",
            str(RENDERED_DIR),
            "-Width",
            "1600",
            "-Height",
            "900",
        ]
        print(">>", " ".join(str(c) for c in cmd))
        subprocess.run(cmd, check=True, timeout=180)
        produced = rendered_pngs()
        return bool(produced)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        print(f"(PowerPoint export failed: {exc} — falling back to PIL placeholder)")
        return False


def write_expected_contact_sheet(files: list[Path], output: Path, *, cols: int = 2) -> None:
    """Write a README-friendly contact sheet for a generated slide subset."""
    from PIL import Image, ImageDraw, ImageFont

    def label_font() -> ImageFont.ImageFont:
        for font_path in (
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ):
            try:
                return ImageFont.truetype(font_path, 18)
            except OSError:
                continue
        return ImageFont.load_default()

    thumb_w, thumb_h = 760, 428
    margin = 22
    label_h = 28
    rows = (len(files) + cols - 1) // cols
    sheet = Image.new(
        "RGB",
        (
            cols * thumb_w + (cols + 1) * margin,
            rows * (thumb_h + label_h) + (rows + 1) * margin,
        ),
        "white",
    )
    draw = ImageDraw.Draw(sheet)
    font = label_font()
    for idx, path in enumerate(files):
        img = Image.open(path).convert("RGB").resize((thumb_w, thumb_h), Image.LANCZOS)
        col = idx % cols
        row = idx // cols
        x = margin + col * (thumb_w + margin)
        y = margin + row * (thumb_h + label_h + margin)
        draw.text((x, y), f"Slide {slide_number(path) or idx + 1}", fill=(0, 0, 0), font=font)
        sheet.paste(img, (x, y + label_h))
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output, optimize=True, compress_level=9)
    print(f"expected contact sheet: {output}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--template",
        default=None,
        help=(
            "Optional path to a real .pptx template. In this mode build_deck.py "
            "overlays thesis.md example content onto native template pages "
            "and writes a generated final.pptx."
        ),
    )
    parser.add_argument(
        "--contact-sheet-cols",
        type=int,
        default=None,
        help=(
            "Override columns in the final contact sheet. Defaults to 2 for "
            "small generated decks and 5 for larger decks."
        ),
    )
    parser.add_argument(
        "--detail-slides",
        default=None,
        help=(
            "Comma-separated 1-based slide numbers to also save as "
            "expected/detail_slide_NN.png (e.g. '1,5'). When omitted in "
            "--template mode the script writes detail snapshots for slides 1 "
            "and 5."
        ),
    )
    args = parser.parse_args()

    if args.template:
        custom = Path(args.template).expanduser().resolve()
        if not custom.is_file():
            print(f"--template path not found: {custom}")
            return 2
        print(f"using custom template: {custom}")
        run([
            sys.executable,
            str(HERE / "build_deck.py"),
            "--template",
            str(custom),
            "--out",
            str(FINAL_PPTX),
            "--real-template",
        ])
    else:
        run([sys.executable, str(HERE / "build_template.py")])
        run([sys.executable, str(HERE / "build_deck.py")])
    run([
        sys.executable,
        str(SCRIPTS / "dump_pptx_content.py"),
        "--pptx",
        str(FINAL_PPTX),
        "--output",
        str(HERE / "dump.md"),
    ])
    run([
        sys.executable,
        str(SCRIPTS / "scan_pptx_text.py"),
        "--pptx",
        str(FINAL_PPTX),
        "--bad",
        "TODO,占位,项目简介,系统框图,算法设计,成果展示,总结分析",
        "--json",
        str(HERE / "scan.json"),
    ])

    used_powerpoint = False
    if can_try_powerpoint():
        used_powerpoint = export_with_powerpoint()
    if not used_powerpoint:
        print("(no PowerPoint COM — using PIL placeholder preview)")
        run([sys.executable, str(HERE / "render_preview.py")])

    rendered = rendered_pngs()
    if args.contact_sheet_cols is not None:
        cols = args.contact_sheet_cols
    elif len(rendered) > 8:
        cols = 5
    else:
        cols = 2

    run([
        sys.executable,
        str(SCRIPTS / "make_contact_sheet.py"),
        "--input",
        str(RENDERED_DIR),
        "--output",
        str(CONTACT_SHEET),
        "--cols",
        str(cols),
        "--thumb-width",
        "800",
        "--thumb-height",
        "450",
    ])

    if args.detail_slides or args.template:
        # Re-extract the rendered list ordered by slide number so that
        # "slide 1" really maps to the first rendered slide (PowerPoint emits
        # locale-named files like "幻灯片1.PNG" in non-zero-padded order).
        rendered_sorted = rendered_pngs()

        if args.detail_slides:
            wanted = [int(x) for x in args.detail_slides.split(",") if x.strip()]
        else:
            wanted = [1, min(5, len(rendered_sorted))]

        EXPECTED_DIR.mkdir(parents=True, exist_ok=True)
        # Clear any stale detail snapshots before writing fresh ones.
        for old in EXPECTED_DIR.glob("detail_slide_*.png"):
            try:
                old.unlink()
            except OSError:
                pass

        for n in wanted:
            if 1 <= n <= len(rendered_sorted):
                src = rendered_sorted[n - 1]
                dst = EXPECTED_DIR / f"detail_slide_{n:02d}.png"
                shutil.copyfile(src, dst)
                print(f"detail snapshot: {dst}")

        if args.template:
            if rendered_sorted:
                write_expected_contact_sheet(
                    rendered_sorted[:4],
                    EXPECTED_DIR / "generated_overview_01.png",
                    cols=2,
                )
            if len(rendered_sorted) > 4:
                write_expected_contact_sheet(
                    rendered_sorted[4:8],
                    EXPECTED_DIR / "generated_overview_02.png",
                    cols=2,
                )

    print()
    print("done. outputs in", HERE)
    print("  preview source:", "PowerPoint COM" if used_powerpoint else "PIL placeholder")
    return 0


if __name__ == "__main__":
    sys.exit(main())
