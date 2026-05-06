"""Run the minimal_markdown thesis-defense-pptx example end-to-end.

When Windows + Microsoft PowerPoint are available, the example uses the
official `export_pptx_png.ps1` to produce a faithful visual preview. On
environments without PowerPoint (Linux / macOS / minimal Windows), it falls
back to `render_preview.py`, which now renders a deliberately spartan
placeholder so the contact sheet does NOT misrepresent the generated PPTX.

Optional: pass --template <path-to-real.pptx> to run the example against a
real institutional PowerPoint template. In that mode the script does NOT
touch build_template.py / build_deck.py — those two are demo-grade and assume
a clean 4-page generated skeleton. Instead it keeps the first N slides of
your real template (default 4, raise with --max-slides or pass --full to keep
the entire deck) and runs the three quality-gate scripts (dump shape/text,
scan stale words, export real PowerPoint PNGs) against the template directly.
In --full mode it can also write README-friendly expected contact sheets and
detail snapshots from the exported PowerPoint PNGs.
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

# Default number of slides kept when --template is given without --full /
# --max-slides (matches the 4-page demo skeleton from build_deck.py).
DEFAULT_PREVIEW_SLIDES = 4


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


def truncate_pptx_to_first_n_slides(src: Path, dst: Path, n: int | None) -> int:
    """Copy ``src`` to ``dst``, optionally keeping only the first ``n`` slides.

    python-pptx does not expose slide deletion, so we manipulate the
    sldIdLst XML directly. The underlying slideN.xml parts remain inside the
    package zip but PowerPoint only renders slides referenced by sldIdLst,
    which is exactly what we want for the example. Pass ``n=None`` to keep
    the whole deck. Returns the number of slides actually kept.
    """
    from pptx import Presentation

    prs = Presentation(str(src))
    sld_id_lst = prs.slides._sldIdLst
    children = list(sld_id_lst)
    total = len(children)
    if n is not None:
        for child in children[n:]:
            sld_id_lst.remove(child)
        kept = min(n, total)
    else:
        kept = total
    dst.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dst))
    return kept


def write_expected_contact_sheet(files: list[Path], output: Path) -> None:
    """Write a README-friendly 5-column contact sheet for a slide subset."""
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

    cols = 5
    thumb_w, thumb_h = 560, 315
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
            "is skipped (it assumes the demo skeleton produced by "
            "build_template.py and would corrupt complex real-world slides). "
            "By default the first 4 slides become final.pptx; raise this with "
            "--max-slides or pass --full to keep the whole deck."
        ),
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help=(
            "Keep every slide of the --template deck. Combine with "
            "--max-slides N to keep all slides up to N (handy for excluding "
            "promotional last pages such as QR-code shoutouts that "
            "community-shared templates often append)."
        ),
    )
    parser.add_argument(
        "--max-slides",
        type=int,
        default=None,
        help=(
            "Upper bound on slides kept from --template. When omitted, the "
            f"default is {DEFAULT_PREVIEW_SLIDES} for the small demo path, "
            "or unbounded when --full is set."
        ),
    )
    parser.add_argument(
        "--contact-sheet-cols",
        type=int,
        default=None,
        help=(
            "Override columns in the final contact sheet. Defaults to 2 for "
            "the demo path / small decks and 5 when running with --full."
        ),
    )
    parser.add_argument(
        "--detail-slides",
        default=None,
        help=(
            "Comma-separated 1-based slide numbers to also save as "
            "expected/detail_slide_NN.png (e.g. '1,5'). When omitted in "
            "--full mode the script picks the cover (slide 1) plus the first "
            "content slide (slide 5)."
        ),
    )
    parser.add_argument(
        "--expected-exclude-slides",
        default="",
        help=(
            "Comma-separated 1-based slide numbers to omit only from the "
            "README expected contact sheets written in --full mode. The "
            "slides are still kept in final.pptx and still go through dump / "
            "scan / export quality gates."
        ),
    )
    args = parser.parse_args()

    if args.template:
        custom = Path(args.template).expanduser().resolve()
        if not custom.is_file():
            print(f"--template path not found: {custom}")
            return 2
        # Resolve final "keep N" semantics:
        #   - --max-slides N always wins (acts as an upper bound, even with --full)
        #   - --full alone keeps everything
        #   - neither: small demo (DEFAULT_PREVIEW_SLIDES)
        if args.max_slides is not None:
            keep_n = max(1, args.max_slides)
        elif args.full:
            keep_n = None
        else:
            keep_n = DEFAULT_PREVIEW_SLIDES
        kept = truncate_pptx_to_first_n_slides(custom, FINAL_PPTX, keep_n)
        print(f"using custom template: {custom}")
        if keep_n is None:
            print(f"  kept all {kept} slides (--full) -> {FINAL_PPTX}")
        elif args.full:
            print(f"  kept first {kept} slides (--full + --max-slides {keep_n}) -> {FINAL_PPTX}")
        else:
            print(f"  kept first {kept} slides -> {FINAL_PPTX}")
        print("  (build_template / build_deck skipped: --template runs the")
        print("   dump/scan/export quality gates against the real template)")
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

    if args.detail_slides or args.full:
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

        if args.full:
            excluded = {
                int(x)
                for x in args.expected_exclude_slides.split(",")
                if x.strip()
            }
            readme_slides = [
                p for p in rendered_sorted
                if slide_number(p) not in excluded
            ]
            if readme_slides:
                write_expected_contact_sheet(
                    readme_slides[:10],
                    EXPECTED_DIR / "contact_sheet_01_10.png",
                )
            if len(readme_slides) > 10:
                write_expected_contact_sheet(
                    readme_slides[10:20],
                    EXPECTED_DIR / "contact_sheet_11_20.png",
                )

    print()
    print("done. outputs in", HERE)
    print("  preview source:", "PowerPoint COM" if used_powerpoint else "PIL placeholder")
    return 0


if __name__ == "__main__":
    sys.exit(main())
