"""Run the minimal_markdown thesis-defense-pptx example end-to-end.

When Windows + Microsoft PowerPoint are available, the example uses the
official `export_pptx_png.ps1` to produce a faithful visual preview. On
environments without PowerPoint (Linux / macOS / minimal Windows), it falls
back to `render_preview.py`, which now renders a deliberately spartan
placeholder so the contact sheet does NOT misrepresent the generated PPTX.
"""
from __future__ import annotations

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
FINAL_PPTX = HERE / "final.pptx"
RENDERED_DIR = HERE / "rendered_slides"
CONTACT_SHEET = HERE / "contact_sheet.png"


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
        produced = list(RENDERED_DIR.glob("*.PNG")) + list(RENDERED_DIR.glob("*.png"))
        return bool(produced)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        print(f"(PowerPoint export failed: {exc} — falling back to PIL placeholder)")
        return False


def main() -> int:
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

    run([
        sys.executable,
        str(SCRIPTS / "make_contact_sheet.py"),
        "--input",
        str(RENDERED_DIR),
        "--output",
        str(CONTACT_SHEET),
        "--cols",
        "2",
        "--thumb-width",
        "800",
        "--thumb-height",
        "450",
    ])
    print()
    print("done. outputs in", HERE)
    print("  preview source:", "PowerPoint COM" if used_powerpoint else "PIL placeholder")
    return 0


if __name__ == "__main__":
    sys.exit(main())
