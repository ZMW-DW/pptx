"""Run the minimal_markdown thesis-defense-pptx example end-to-end."""
from __future__ import annotations

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


def run(cmd: list[str]) -> None:
    print(">>", " ".join(str(c) for c in cmd))
    subprocess.run(cmd, check=True)


def main() -> int:
    run([sys.executable, str(HERE / "build_template.py")])
    run([sys.executable, str(HERE / "build_deck.py")])
    run([
        sys.executable,
        str(SCRIPTS / "dump_pptx_content.py"),
        "--pptx",
        str(HERE / "final.pptx"),
        "--output",
        str(HERE / "dump.md"),
    ])
    run([
        sys.executable,
        str(SCRIPTS / "scan_pptx_text.py"),
        "--pptx",
        str(HERE / "final.pptx"),
        "--bad",
        "TODO,占位,项目简介,系统框图,算法设计,成果展示,总结分析",
        "--json",
        str(HERE / "scan.json"),
    ])
    run([sys.executable, str(HERE / "render_preview.py")])
    run([
        sys.executable,
        str(SCRIPTS / "make_contact_sheet.py"),
        "--input",
        str(HERE / "rendered_slides"),
        "--output",
        str(HERE / "contact_sheet.png"),
        "--cols",
        "2",
        "--thumb-width",
        "480",
        "--thumb-height",
        "270",
    ])
    print()
    print("done. outputs in", HERE)
    return 0


if __name__ == "__main__":
    sys.exit(main())
