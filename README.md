# Thesis Defense PPTX Skill

A reusable Codex/Agent Skill for creating editable thesis defense PowerPoint decks from local thesis files and an existing `.pptx` visual template.

The workflow is designed for cases where template fidelity matters: university defense decks, lab report decks, branded academic presentations, and local private thesis projects.

## What It Does

- Reads thesis source material from a local PDF/LaTeX project.
- Preserves an existing PowerPoint template's cover, typography, colors, navigation, card style, and slide proportions.
- Builds editable `.pptx` decks instead of image-only slides.
- Exports slides to PNG for visual review.
- Generates a contact sheet for whole-deck inspection.
- Uses PowerPoint COM to detect risky text overflow based on real PowerPoint rendering.
- Scans for stale template text and placeholders before delivery.

## Repository Layout

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

## Requirements

- Windows is recommended for full quality checking.
- Microsoft PowerPoint is required for COM-based slide cloning, PNG export, and overflow inspection.
- Python 3.10+.
- Python packages:
  - `python-pptx`
  - `Pillow`
  - `PyMuPDF` or `pypdf` for PDF extraction

Install the Python dependencies:

```powershell
python -m pip install python-pptx Pillow PyMuPDF pypdf
```

## Install Locally

Copy the skill folder into your Codex skills directory:

```powershell
Copy-Item -Recurse -Force `
  .\skills\thesis-defense-pptx `
  "$env:USERPROFILE\.codex\skills\thesis-defense-pptx"
```

Then start a new Codex session and ask for a thesis defense PPTX, or explicitly mention:

```text
Use the thesis-defense-pptx skill.
```

## Core Workflow

1. Extract thesis context from local source files.
2. Analyze the supplied PowerPoint template.
3. Clone native template slides with PowerPoint COM.
4. Fill concise defense content into editable PPTX shapes.
5. Export slides to PNG.
6. Create a contact sheet and inspect it.
7. Run text overflow and stale-template checks.
8. Iterate until the deck passes the quality gate.

## Example Commands

Extract thesis context:

```powershell
python .\skills\thesis-defense-pptx\scripts\extract_thesis_context.py `
  --input "D:\path\to\thesis-project" `
  --output "D:\path\to\thesis_context.md"
```

Export a deck to PNG:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\thesis-defense-pptx\scripts\export_pptx_png.ps1 `
  -Pptx "D:\path\to\deck.pptx" `
  -OutDir "D:\path\to\visual_check" `
  -Width 1600 -Height 900
```

Inspect text overflow:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\thesis-defense-pptx\scripts\inspect_pptx_overflow.ps1 `
  -Pptx "D:\path\to\deck.pptx" `
  -Tolerance 40
```

Create a contact sheet:

```powershell
python .\skills\thesis-defense-pptx\scripts\make_contact_sheet.py `
  --input "D:\path\to\visual_check" `
  --output "D:\path\to\contact_sheet.png"
```

## Notes

This skill intentionally does not ship a built-in slide template. It is meant to preserve the user's supplied template rather than impose a generic presentation style.

## License

MIT
