# expected/

This directory stores reference visual outputs for the minimal example.

`contact_sheet.png` is the screenshot referenced from the top-level README. It
is intentionally generated against a **real institutional PowerPoint template**
(in this case a Zhengzhou University defense/report template) by passing the
template path via `--template`, so the screenshot shows what the skill's
quality-gate scripts produce on a realistic 16:9 deck instead of the
deliberately spartan demo skeleton:

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-template.pptx
cp examples/minimal_markdown/contact_sheet.png \
   examples/minimal_markdown/expected/contact_sheet.png
```

In `--template` mode the example keeps only the first 4 slides of the real
template as `final.pptx`, **skips** `build_template.py` / `build_deck.py`
(those assume the generated demo skeleton and would corrupt complex
real-world slides), and runs:

- `dump_pptx_content.py` to dump the template's shape/text/table inventory,
- `scan_pptx_text.py` to flag stale placeholder words such as `占位` /
  `Copy paste fonts.` / `添加标题`,
- `export_pptx_png.ps1` to render real PNGs via PowerPoint COM,
- `make_contact_sheet.py` to assemble the 2x2 grid you see committed here.

Without `--template`, `run_example.py` falls back to the demo path: it
synthesizes `sample_template.pptx` from `build_template.py`, runs
`build_deck.py` to rebuild the four pages, and (if PowerPoint COM is not
available) renders a deliberately spartan PIL placeholder. That output is
fully reproducible but visually plain — the committed `expected/contact_sheet.png`
is **not** what the default path produces.

For real thesis delivery, export slides with PowerPoint
(`export_pptx_png.ps1`), scan them with `inspect_pptx_overflow.ps1`, and run
`make_contact_sheet.py` on the exported PNG directory.

## Acknowledging template provenance

The committed PNG shows fragments of a publicly shared Zhengzhou-University-
themed PPT template (the cover line "汇报人：郑大球知道", the school crest,
and the watermark "DESIGNED BY QIANKU" / "made by jiaoyang" all come from
that template's own designers). They are reproduced here only as part of the
quality-gate demonstration. This repository does not redistribute the
template itself; users must obtain templates from their own institutional
or licensed sources.
