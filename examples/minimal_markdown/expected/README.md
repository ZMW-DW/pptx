# expected/

This directory stores reference visual outputs for the minimal example.

The committed images are generated against a **real institutional PowerPoint
template** (a Zhengzhou University defense/report template) by passing the
template path via `--template --full`. They show what the skill's quality-gate
scripts produce on a realistic 16:9 deck instead of the deliberately spartan
demo skeleton.

## Files

- `contact_sheet_01_10.png` — PowerPoint-COM renders of slides 1-10,
  assembled as a 5x2 overview.
- `contact_sheet_11_20.png` — PowerPoint-COM renders of slides 11-19 and 21,
  assembled as a second 5x2 overview.
- `detail_slide_01.png` — full-resolution cover slide detail.
- `detail_slide_14.png` — full-resolution content slide detail.

Slide 20 in this specific publicly shared template is a QR-code promotional
page. It is still kept in `final.pptx` and exported under `rendered_slides/`
when the example is run, so the quality gate remains full-deck; it is omitted
only from the public README contact sheets.

Generate the images with:

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-template.pptx \
    --full \
    --expected-exclude-slides "20" \
    --detail-slides "1,14"
```

In `--template` mode the example **skips** `build_template.py` /
`build_deck.py`: those scripts assume the generated demo skeleton and would
corrupt complex real-world slides if they tried to clear/rebuild them. The
real template pages are kept natively, then the example runs:

- `dump_pptx_content.py` to dump the template's shape/text/table inventory,
- `scan_pptx_text.py` to flag stale placeholder words,
- `export_pptx_png.ps1` to render real PNGs via PowerPoint COM,
- `make_contact_sheet.py` plus README-specific contact-sheet generation.

Without `--template`, `run_example.py` falls back to the demo path: it
synthesizes `sample_template.pptx` from `build_template.py`, runs
`build_deck.py` to rebuild four pages, and (if PowerPoint COM is not
available) renders a deliberately spartan PIL placeholder. That output is
fully reproducible but visually plain; the committed images here demonstrate
behaviour on a realistic institutional template.

## Acknowledging Template Provenance

The committed PNGs show fragments of a publicly shared Zhengzhou-University-
themed PPT template. The cover line "汇报人：郑大球知道", the school crest,
and the watermark text such as "DESIGNED BY QIANKU" / "made by jiaoyang"
come from that template's own designers. They are reproduced here only as
part of the quality-gate demonstration. This repository does not redistribute
the template itself; users must obtain templates from their own institutional
or licensed sources.
