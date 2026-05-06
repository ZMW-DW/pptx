# expected/

This directory stores reference visual outputs for the `minimal_markdown`
example.

The committed images are generated against a **real institutional PowerPoint
template** (a Zhengzhou University defense/report template), but they are
not raw template-page screenshots. They are renders of the generated
`final.pptx` after `build_deck.py --real-template` overlays content from
`thesis.md` onto native template pages.

## Files

- `generated_overview_01.png` — PowerPoint-COM renders of generated slides
  1-4, assembled as a 2x2 overview.
- `generated_overview_02.png` — PowerPoint-COM renders of generated slides
  5-8, assembled as a second 2x2 overview.
- `detail_slide_01.png` — full-resolution generated cover slide.
- `detail_slide_06.png` — full-resolution generated method slide, including
  generated bullet text, workflow cards, and a small table.

Every generated slide contains a `Generated from thesis.md` badge so readers
can distinguish these images from raw template screenshots.

Generate the images with:

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-template.pptx \
    --detail-slides "1,6"
```

In this mode the example:

1. opens the real template without clearing its complex native shapes,
2. overlays generated `thesis.md` content onto selected template pages,
3. writes an editable `final.pptx`,
4. runs `dump_pptx_content.py` and `scan_pptx_text.py`,
5. exports faithful PNGs via PowerPoint COM,
6. writes the README contact sheets and detail snapshots from those PNGs.

Without `--template`, `run_example.py` falls back to the demo path: it
synthesizes `sample_template.pptx` from `build_template.py`, runs
`build_deck.py` to rebuild four pages, and (if PowerPoint COM is not
available) renders a deliberately spartan PIL placeholder. That output is
fully reproducible but visually plain; the committed images here demonstrate
the same generated example content on a realistic institutional template.

## Acknowledging Template Provenance

The committed PNGs show fragments of a publicly shared Zhengzhou-University-
themed PPT template. The school crest and watermark text such as
"DESIGNED BY QIANKU" / "made by jiaoyang" come from that template's own
designers. They are reproduced here only as part of the quality-gate
demonstration. This repository does not redistribute the template itself;
users must obtain templates from their own institutional or licensed sources.
