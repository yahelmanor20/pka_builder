---
name: hebrew-pka-docx
description: >-
  Generate Hebrew RTL Word (DOCX) documents for the 8222 / ביה"ס ללחימה project —
  PKA (פקודת אימון) and training-requirements (דרישות אימון) docs — and any other
  Hebrew RTL DOCX. Use when the user asks to build, produce, or export a פק"א, a
  דרישות אימון sheet, or any Hebrew Word document that must render right-to-left
  with correct bidi, David font, gray fills, and blue headers. Wraps the project
  generator tools/gen_docx.py and folds in Hebrew First RTL/bidi/font expertise.
  Merged from skills-il/hebrew-document-generator (MIT) + the pka_builder project.
license: MIT
allowed-tools: Bash(python:*) Bash(pip:*)
compatibility: >-
  Requires Python 3.9+ with python-docx (and python-bidi for any PDF path).
  Hebrew fonts (David, Heebo) should be available on the system for correct preview.
---

# Hebrew PKA / Requirements DOCX Builder

This skill merges the generic **Hebrew Document Generator** (RTL/bidi knowhow, font
catalog, python-docx techniques) with the **pka_builder** project's own canonical
generator and formatting rules. For PKA and requirements documents, always prefer the
project generator — it already encodes the exact template styling. Use the generic
Hebrew-First techniques only for one-off documents the project generator does not cover.

## Decision: which path to take

| Task | Path | Tool |
|------|------|------|
| פק"א מלאה (full training order) | **Project generator** | `tools/gen_docx.py` |
| דרישות אימון (requirements sheet) | Project derivation rules | see `requirements-builder.skill` + CLAUDE.md |
| Any other Hebrew RTL Word doc | Generic python-docx | `scripts/generate_doc.py` + `references/rtl-docx.md` |
| Hebrew PDF (invoice/receipt/etc.) | Generic reportlab | `scripts/generate_doc.py` |

## Step 1 — PKA / Requirements (the primary use case)

The canonical, project-specific generator is **[tools/gen_docx.py](../../../tools/gen_docx.py)**.
It already implements every formatting rule from `CLAUDE.md`:

- `David` font everywhere, size 11 body / blue `1F4E79` titles
- Gray fills `GRAY_HEAD=BFBFBF` (header rows), `GRAY_DARK=A6A6A6`
- Full-width tables (`FULL_W = 9360` twips), fixed layout
- Table-level `w:bidiVisual` for correct RTL column mirroring
- Paragraph-level `w:bidi` + run-level `w:rtl` on every run (`_run` / `_prtl` / `_cell` helpers)
- One table per day, page break between days; merged rows keep the hour cell rightmost

**Do not re-implement RTL DOCX from scratch for a PKA.** Extend `tools/gen_docx.py`
instead — its helpers (`hdr`, `title`, `bullet`, `plain`, `grid_table`, `_cell`,
`_full_width`, `_borders`) are the building blocks. Output filename convention:
`פקא לאימון גדוד X שבוע Y`.

The PKA-specific mapping/format rules (column order, parallel-activity routing,
location defaults, אמ"מ 2026 day models) live in `CLAUDE.md` — consult it before
building achievement lists or laying out the daily לו"ז.

## Step 2 — Generic Hebrew RTL DOCX (everything else)

For ad-hoc Hebrew Word documents not covered by the project template, use python-docx
with the RTL helpers. The two mandatory pieces of RTL that python-docx does NOT set by
default:

```python
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def set_paragraph_rtl(paragraph):
    """Paragraph reads right-to-left."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.insert(0, pPr.makeelement(qn('w:bidi'), {}))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def set_run_rtl(run):
    """Run renders as RTL (Hebrew) text."""
    rPr = run._r.get_or_add_rPr()
    rPr.append(rPr.makeelement(qn('w:rtl'), {}))
```

For tables that must mirror their columns right-to-left, also append `w:bidiVisual`
to the table's `tblPr` (see `_borders`/`grid_table` in `tools/gen_docx.py` for the
exact pattern). Full technique notes, gotchas, and font choices are in
`references/rtl-docx.md` and `references/hebrew-fonts.md`.

## Step 3 — Hebrew PDF (optional)

`scripts/generate_doc.py` (from Hebrew First) produces RTL PDFs with reportlab +
python-bidi (invoices/receipts). Run `python scripts/generate_doc.py --help`. Not used
for PKA output, which is always DOCX, but available if a PDF deliverable is requested.

## Critical RTL gotchas (apply to every path)

- **bidi is not idempotent.** With reportlab, call `get_display()` once per line at draw
  time. Never pre-process a whole document through it and then call it again. (python-docx
  does NOT need `get_display` — Word does the bidi at render time given `w:rtl`/`w:bidi`.)
- **Pass mixed Hebrew+number lines whole** to the bidi algorithm so digits, currency, and
  punctuation land in the right LTR positions. Do not split and reorder yourself.
- **Font must cover U+0590–U+05FF.** David / Heebo / Frank Ruhl Libre are safe; many Latin
  decorative fonts are not.
- **No nikud** in formal/business/military Hebrew — it looks unprofessional.
- **Dates** are DD/MM/YYYY in secular context.

## Bundled resources

- `scripts/generate_doc.py` — generic Hebrew PDF/DOCX generator (reportlab + python-bidi).
- `references/rtl-docx.md` — python-docx RTL techniques, bidi rules, troubleshooting.
- `references/hebrew-fonts.md` — Hebrew font catalog, pairings, install + registration.

## Provenance

The generic Hebrew document-generation material is adapted from
[skills-il/hebrew-document-generator](https://github.com/skills-il/localization/tree/master/hebrew-document-generator)
(MIT License). The PKA-specific generator, template, and formatting rules are original to
the pka_builder project.
