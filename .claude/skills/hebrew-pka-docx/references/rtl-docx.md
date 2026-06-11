# Hebrew RTL in Word (DOCX) — techniques & gotchas

Adapted from skills-il/hebrew-document-generator (MIT) and the pka_builder
`tools/gen_docx.py` generator.

## The three layers of RTL in a .docx

Word needs RTL set at three levels. python-docx sets *none* of them by default — you
must add the raw OOXML elements yourself.

| Level | OOXML element | Effect | Helper in `gen_docx.py` |
|-------|---------------|--------|--------------------------|
| Run | `w:rtl` (in `rPr`) | Characters render right-to-left | `_run()` |
| Paragraph | `w:bidi` (in `pPr`) | Paragraph base direction is RTL; align right | `_prtl()` |
| Table | `w:bidiVisual` (in `tblPr`) | Columns mirror right-to-left | `grid_table()` / `_borders()` |
| Section | `w:bidi` (in `sectPr`) | Whole section is RTL | template `pka_template.docx` |

Miss `w:bidiVisual` and your table columns render left-to-right even though each cell's
text is RTL — the single most common PKA layout bug.

## Minimal helpers

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), v)
    return e

def set_paragraph_rtl(paragraph):
    paragraph._p.get_or_add_pPr().insert(0, _el('w:bidi'))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def set_run_rtl(run):
    run._r.get_or_add_rPr().append(_el('w:rtl', **{'w:val': '1'}))

def set_table_rtl(table):
    table._tbl.tblPr.append(_el('w:bidiVisual'))
```

Always set the run font to a Hebrew-capable face on `w:rFonts` `ascii`/`hAnsi`/**`cs`**
(the `cs` = complex-script slot is what Word actually uses for Hebrew). See `_run()` in
`tools/gen_docx.py`:

```python
rf = _el('w:rFonts')
rf.set(qn('w:ascii'), 'David'); rf.set(qn('w:hAnsi'), 'David'); rf.set(qn('w:cs'), 'David')
rPr.insert(0, rf)
rPr.append(_el('w:szCs', **{'w:val': str(int(size * 2))}))  # complex-script size, half-points
```

## RTL column order for PKA tables

From `CLAUDE.md` (RTL, rightmost first):
- Normal day: שעה → פעילות → מיקום → גורם מעביר → הערות
- Split day: שעה → פלוגה א' → ב' → ג'
- Merged rows: the hour cell is always first (rightmost).

Because `w:bidiVisual` mirrors the columns, build the `headers`/`rows` lists in *logical*
(LTR source) order and let Word mirror them — verify against an example in `examples-pka/`.

## Gotchas

- **No `get_display()` for DOCX.** python-bidi's `get_display` is for fixed-layout
  renderers (reportlab PDF). Word reorders bidi at render time from `w:rtl`/`w:bidi`; if
  you also pre-reorder the string you get double-reversed (scrambled) text.
- **Numbers/punctuation in mixed lines** sort themselves out from the paragraph's RTL base
  direction. Don't hand-reorder `1,500.00 ש"ח`.
- **`cs` font slot, not just `ascii`.** Setting only `ascii`/`hAnsi` leaves Hebrew on
  Word's default complex-script font.
- **No nikud** in formal/military documents.
- **`.docx` files in this repo may be UTF-8 text, not OOXML** (per CLAUDE.md) — `cat`
  first before opening with python-docx.

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| Hebrew shows as boxes/`?` | Font lacks Hebrew or `cs` slot unset | Set `w:rFonts` `cs` to David/Heebo |
| Text left-aligned / LTR | Missing `w:bidi` / `w:rtl` | `set_paragraph_rtl` + `set_run_rtl` |
| Table columns not mirrored | Missing `w:bidiVisual` | `set_table_rtl(table)` |
| Numbers in wrong place | String pre-reordered, or LTR base | Don't `get_display` for DOCX |
