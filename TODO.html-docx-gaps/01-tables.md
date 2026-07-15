# 1. Tables not converted in DOCX output

## Problem

The HTML `rice.html` contains 2 `<table>` elements (a large defects table and an interlaboratory results table). The generated DOCX contains **0 `<w:tbl>` elements**. Table content is either dropped or rendered as flat paragraphs.

## Root cause

`DocxConverter#convert_div` (line 150-152) handles `<table>` children by calling `convert_paragraph(child)`, which just treats the table as a generic inline container — it does NOT call `convert_table`:

```ruby
# docx_converter.rb:150-152
when "table"
  # Tables are handled separately - would need to return both
  paragraphs << convert_paragraph(child)
```

`convert_body` (line 83) DOES call `convert_table` for top-level tables, but in `rice.html` the tables are nested inside `<div>` elements, so they never hit that path.

Additionally, `convert_div` only returns paragraphs (Array of Paragraph), so even if it called `convert_table`, the return type is incompatible — tables would need to be returned alongside paragraphs.

## Required changes

1. **`convert_div` must return mixed content** — both Paragraphs and Tables. Change the return type to a flat array that can contain both, or return a structured result. The caller `convert_body` needs to handle this mixed return.

2. **`convert_div` must call `convert_table`** for `<table>` children instead of `convert_paragraph`.

3. **Table cell content** needs special handling — cells contain `<p>`, `<span class="stem">`, `<a>`, and nested `<div class="Note">`. The current `convert_table` method handles `<p>` but not nested divs or inline math in cells.

4. **Table properties** need mapping from CSS inline styles:
   - `border-*` styles → `w:tblBorders` / `w:tcBorders`
   - `border-collapse` → table properties
   - `rowspan` / `colspan` → `w:gridSpan` / `w:vMerge`
   - `align` attribute → `w:jc`
   - `class="MsoISOTable"` → style reference (map via `StyleLoader`)
   - `width` → `w:tblW` / `w:tcW`

## Scope in rice.html

- **Table 1** (line 1223): 16 rows x 5 cols with `rowspan`, `colspan`, thead/tbody/tfoot, border styles, table footnotes, and inline stem math in cells.
- **Table D-1** (line 1519): interlaboratory test results with similar complexity.

## Acceptance criteria

- Both tables in rice.html produce valid `<w:tbl>` elements in the DOCX
- Cell content (text, bold, hyperlinks, stem math) is preserved
- Rowspan/colspan produce correct `gridSpan`/`vMerge`
- Borders and alignment from inline CSS are mapped
