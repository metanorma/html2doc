# 5. Multi-section support: headers, footers, and section properties

## Problem

The HTML `rice.html` uses 3 `WordSection` divs (WordSection1 = cover, WordSection2 = copyright/TOC, WordSection3 = body). The DOCX output has **only 1 `w:sectPr`** (on the body element) with default margins and **no header/footer references**.

## Root cause

### 1. Section boundaries not detected from HTML

`DocxConverter#apply_sections` (line 1016-1056) only activates when `parse_header_file` returns section data. Without a header file, no sections are created. The method has a TODO comment:

```ruby
# TODO: use landscape divs or explicit page breaks to determine section boundaries.
```

The code currently distributes paragraphs evenly across sections, which is wrong — sections should be determined by `<div class="WordSection1/2/3">` boundaries and `<br class="section">` markers.

### 2. No header file passed in test invocation

The test invocation `Html2Doc.new(output_format: :docx, filename: ..., imagedir: ...)` doesn't pass `header_file:`. The ISO header.html contains the header/footer definitions per section (eh1/h1/ef1/f1 for section 1, etc.).

### 3. Page size/margins not extracted from CSS

`build_section_properties` (line 1058-1083) hardcodes A4 with 1-inch margins. The actual page dimensions are defined in the CSS `@page WordSection1/2/3` rules and should be parsed from the stylesheet, similar to how `locate_landscape` and `page_dimensions` work in `mime.rb`.

## Required changes

1. **Detect WordSection boundaries** from `<div class="WordSection1">`, `<div class="WordSection2">`, etc. in the HTML body. These correspond to document sections.

2. **Map section breaks** — `<br clear="all" class="section" />` and `<br style="page-break-before:always" />` between WordSection divs should trigger a `w:sectPr` on the last paragraph of the preceding section.

3. **Parse CSS `@page` rules** — extract page size and margins from `@page WordSection1/2/3` in the stylesheet. The existing `page_dimensions` and `find_page_size` methods in `mime.rb` already do this; refactor or reuse them.

4. **Wire header/footer references** — when a `header_file` is provided, the existing `parse_header_file`, `build_header_footer_parts`, and `wire_header_footer_refs` methods should work. Test with the ISO `header.html`.

5. **Body-level sectPr** — the last section's properties go on `<w:body><w:sectPr>`. Intermediate sections go on the last `<w:p><w:pPr><w:sectPr>` of each section. Verify Uniword supports this correctly.

6. **Page numbering** — `build_section_properties` sets `pgNumType start="1"` for section index >= 2 (body). Verify this matches the legacy MHT output.

## Scope in rice.html

- **WordSection1** (line 690): Cover page — different page size/margins, no header/footer on first page
- **WordSection2** (line 762): Copyright + TOC — has header/footer, page numbering starts at Roman numerals
- **WordSection3** (line 1141): Body content — has header/footer, page numbering restarts at Arabic "1"

Each section has distinct `@page` rules in the CSS with different sizes and margins.

## Acceptance criteria

- 3 sections in the DOCX with distinct `w:sectPr` elements
- Cover page (WordSection1) has no header/footer on first page (`w:titlePg`)
- Copyright/TOC section (WordSection2) has appropriate header/footer
- Body section (WordSection3) has header/footer with PAGE number field
- Page size and margins match the CSS `@page` rules
- Page numbering restarts at "1" for the body section
