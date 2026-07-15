# 4. Page breaks not generated in DOCX output

## Problem

The HTML `rice.html` uses CSS `page-break-before:always` and `<br clear="all" style="page-break-before:always">` to create page breaks between sections. The DOCX output contains **17 `<w:br/>` elements** (line breaks) but **0 `<w:br w:type="page"/>`** page breaks.

## Root cause

### 1. `<br>` at body level is not handled

In `rice.html`, page breaks appear as direct children of `<body>`:
```html
<br clear="all" style="mso-special-character:line-break;page-break-before:always" />
```
and:
```html
<br clear="all" class="section" />
```

`DocxConverter#convert_body` (line 75-97) has no `when "br"` case — `<br>` elements at body level fall through to the `else` branch and get wrapped as generic paragraphs via `convert_paragraph(node)`. Inside `convert_paragraph`, `convert_inline_content` iterates children of the `<br>` element (which has none), producing an empty paragraph.

### 2. `create_break_run` logic is correct but unreachable

`create_break_run` (line 405-413) does check for page-break:
```ruby
def create_break_run(element)
  run = Uniword::Wordprocessingml::Run.new
  if element["style"]&.include?("page-break")
    run.break = Uniword::Wordprocessingml::Break.new(type: "page")
  else
    run.break = Uniword::Wordprocessingml::Break.new
  end
  run
end
```

This would work for `<br>` inside a paragraph, but `<br>` at body level never reaches this code.

### 3. CSS `page-break-before` on headings/divs not handled

CSS classes like `ForewordTitle`, `IntroTitle`, `ANNEX` have `page-break-before:always` in their styles. This needs to be detected when converting those elements and translated to either:
- A `<w:br w:type="page"/>` run at the start of the paragraph, or
- A section break with appropriate properties

## Required changes

1. **Handle `<br>` at body level** in `convert_body` — add a `when "br"` case that creates a paragraph containing a page break run:
   ```ruby
   when "br"
     para = Uniword::Wordprocessingml::Paragraph.new
     para.runs << create_break_run(node)
     body.paragraphs << para
   ```

2. **Detect page-break CSS** on block elements — when converting `<h1>`, `<div>`, `<p>` etc., check if the element's CSS class or inline style includes `page-break-before:always`. If so, prepend a page break run to the paragraph.

3. **Handle `<br clear="all" class="section">`** — this is a section break marker. In multi-section documents, this should trigger a new `w:sectPr` on the preceding paragraph (related to item 5).

## Scope in rice.html

- `<br clear="all" style="...page-break-before:always" />` between Foreword/Introduction/body (lines 1101, 1123, 1138)
- `<br clear="all" class="section" />` between WordSection2 and WordSection3 (line 761)
- CSS `page-break-before:always` on classes: `ForewordTitle`, `IntroTitle`, `ANNEX`, `BiblioTitle`, and others

## Acceptance criteria

- Page breaks between Foreword, Introduction, and body render as `<w:br w:type="page"/>` in the DOCX
- CSS `page-break-before:always` on styled elements is honored
- Section markers produce proper section breaks (when multi-section support is complete)
