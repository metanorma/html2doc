# Phase 3: Paragraph and Run Conversion (HTML → OOXML Body)

## Goal
Convert the cleaned Nokogiri HTML document's body content (paragraphs, headings, inline formatting) into Uniword `Paragraph` and `Run` model objects with correct ISO style assignments.

## Background
After html2doc's `cleanup` pipeline, the document is a Nokogiri XML tree with:
- `<p>` elements with `class` attributes mapping to Word styles
- Inline formatting: `<strong>`, `<em>`, `<u>`, `<span style="...">`
- Headings (`<h1>`-`<h6>`) with style classes
- Bookmarks (`<a name="...">`) already inserted
- Math OOXML already embedded inline
- Lists already converted to `MsoListParagraphCxSp*` styled paragraphs

The challenge is walking this HTML tree and producing Uniword's `Paragraph > Run` model graph with:
- Correct style references (ISO style IDs from Phase 2)
- Correct run properties (bold, italic, underline, font, size, color)
- Proper nesting of formatting elements

## Tasks

- [ ] Implement `DocxConverter#convert_body(docxml)` — walks `<body>` children
  ```ruby
  def convert_body(docxml)
    body = Uniword::Wordprocessingml::Body.new
    docxml.xpath("//body/*").each do |element|
      case element.name
      when "p" then body.paragraphs << convert_paragraph(element)
      when "table" then body.tables << convert_table(element)
      when "div" then convert_div(element).each { |p| body.paragraphs << p }
      # ... other elements
      end
    end
    body
  end
  ```

- [ ] Implement `DocxConverter#convert_paragraph(element)` — maps `<p>` to `Paragraph`
  - Extract `class` attribute → resolve to ISO style ID via `class_to_style_map`
  - Extract `style` attribute → parse CSS inline styles into `ParagraphProperties`
  - Handle `align` attribute → `Justification` property
  - Extract `lang` attribute → `Language` property

- [ ] Implement `DocxConverter#convert_runs(element)` — walks inline children
  - TEXT nodes → `Run` with text
  - `<strong>`/`<b>` → `Run` with `Properties::Bold`
  - `<em>`/`<i>` → `Run` with `Properties::Italic`
  - `<u>` → `Run` with `Properties::Underline`
  - `<sub>`/`<sup>` → `Run` with `Properties::VerticalAlign`
  - `<span style="...">` → `Run` with parsed CSS properties (color, font-size, font-family)
  - `<a href="...">` → `Hyperlink` with `Run` children
  - `<a name="...">` → `BookmarkStart` + `BookmarkEnd`
  - Nested formatting: `<strong><em>text</em></strong>` → single Run with both Bold and Italic

- [ ] Implement CSS inline style parsing → `RunProperties`:
  - `font-weight: bold` → `Bold`
  - `font-style: italic` → `Italic`
  - `text-decoration: underline` → `Underline`
  - `color: #RRGGBB` → `Color`
  - `font-size: NNpt` → `FontSize` (convert pt to half-points)
  - `font-family: Name` → `RunFonts`
  - `mso-` prefixed properties → direct OOXML properties

- [ ] Handle special paragraph types:
  - List paragraphs (`MsoListParagraphCxSp*`) → assign `numId` and `ilvl` from stylesheet list styles
  - TOC paragraphs → `Sisluet1`-`Sisluet9` styles with `TabStop` leaders
  - Title/cover page paragraphs → `zzCover`, `zzSTDTitle` styles

- [ ] Handle section breaks:
  - `<br style="page-break-before:always">` or explicit section divs
  - Create `SectionProperties` with page size, margins, header/footer references
  - Map `@page` CSS to `PageSize`, `PageMargin` properties

## Key Files to Create/Modify
- `lib/html2doc/docx_converter.rb` — all conversion methods
- `lib/html2doc/css_parser.rb` — CSS inline style to OOXML properties

## Success Criteria
- Simple HTML paragraphs with formatting convert to Uniword `Paragraph` objects
- Style IDs correctly resolved from CSS classes
- Inline formatting (bold, italic, underline, color, font) correctly mapped
- Nested formatting preserved (no duplicated runs)
