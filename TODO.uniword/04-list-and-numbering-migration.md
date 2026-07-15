# Phase 4: List and Numbering Migration

## Goal
Convert html2doc's list processing output (paragraphs with `MsoListParagraphCxSp*` styles and `mso-list:` inline styles) into Uniword's native OOXML numbering system (`NumberingConfiguration` with `abstractNum`/`num` pairs, paragraphs with `numId`/`ilvl` properties).

## Background
html2doc already converts HTML `<ul>`/`<ol>` lists in its `cleanup` pipeline:
- `lists.rb` adds `mso-list:levelN lfoN` inline styles to `<li>` elements
- `list2para` converts `<li>` to `<p class="MsoListParagraphCxSpFirst/Middle/Last">`
- Ordered list start numbers handled by dynamically generating `@list` CSS rules

For DOCX output, we need a different approach:
- OOXML uses `<w:numPr>` in paragraph properties with `<w:numId>` and `<w:ilvl>`
- Numbering definitions are in `word/numbering.xml` as `abstractNum` + `num` pairs
- The ISO fixture has 13 numbering definitions covering decimal, bullet, and annex numbering

## Tasks

- [ ] Decide approach: **intercept before** `lists.rb` processing or **convert after**
  - **Recommended: Convert after** — let html2doc's existing list processing run, then translate the `mso-list:` CSS properties into OOXML numbering references
  - This avoids duplicating html2doc's complex list nesting logic

- [ ] Implement `DocxConverter#convert_list_paragraph(element)`:
  ```ruby
  def convert_list_paragraph(element)
    para = convert_paragraph(element)
    style = element["style"] || ""
    if mso_list = style.match(/mso-list:(\w+)\s+level(\d+)\s+lfo(\d+)/)
      liststyle, level, lfo = mso_list.captures
      para_props = para.properties || ParagraphProperties.new
      para_props.num_id = resolve_num_id(liststyle, lfo)
      para_props.ilvl = level.to_i - 1  # 0-indexed in OOXML
      para.properties = para_props
    end
    para
  end
  ```

- [ ] Build `NumberingConfiguration` from html2doc's parsed stylesheet list styles:
  - `parse_stylesheet_line_styles` in `lists.rb` already extracts `@list` definitions
  - Map each `@list` CSS definition to an OOXML `abstractNum`
  - Map each `lfo` reference to an OOXML `num` instance
  - Handle `mso-level-start-at` for ordered list start numbers

- [ ] Merge ISO numbering definitions with html2doc's dynamic definitions:
  - Start with ISO's 13 numbering definitions as the base
  - Add any additional definitions from the document's stylesheets
  - Ensure `numId` values don't collide (use ISO's max+1 as starting offset)

- [ ] Map html2doc bullet/number formats to OOXML `numFmt` values:
  | html2doc format | OOXML numFmt |
  |---|---|
  | bullet (Symbol font) | `bullet` |
  | decimal `%1.` | `decimal` |
  | lowerLetter `%1)` | `lowerLetter` |
  | lowerRoman `%i.` | `lowerRoman` |
  | upperLetter `%1)` | `upperLetter` |

- [ ] Handle list indentation:
  - html2doc uses CSS margin/padding on `<p>` elements
  - OOXML uses `w:ind` with `w:left`, `w:hanging`, `w:firstLine` in twips
  - Convert from CSS px/pt to OOXML twips (1pt = 20 twips)

## Key Files to Create/Modify
- `lib/html2doc/docx_converter.rb` — list paragraph conversion
- `lib/html2doc/numbering_converter.rb` — CSS list styles → OOXML numbering
- `lib/html2doc/lists.rb` — minor: export parsed list style data for DOCX path

## Success Criteria
- Nested ordered and unordered lists convert to OOXML numbering
- List indentation preserved
- Start numbers for ordered lists preserved
- Bullet characters match ISO template's Symbol font bullets
- Mixed nesting (ordered within unordered, etc.) works correctly
