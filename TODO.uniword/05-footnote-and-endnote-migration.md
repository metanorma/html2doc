# Phase 5: Footnote and Endnote Migration

## Goal
Convert html2doc's footnote processing output into Uniword's native OOXML footnote/endnote model objects (`Footnote`, `Footnotes`, `FootnoteReference`, etc.).

## Background
html2doc's `notes.rb` already:
- Detects footnote links (`<a class="Footnote">` or `<a epub:type="footnote">`)
- Collects footnote content from `<aside>` / `<div>` targets
- Creates `<div style="mso-element:footnote">` containers with auto-numbered references
- Replaces footnote body links with `MsoFootnoteText` styled content

For DOCX output, the architecture is fundamentally different:
- Footnotes are a separate XML part (`word/footnotes.xml`), not embedded in document body
- References are `<w:footnoteReference w:id="N"/>` inside `<w:run>` elements
- Footnote separators are system footnotes (id -1, 0)
- The ISO fixture has 34 footnotes with hyperlink references

## Tasks

- [ ] Decide approach: **intercept before** `notes.rb` or **convert after**
  - **Recommended: Intercept before** — for DOCX output, skip `notes.rb` entirely and convert footnotes directly to Uniword model objects
  - Rationale: The MHT approach embeds footnotes in the document body; DOCX keeps them separate. Trying to reverse-engineer MHT footnote markup into OOXML is harder than starting from the original HTML

- [ ] Create `DocxConverter#extract_footnotes(docxml)` — runs before body conversion
  ```ruby
  def extract_footnotes(docxml)
    footnotes = Uniword::Wordprocessingml::Footnotes.new
    idx = 1  # system footnotes use -1 and 0
    docxml.xpath("//a[@epub:type='footnote' or @class='Footnote']").each do |link|
      href = link["href"].sub(/^#/, "")
      target = docxml.at("//*[@id='#{href}' or @name='#{href}']")
      next unless target

      # Create Footnote model
      fn = Uniword::Wordprocessingml::Footnote.new(id: idx)
      fn.paragraphs = convert_footnote_content(target)
      footnotes.add_footnote(fn)

      # Replace link with reference marker
      link.replace(create_footnote_reference(idx))
      idx += 1
    end
    footnotes
  end
  ```

- [ ] Implement `DocxConverter#create_footnote_reference(idx)`:
  - Returns a Nokogiri placeholder element that will be converted to a `Run` with `FootnoteReference` during body conversion
  - Use a recognizable marker: `<span data-footnote-ref="#{idx}">`

- [ ] Implement `DocxConverter#convert_footnote_content(target)`:
  - Convert footnote `<aside>`/`<div>` children to `Paragraph` objects
  - Apply `MsoFootnoteText` / ISO `Alaviitteenteksti` style
  - Handle footnote auto-number markers (`MsoFootnoteReference` spans)
  - Preserve hyperlinks within footnotes (they'll need entries in `document.xml.rels`)

- [ ] Handle footnote hyperlink relationships:
  - Footnotes in the ISO fixture have 30+ external hyperlinks
  - These go into `word/_rels/footnotes.xml.rels` (separate from document rels)
  - Create `Relationships` object for footnote relationships

- [ ] Add system footnotes:
  - Separator (id: -1): `<w:p><w:r><w:separator/></w:r></w:p>`
  - Continuation separator (id: 0): similar with `<w:continuationSeparator/>`

- [ ] Add endnote support (same pattern, `word/endnotes.xml`):
  - If `epub:type="endnote"` is detected, route to endnotes instead

## Key Files to Create/Modify
- `lib/html2doc/docx_converter.rb` — footnote extraction and conversion
- `lib/html2doc/base.rb` — skip `notes.rb` for DOCX output path, call DocxConverter instead

## Success Criteria
- HTML footnotes converted to OOXML `Footnote` model objects
- Footnote references (`w:footnoteReference`) correctly inserted inline
- Footnote content preserves formatting (bold, italic, hyperlinks)
- System footnote separators included
- External hyperlinks in footnotes have relationship entries
