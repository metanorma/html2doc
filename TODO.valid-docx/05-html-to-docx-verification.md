# Plan 05: HTML-to-DOCX Output Verification

## Goal
Verify that html2doc's DOCX output matches the expected output in spec/fixtures and passes Uniword validation. This is the end-to-end test that the library actually works for its primary purpose.

## Current html2doc DOCX pipeline

```
HTML input → Html2Doc.new.process(html)
  → to_xhtml (Nokogiri HTML → XHTML)
  → cleanup (lists, footnotes, math)
  → DocxConverter.convert(docxml)
    → extract_footnotes
    → convert_body (paragraphs, headings, tables, lists, math, images)
    → load_styles (ISO template from data/iso_styles.xml)
    → convert_numbering
    → parse_header_file (headers/footers from MS Word HTML)
    → apply_sections
    → assemble_package (Uniword::Docx::Package)
  → save_to_file
```

## Existing tests

### spec/docx_output_spec.rb
Tests that DOCX output:
- Contains required OOXML parts
- Has well-formed XML in all parts
- Has document body with paragraphs
- Handles headings, formatting, lists, tables, footnotes, math, images, bookmarks, hyperlinks

Uses `shared_examples "valid docx package"` for basic validity checks.

### spec/docx_converter_spec.rb
Unit tests for DocxConverter:
- Paragraph conversion and style mapping
- Heading mapping (h1→Heading1, etc.)
- Inline formatting (bold, italic, nested)
- Hyperlink conversion
- Table conversion
- Paragraph style parsing

### spec/style_loader_spec.rb
Tests ISO template loading:
- Styles, numbering, fonts, theme, settings

## Gaps

### G1. No baseline comparison
**Problem:** There's no reference DOCX to compare against. `spec/fixtures/iso-damd-fdis-sample.docx` exists but no test compares html2doc output to it.

**Solution:** For key HTML fixtures, generate a "golden" DOCX and store in `spec/fixtures/`. Tests verify output matches golden files (within tolerance for timestamps and random IDs).

### G2. No validation of generated DOCX
**Problem:** `docx_output_spec.rb` checks for required parts and well-formed XML, but doesn't run Uniword's validation pipeline.

**Solution:** After generating DOCX, run `Uniword::Validation::DocumentValidator.new.valid?(path)` and assert no errors.

### G3. Missing coverage for some features
The existing tests cover basics, but these are not tested for DOCX output:
- Multi-section headers/footers
- PAGE field codes in headers/footers
- Complex list numbering (nested lists, custom formats)
- Image embedding (binary data in word/media/)
- Math equations (OMML passthrough)
- Bookmark/hyperlink cross-references

**Solution:** Add test cases for each missing feature.

### G4. ISO template styles verification
**Problem:** No test verifies that the ISO template styles are correctly applied to generated content (e.g., that list paragraphs get MsoListParagraph styles, that headings have correct Finnish style names).

**Solution:** Add tests that check styleId values on generated paragraphs.

## Test plan

### Test 1: Basic DOCX generation
```ruby
it "generates a valid DOCX from simple HTML" do
  html = '<html><head><title>Test</title></head><body><p>Hello World</p></body></html>'
  entries = generate_docx(html)
  # Verify required parts
  # Run Uniword validation
  expect(Uniword::Validation::DocumentValidator.new.valid?(docx_path)).to be true
end
```

### Test 2: Heading hierarchy
```ruby
it "maps HTML headings to Word heading styles" do
  html = '<html><head><title>Test</title></head><body>
    <h1>Heading 1</h1><h2>Heading 2</h2><h3>Heading 3</h3>
  </body></html>'
  entries = generate_docx(html)
  doc = parse_xml(entries["word/document.xml"])
  h1 = doc.xpath("//w:p[w:pPr/w:pStyle/@w:val='Heading1']", "w" => W_NS)
  expect(h1.length).to eq(1)
end
```

### Test 3: Numbering preservation
```ruby
it "creates valid numbering for ordered lists" do
  html = '<html><head><title>Test</title></head><body>
    <ol><li>First</li><li>Second</li></ol>
  </body></html>'
  entries = generate_docx(html)
  numbering = parse_xml(entries["word/numbering.xml"])
  abstract_nums = numbering.xpath("//w:abstractNum", "w" => W_NS)
  expect(abstract_nums.length).to be > 0
  # Validate numbering structure
  expect(Uniword::Validation::DocumentValidator.new.valid?(docx_path)).to be true
end
```

### Test 4: Footnotes
```ruby
it "creates valid footnotes structure" do
  # Use HTML with footnote markup that html2doc expects
  entries = generate_docx(html_with_footnotes)
  expect(entries).to have_key("word/footnotes.xml")
  footnotes = parse_xml(entries["word/footnotes.xml"])
  # Verify separator entries
  expect(footnotes.xpath("//w:footnote[@w:type='separator']", "w" => W_NS).length).to be > 0
end
```

### Test 5: Table conversion
```ruby
it "creates valid table structure" do
  html = '<table><tr><td>A</td><td>B</td></tr></table>'
  entries = generate_docx(html)
  doc = parse_xml(entries["word/document.xml"])
  tables = doc.xpath("//w:tbl", "w" => W_NS)
  expect(tables.length).to eq(1)
end
```

### Test 6: Image embedding
```ruby
it "embeds images in word/media/" do
  # HTML with local image reference
  entries = generate_docx(html_with_image, imagedir: "spec/fixtures/images")
  media_files = entries.keys.select { |k| k.start_with?("word/media/") }
  expect(media_files.length).to be > 0
end
```

### Test 7: Validation pass on all generated DOCX
```ruby
RSpec.describe "DOCX output validation" do
  [simple_html, heading_html, list_html, table_html, footnote_html].each do |html|
    it "passes Uniword validation for #{html_description(html)}" do
      entries = generate_docx(html)
      path = entries[:path]
      validator = Uniword::Validation::DocumentValidator.new
      report = validator.validate(path)
      expect(report.valid?).to be true, "Validation errors: #{report.errors}"
    end
  end
end
```

## Implementation steps

1. **Create golden DOCX fixtures** for key test cases in `spec/fixtures/`
2. **Add Uniword validation** to `shared_examples "valid docx package"`
3. **Add feature-specific tests** for missing coverage (G3)
4. **Add ISO style verification tests** (G4)
5. **Run full DOCX output test suite** and fix any failures
6. **Add to CI** to prevent regressions

## Success criteria
- All DOCX output tests pass
- All generated DOCX files pass Uniword validation
- html2doc can generate a DOCX from an ISO HTML document that opens in Word without errors
