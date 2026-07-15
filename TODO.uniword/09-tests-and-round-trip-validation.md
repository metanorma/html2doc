# Phase 9: Testing and Round-Trip Validation

## Goal
Build comprehensive test coverage for the DOCX output path, validate output against the ISO fixture quality bar, and ensure round-trip fidelity (DOCX → Uniword → back to DOCX produces identical output).

## Background
The existing test suite (`spec/html2doc_spec.rb`) tests the MHT path. We need parallel tests for DOCX output. The `canon` gem (already added in spec/canon branch) provides semantic XML comparison for test assertions.

## Tasks

- [ ] Create `spec/docx_output_spec.rb` — end-to-end DOCX generation tests:
  - Simple document with paragraphs, formatting
  - Document with headings (h1-h6)
  - Document with bold, italic, underline, colored text
  - Document with bullet lists (nested)
  - Document with numbered lists (with start numbers)
  - Document with footnotes
  - Document with images (local, various formats)
  - Document with math equations (inline and block)
  - Document with tables
  - Document with bookmarks/anchors
  - Multi-section document with headers/footers

- [ ] Create `spec/docx_converter_spec.rb` — unit tests for conversion methods:
  - `convert_paragraph` — style mapping, run extraction
  - `convert_runs` — inline formatting preservation
  - `convert_list_paragraph` — numbering assignment
  - `convert_footnotes` — footnote extraction and reference creation
  - `convert_images` — image embedding
  - `convert_math` — OMML passthrough

- [ ] Create `spec/style_loader_spec.rb` — style loading tests:
  - ISO styles load without errors
  - Class-to-style mapping resolves correctly
  - Numbering definitions load correctly
  - Font table and theme load correctly

- [ ] Validate DOCX output quality:
  - Each generated DOCX must be a valid ZIP archive
  - All required parts present (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, etc.)
  - XML well-formed in all parts
  - Relationships are internally consistent (every rId has a target)
  - Content types cover all parts

- [ ] Use Uniword to round-trip test:
  ```ruby
  it "round-trips through Uniword" do
    # Generate DOCX via html2doc
    Html2Doc.new(filename: "test", output_format: :docx).process(html)

    # Load via Uniword
    pkg = Uniword::Ooxml::DocxPackage.from_file("test.docx")

    # Save again
    pkg.save("test_roundtrip.docx")

    # Compare semantically using canon
    expect("test.docx").to canon_match("test_roundtrip.docx")
  end
  ```

- [ ] Compare with ISO fixture quality:
  - Generate a DOCX from ISO 690:2021 HTML source
  - Compare the generated DOCX structure against the reference ISO DOCX fixture
  - Validate that styles used in the document are defined in styles.xml
  - Validate that numbering references have corresponding definitions

- [ ] Performance benchmarks:
  - Measure DOCX generation time vs MHT generation time
  - Profile memory usage for large documents
  - Identify bottlenecks (likely Nokogiri → OOXML model conversion)

## Key Files to Create/Modify
- `spec/docx_output_spec.rb` — end-to-end DOCX tests
- `spec/docx_converter_spec.rb` — unit conversion tests
- `spec/style_loader_spec.rb` — style loading tests
- `spec/html2doc_spec.rb` — add DOCX variants of existing MHT tests

## Success Criteria
- All new tests pass
- DOCX output validates as well-formed OOXML
- Round-trip through Uniword produces semantically equivalent output
- No regressions in existing MHT test suite
