# Plan 04: Round-Trip Test Suite

## Goal
Test that Uniword can load → save any DOCX file without introducing Word errors. Build a systematic test suite covering diverse document types.

## Test methodology

### Round-trip test pattern
```ruby
RSpec.describe "DOCX round-trip" do
  def round_trip_succeeds?(input_path)
    # 1. Load via Uniword
    pkg = Uniword::Docx::Package.from_file(input_path)

    # 2. Save to temp file
    output_path = "/tmp/roundtrip_#{File.basename(input_path)}"
    pkg.save(output_path)

    # 3. Run Uniword validation
    validator = Uniword::Validation::DocumentValidator.new
    report = validator.validate(output_path)
    return false unless report.valid?

    # 4. Optional: compare XML parts for content preservation
    true
  end
end
```

### Swap test pattern (for isolating failures)
When a round-trip fails, isolate which part causes the failure:
1. Take the working repaired file as base
2. Swap in one part from the broken file
3. If swap fails → that part is the problem
4. Diff broken vs repaired for that part

### Content preservation test pattern
```ruby
def preserves_content?(input_path, output_path)
  Zip::File.open(input_path) do |input|
    Zip::File.open(output_path) do |output|
      %w[word/document.xml word/styles.xml word/numbering.xml
         word/settings.xml word/fontTable.xml].each do |part|
        next unless input.find_entry(part) && output.find_entry(part)
        input_xml = Nokogiri::XML(input.read(part))
        output_xml = Nokogiri::XML(output.read(part))
        # Compare canonical XML (ignoring namespace ordering, whitespace)
        expect(canonicalize(output_xml)).to eq(canonicalize(input_xml))
      end
    end
  end
end
```

## Test matrix

### Tier 1: Real-world ISO documents (existing in Uniword repo)
These are the most important — they represent actual Metanorma output.

| File | Source | Key features |
|------|--------|-------------|
| ISO 8601-1:2019/Amd1 | Uniword spec fixtures | Complex structure, footnotes, tables |
| ISO 690:2021 | Uniword spec fixtures | Standard document |
| word-template-paper-with-cover-and-toc | Uniword spec fixtures | Multi-section, headers/footers |
| word-template-apa-style-paper | Uniword spec fixtures | APA styling |

### Tier 2: Rice document (html2doc's test document)
| File | Source | Key features |
|------|--------|-------------|
| rice.docx | html2doc spec/examples | Full ISO document with numbering, footnotes, images |

### Tier 3: Synthetic minimal documents
Test edge cases with minimal documents:

| Test case | What it tests |
|-----------|--------------|
| Empty document | Just one paragraph |
| Document with footnotes | footnotePr + footnotes.xml |
| Document with endnotes | endnotePr + endnotes.xml |
| Document with images | blip references to media |
| Document with tables | complex table with merges |
| Document with math | oMath/oMathPara |
| Document with bookmarks | bookmarkStart/bookmarkEnd |
| Document with hyperlinks | internal + external |
| Multi-section document | Multiple sectPr with different headers |
| Document with numbering | ol/ul lists, multiple abstractNums |

### Tier 4: html2doc-generated DOCX files
These test the end-to-end html2doc pipeline:

| Input HTML | Expected features |
|------------|------------------|
| Basic paragraph | Simple paragraph output |
| Heading hierarchy | h1-h6 mapped to Heading1-Heading6 |
| Bold/italic/underline | Run formatting |
| Ordered list | Numbering instance |
| Unordered list | Bullet numbering |
| Table | Word table with grid |
| Footnotes | footnotes.xml with references |
| Image | Image in word/media/ |
| Math (OMML) | oMath passthrough |
| Multi-section | Multiple sectPr |

## Existing Uniword round-trip tests
Uniword already has `spec/integration/round_trip_spec.rb` (20 examples, all passing) and `spec/integration/docx_roundtrip_spec.rb`. Leverage and extend these.

## Implementation steps

1. **Create shared round-trip helper** in Uniword spec helpers
2. **Add Tier 1 tests** — extend existing round_trip_spec.rb with validation checks
3. **Add rice.docx round-trip test** in html2doc specs
4. **Create synthetic test documents** using Uniword Builder API
5. **Add html2doc generation tests** using existing docx_output_spec.rb pattern
6. **Add content preservation assertions** for critical parts (numbering, styles)
7. **Run validation after each round-trip** using the new validation rules from Plan 02

## Success criteria
- All Tier 1-3 round-trips pass without validation errors
- rice.docx round-trip opens in Word without errors or repair
- html2doc-generated DOCX files pass Uniword validation
- Content preservation verified for critical parts
