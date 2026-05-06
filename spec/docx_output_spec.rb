require "spec_helper"
require "zip"
require "html2doc"
require "uniword/validation/rules"

RSpec.describe "DOCX output" do
  def generate_docx(html, options = {})
    filename = options.delete(:filename) || "/tmp/docx_test_output"
    stylesheet = options.delete(:stylesheet)
    header_file = options.delete(:header_file)

    Html2Doc.new(
      filename: filename,
      stylesheet: stylesheet,
      header_file: header_file,
      output_format: :docx,
    ).process(html)

    path = "#{filename}.docx"
    expect(File.exist?(path)).to be true

    entries = {}
    Zip::File.open(path) do |z|
      z.entries.each do |e|
        entries[e.name] = z.read(e.name).force_encoding("UTF-8")
      end
    end
    entries[:path] = path
    entries
  end

  def docx_path
    entries[:path]
  end

  def parse_xml(xml)
    Nokogiri::XML(xml) { |config| config.strict }
  end

  shared_examples "valid docx package" do
    it "contains required OOXML parts" do
      expect(entries).to have_key("[Content_Types].xml")
      expect(entries).to have_key("_rels/.rels")
      expect(entries).to have_key("word/document.xml")
      expect(entries).to have_key("word/styles.xml")
      expect(entries).to have_key("word/fontTable.xml")
      expect(entries).to have_key("word/settings.xml")
      expect(entries).to have_key("word/_rels/document.xml.rels")
    end

    it "has well-formed XML in all parts" do
      entries.each do |name, content|
        next unless name.end_with?(".xml")
        doc = parse_xml(content)
        expect(doc.errors).to be_empty, "XML errors in #{name}: #{doc.errors.map(&:to_s).join(', ')}"
      end
    end

    it "has a document body with paragraphs" do
      doc = parse_xml(entries["word/document.xml"])
      paras = doc.xpath("//w:p", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(paras.size).to be > 0
    end

    it "passes Uniword DOC-100..DOC-109 validation rules" do
      path = docx_path
      context = Uniword::Validation::Rules::DocumentContext.new(path)
      issues = Uniword::Validation::Rules::Registry.all.flat_map do |rule|
        rule.applicable?(context) ? rule.check(context) : []
      end
      context.close

      errors = issues.select { |i| i.severity == "error" }
      expect(errors).to be_empty,
        "Validation errors:\n#{errors.map { |e| "  #{e.code}: #{e.message}" }.join("\n")}"
    end
  end

  describe "simple document" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">Hello world</p>
        <p class="MsoNormal">Second paragraph</p>
      </body></html>')
    end

    include_examples "valid docx package"

    it "contains paragraph text" do
      doc = parse_xml(entries["word/document.xml"])
      texts = doc.xpath("//w:t", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      all_text = texts.map(&:text).join
      expect(all_text).to include("Hello world")
      expect(all_text).to include("Second paragraph")
    end

    it "applies Normal style" do
      doc = parse_xml(entries["word/document.xml"])
      styles = doc.xpath("//w:pStyle/@w:val", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(styles.map(&:value)).to all(eq("Normal"))
    end
  end

  describe "headings" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <h1>Heading 1</h1>
        <h2>Heading 2</h2>
        <h3>Heading 3</h3>
        <h4>Heading 4</h4>
      </body></html>')
    end

    include_examples "valid docx package"

    it "maps h1-h4 to ISO heading styles" do
      doc = parse_xml(entries["word/document.xml"])
      styles = doc.xpath("//w:pStyle/@w:val", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      values = styles.map(&:value)
      expect(values).to include("Heading1", "Heading2", "Heading3", "Heading4")
    end
  end

  describe "inline formatting" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">
          Normal <b>bold</b> <i>italic</i> <u>underline</u> <b><i>bold italic</i></b>
        </p>
      </body></html>')
    end

    include_examples "valid docx package"

    it "creates bold runs" do
      doc = parse_xml(entries["word/document.xml"])
      bold = doc.xpath("//w:b", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(bold.size).to be >= 2  # bold + bold italic
    end

    it "creates italic runs" do
      doc = parse_xml(entries["word/document.xml"])
      italic = doc.xpath("//w:i", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(italic.size).to be >= 2  # italic + bold italic
    end

    it "creates underline runs" do
      doc = parse_xml(entries["word/document.xml"])
      underline = doc.xpath("//w:u", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(underline.size).to be >= 1
    end
  end

  describe "lists" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <p class="MsoListParagraphCxSpFirst" style="mso-list:l0 level1 lfo1">Item 1</p>
        <p class="MsoListParagraphCxSpMiddle" style="mso-list:l0 level1 lfo1">Item 2</p>
        <p class="MsoListParagraphCxSpLast" style="mso-list:l0 level1 lfo1">Item 3</p>
      </body></html>')
    end

    include_examples "valid docx package"

    it "creates numbering properties" do
      doc = parse_xml(entries["word/document.xml"])
      num_pr = doc.xpath("//w:numPr", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(num_pr.size).to eq(3)
    end
  end

  describe "bookmarks" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">Before <a name="anchor1"></a>after</p>
      </body></html>')
    end

    include_examples "valid docx package"

    it "does not create Word bookmarks from HTML anchors" do
      doc = parse_xml(entries["word/document.xml"])
      starts = doc.xpath("//w:bookmarkStart", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      # Reference DOCX has minimal bookmarks (1 total); HTML anchors don't become Word bookmarks
      expect(starts.size).to eq(0)
    end
  end

  describe "tables" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <table>
          <tr><th>Header 1</th><th>Header 2</th></tr>
          <tr><td>Cell 1</td><td>Cell 2</td></tr>
        </table>
      </body></html>')
    end

    include_examples "valid docx package"

    it "creates table structure" do
      doc = parse_xml(entries["word/document.xml"])
      tables = doc.xpath("//w:tbl", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(tables.size).to eq(1)
      rows = doc.xpath("//w:tr", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(rows.size).to eq(2)
      cells = doc.xpath("//w:tc", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(cells.size).to eq(4)
    end
  end

  describe "with headers and footers" do
    let(:entries) do
      generate_docx(
        '<html><head><title>Test</title></head><body>
          <p class="MsoNormal">Content</p>
        </body></html>',
        header_file: File.expand_path("header.html", __dir__)
      )
    end

    include_examples "valid docx package"

    it "includes header and footer files" do
      headers = entries.keys.select { |k| k.match?(/header\d+\.xml/) }
      footers = entries.keys.select { |k| k.match?(/footer\d+\.xml/) }
      expect(headers.size).to be > 0
      expect(footers.size).to be > 0
    end

    it "has section properties with header/footer references" do
      doc = parse_xml(entries["word/document.xml"])
      hdr_refs = doc.xpath("//w:headerReference", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      ftr_refs = doc.xpath("//w:footerReference", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(hdr_refs.size).to be > 0
      expect(ftr_refs.size).to be > 0
    end

    it "includes PAGE field code in footer" do
      footer_key = entries.keys.find { |k| k.match?(/footer\d+\.xml/) }
      expect(footer_key).not_to be_nil
      doc = parse_xml(entries[footer_key])
      fld_chars = doc.xpath("//w:fldChar", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(fld_chars.size).to be > 0
    end
  end

  describe "alignment" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <p class="MsoNormal" style="text-align:center">Centered text</p>
        <p class="MsoNormal" style="text-align:right">Right-aligned text</p>
      </body></html>')
    end

    include_examples "valid docx package"

    it "sets center alignment" do
      doc = parse_xml(entries["word/document.xml"])
      jc = doc.xpath("//w:jc/@w:val", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
      expect(jc.map(&:value)).to include("center")
    end
  end

  describe "without header file" do
    let(:entries) do
      generate_docx('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">No header</p>
      </body></html>')
    end

    include_examples "valid docx package"

    it "does not include header/footer files" do
      headers = entries.keys.select { |k| k.match?(/header\d+\.xml/) }
      footers = entries.keys.select { |k| k.match?(/footer\d+\.xml/) }
      expect(headers).to be_empty
      expect(footers).to be_empty
    end
  end

  describe "reference fixture validation" do
    it "spec/fixtures/iso-damd-fdis-sample.docx passes Uniword validation" do
      path = File.expand_path("fixtures/iso-damd-fdis-sample.docx", __dir__)
      skip "Reference fixture not found" unless File.exist?(path)

      ctx = Uniword::Validation::Rules::DocumentContext.new(path)
      issues = Uniword::Validation::Rules::Registry.all.flat_map do |rule|
        rule.applicable?(ctx) ? rule.check(ctx) : []
      end
      ctx.close

      errors = issues.select { |i| i.severity == "error" }
      expect(errors).to be_empty,
        "Reference DOCX validation errors:\n#{errors.map { |e| "  #{e.code}: #{e.message}" }.join("\n")}"
    end
  end
end
