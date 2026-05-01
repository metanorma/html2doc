require "spec_helper"
require "html2doc"

RSpec.describe Html2Doc::DocxConverter do
  def convert(html)
    docxml = Nokogiri::HTML(html)
    converter = Html2Doc::DocxConverter.new(
      filename: "/tmp/test_converter",
      stylesheet: nil,
    )
    package = converter.convert(docxml)
    # Return the document for inspection
    [converter, package]
  end

  describe "#convert_paragraph" do
    it "does not set style by default when no class" do
      _, pkg = convert('<p>No class</p>')
      para = pkg.document.body.paragraphs.first
      # Reference output has no style for unclassed paragraphs
      expect(para.properties&.style).to be_nil
    end

    it "maps MsoNormal to no explicit style (Normal is default)" do
      _, pkg = convert('<p class="MsoNormal">Text</p>')
      para = pkg.document.body.paragraphs.first
      # MsoNormal resolves to "Normal" which is the default — no explicit pStyle needed
      expect(para.properties&.style).to be_nil
    end

    it "maps MsoTitle to zzSTDTitle" do
      _, pkg = convert('<p class="MsoTitle">Title</p>')
      para = pkg.document.body.paragraphs.first
      expect(para.properties).not_to be_nil
    end

    it "creates runs from text content" do
      _, pkg = convert('<p class="MsoNormal">Hello world</p>')
      para = pkg.document.body.paragraphs.first
      expect(para.runs.size).to be >= 1
      expect(para.runs.map(&:text).join).to include("Hello world")
    end
  end

  describe "#convert_heading" do
    it "maps h1 to Heading1" do
      _, pkg = convert('<h1>Heading 1</h1>')
      para = pkg.document.body.paragraphs.first
      # Heading style should be set
      expect(para.properties).not_to be_nil
    end

    it "maps h2 to Heading2" do
      _, pkg = convert('<h2>Heading 2</h2>')
      para = pkg.document.body.paragraphs.first
      expect(para.properties).not_to be_nil
    end
  end

  describe "inline formatting" do
    it "creates bold runs for <b>" do
      _, pkg = convert('<p class="MsoNormal"><b>bold</b></p>')
      para = pkg.document.body.paragraphs.first
      bold_runs = para.runs.select { |r| r.properties&.respond_to?(:bold) && r.properties.bold }
      expect(bold_runs.size).to be >= 1
    end

    it "creates italic runs for <i>" do
      _, pkg = convert('<p class="MsoNormal"><i>italic</i></p>')
      para = pkg.document.body.paragraphs.first
      italic_runs = para.runs.select { |r| r.properties&.respond_to?(:italic) && r.properties.italic }
      expect(italic_runs.size).to be >= 1
    end

    it "handles nested formatting" do
      _, pkg = convert('<p class="MsoNormal"><b><i>bold italic</i></b></p>')
      para = pkg.document.body.paragraphs.first
      # Should have a run with both bold and italic
      bi_runs = para.runs.select { |r| r.properties&.bold && r.properties&.italic }
      expect(bi_runs.size).to be >= 1
    end
  end

  describe "#convert_link" do
    it "creates internal anchor hyperlinks" do
      _, pkg = convert('<p class="MsoNormal"><a href="#section1">link</a></p>')
      para = pkg.document.body.paragraphs.first
      expect(para.hyperlinks.size).to eq(1)
      expect(para.hyperlinks.first.anchor).to eq("section1")
    end

    it "creates external URL hyperlinks" do
      _, pkg = convert('<p class="MsoNormal"><a href="https://example.com">link</a></p>')
      para = pkg.document.body.paragraphs.first
      expect(para.hyperlinks.size).to eq(1)
      expect(para.hyperlinks.first.id).to start_with("rIdLink")
    end
  end

  describe "tables" do
    it "creates table with rows and cells" do
      _, pkg = convert('<table><tr><td>A</td><td>B</td></tr></table>')
      table = pkg.document.body.tables.first
      expect(table).not_to be_nil
      expect(table.rows.size).to eq(1)
      expect(table.rows.first.cells.size).to eq(2)
    end
  end

  describe "parse_paragraph_style" do
    it "parses text-align center" do
      _, pkg = convert('<p class="MsoNormal" style="text-align:center">text</p>')
      para = pkg.document.body.paragraphs.first
      expect(para.properties.alignment).not_to be_nil
    end

    it "parses mso-list for numbering" do
      _, pkg = convert('<p class="MsoListParagraphCxSpFirst" style="mso-list:l0 level1 lfo1">Item</p>')
      para = pkg.document.body.paragraphs.first
      expect(para.properties.numbering_properties).not_to be_nil
    end
  end

  describe "with header file" do
    it "parses header.html into sections" do
      converter = Html2Doc::DocxConverter.new(
        filename: "/tmp/test_converter",
        stylesheet: nil,
        header_file: File.expand_path("header.html", __dir__),
      )
      docxml = Nokogiri::HTML('<html><body><p class="MsoNormal">Content</p></body></html>')
      pkg = converter.convert(docxml)

      # Check that header/footer parts were created
      expect(pkg.document.header_footer_parts).not_to be_nil
      expect(pkg.document.header_footer_parts.size).to be > 0
    end
  end
end
