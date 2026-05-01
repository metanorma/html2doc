require "spec_helper"
require "html2doc"

RSpec.describe "MHT output via Uniword" do
  def generate_mht(html, options = {})
    filename = options.delete(:filename) || "/tmp/mht_test_output"
    stylesheet = options.delete(:stylesheet)
    header_file = options.delete(:header_file)

    Html2Doc.new(
      filename: filename,
      stylesheet: stylesheet,
      header_file: header_file,
      output_format: :mht,
    ).process(html)

    path = "#{filename}.doc"
    expect(File.exist?(path)).to be true

    File.read(path, encoding: "UTF-8")
  end

  # Decode QP-encoded MHT for content assertions.
  def decode_mht(mht)
    mht.gsub(/=\r\n/, "").gsub(/=([0-9A-Fa-f]{2})/) { $1.hex.chr }
  end

  shared_examples "valid mht package" do
    it "starts with MIME headers" do
      expect(mht).to start_with("MIME-Version: 1.0")
    end

    it "has multipart/related content type" do
      expect(mht).to include("Content-Type: multipart/related")
    end

    it "contains text/html content part" do
      expect(mht).to include('Content-Type: text/html')
    end

    it "has a MIME boundary terminator" do
      boundary = mht.match(/boundary="([^"]+)"/)[1]
      expect(mht).to include("--#{boundary}--")
    end
  end

  describe "simple document" do
    let(:mht) do
      generate_mht('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">Hello world</p>
        <p class="MsoNormal">Second paragraph</p>
      </body></html>')
    end

    include_examples "valid mht package"

    it "contains the paragraph text" do
      decoded = decode_mht(mht)
      expect(decoded).to include("Hello world")
      expect(decoded).to include("Second paragraph")
    end
  end

  describe "headings" do
    let(:mht) do
      generate_mht('<html><head><title>Test</title></head><body>
        <h1>Heading 1</h1>
        <h2>Heading 2</h2>
      </body></html>')
    end

    include_examples "valid mht package"

    it "contains heading text" do
      decoded = decode_mht(mht)
      expect(decoded).to include("Heading 1")
      expect(decoded).to include("Heading 2")
    end
  end

  describe "inline formatting" do
    let(:mht) do
      generate_mht('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">
          Normal <b>bold</b> <i>italic</i> <b><i>bold italic</i></b>
        </p>
      </body></html>')
    end

    include_examples "valid mht package"

    it "contains formatted text" do
      decoded = decode_mht(mht)
      expect(decoded).to include("bold")
      expect(decoded).to include("italic")
    end
  end

  describe "tables" do
    let(:mht) do
      generate_mht('<html><head><title>Test</title></head><body>
        <table>
          <tr><th>Header 1</th><th>Header 2</th></tr>
          <tr><td>Cell 1</td><td>Cell 2</td></tr>
        </table>
      </body></html>')
    end

    include_examples "valid mht package"

    it "renders table structure" do
      decoded = decode_mht(mht)
      expect(decoded).to include("<table>")
      expect(decoded).to include("<tr>")
      expect(decoded).to include("<td>")
    end
  end

  describe "output file extension" do
    it "writes to .doc extension" do
      generate_mht('<html><head><title>Test</title></head><body>
        <p class="MsoNormal">Content</p>
      </body></html>', filename: "/tmp/mht_ext_test")

      expect(File.exist?("/tmp/mht_ext_test.doc")).to be true
    end
  end
end
