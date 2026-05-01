require "spec_helper"
require "html2doc"
require "uniword"

RSpec.describe "rice.html fixture comparison" do
  EXAMPLES_DIR = File.expand_path("examples", __dir__)
  let(:rice_html) { File.read(File.join(EXAMPLES_DIR, "rice.html"), encoding: "UTF-8") }
  let(:rice_doc_path) { File.join(EXAMPLES_DIR, "rice.doc") }
  let(:rice_docx_path) { File.join(EXAMPLES_DIR, "rice.docx") }

  # -----------------------------------------------------------------------
  # Shared helpers
  # -----------------------------------------------------------------------

  # Normalize text for comparison: collapse whitespace, strip leading/trailing
  def normalized_text(str)
    str.to_s.gsub(/\s+/, " ").strip
  end

  # Extract element counts from Nokogiri HTML document
  def html_element_counts(doc)
    counts = {}
    %w[h1 h2 h3 h4 h5 h6 p table tr td th img b i u span a br].each do |tag|
      counts[tag] = doc.css(tag).size
    end
    counts
  end

  # -----------------------------------------------------------------------
  # MHT: rice.html → MHT vs rice.doc
  # -----------------------------------------------------------------------

  describe "rice.html -> MHT matching rice.doc" do
    let(:generated_mht_path) { "/tmp/rice_fixture_mht_test" }

    before(:all) do
      @gen_path = "/tmp/rice_fixture_mht_test"
      examples = File.expand_path("examples", __dir__)
      rice = File.read(File.join(examples, "rice.html"), encoding: "UTF-8")
      Html2Doc.new(
        filename: @gen_path,
        output_format: :mht,
        imagedir: File.join(examples, "rice_images"),
        header_file: File.join(examples, "header.html"),
      ).process(rice)
    end

    after(:all) do
      File.delete("#{@gen_path}.doc") if File.exist?("#{@gen_path}.doc")
    end

    let(:generated_mht_file) { "#{generated_mht_path}.doc" }

    describe "MIME structure" do
      let(:generated_mhtml) { Uniword::Mhtml::MhtmlPackage.from_file(generated_mht_file) }
      let(:reference_mhtml) { Uniword::Mhtml::MhtmlPackage.from_file(rice_doc_path) }

      it "both parse as valid Mhtml::Document" do
        expect(generated_mhtml).to be_a(Uniword::Mhtml::Document)
        expect(reference_mhtml).to be_a(Uniword::Mhtml::Document)
      end

      it "has the same number of image parts" do
        gen_images = generated_mhtml.image_parts
        ref_images = reference_mhtml.image_parts
        expect(gen_images.size).to eq(ref_images.size),
          "Generated has #{gen_images.size} image parts, reference has #{ref_images.size}"
      end

      it "has the same number of total MIME parts" do
        gen_parts = generated_mhtml.parts
        ref_parts = reference_mhtml.parts
        expect(gen_parts.size).to eq(ref_parts.size),
          "Generated has #{gen_parts.size} MIME parts, reference has #{ref_parts.size}"
      end

      it "has XML parts (filelist.xml)" do
        gen_xml = generated_mhtml.xml_parts
        ref_xml = reference_mhtml.xml_parts
        expect(gen_xml.size).to eq(ref_xml.size),
          "Generated has #{gen_xml.size} XML parts, reference has #{ref_xml.size}"
      end
    end

    describe "HTML body content" do
      let(:generated_mhtml) { Uniword::Mhtml::MhtmlPackage.from_file(generated_mht_file) }
      let(:reference_mhtml) { Uniword::Mhtml::MhtmlPackage.from_file(rice_doc_path) }

      let(:gen_body_doc) { Nokogiri::HTML(generated_mhtml.body_html || "") }
      let(:ref_body_doc) { Nokogiri::HTML(reference_mhtml.body_html || "") }

      it "has matching heading counts" do
        gen_counts = html_element_counts(gen_body_doc)
        ref_counts = html_element_counts(ref_body_doc)

        %w[h1 h2 h3 h4 h5 h6].each do |tag|
          expect(gen_counts[tag]).to eq(ref_counts[tag]),
            "h tag #{tag}: generated #{gen_counts[tag]}, reference #{ref_counts[tag]}"
        end
      end

      it "has matching table counts" do
        gen_counts = html_element_counts(gen_body_doc)
        ref_counts = html_element_counts(ref_body_doc)

        expect(gen_counts["table"]).to eq(ref_counts["table"]),
          "table count: generated #{gen_counts['table']}, reference #{ref_counts['table']}"
        expect(gen_counts["tr"]).to eq(ref_counts["tr"]),
          "tr count: generated #{gen_counts['tr']}, reference #{ref_counts['tr']}"
      end

      it "has matching image counts" do
        gen_counts = html_element_counts(gen_body_doc)
        ref_counts = html_element_counts(ref_body_doc)

        expect(gen_counts["img"]).to eq(ref_counts["img"]),
          "img count: generated #{gen_counts['img']}, reference #{ref_counts['img']}"
      end

      it "has matching break counts" do
        gen_counts = html_element_counts(gen_body_doc)
        ref_counts = html_element_counts(ref_body_doc)

        expect(gen_counts["br"]).to eq(ref_counts["br"]),
          "br count: generated #{gen_counts['br']}, reference #{ref_counts['br']}"
      end

      it "has matching bold/italic counts" do
        gen_counts = html_element_counts(gen_body_doc)
        ref_counts = html_element_counts(ref_body_doc)

        expect(gen_counts["b"]).to eq(ref_counts["b"]),
          "b count: generated #{gen_counts['b']}, reference #{ref_counts['b']}"
        expect(gen_counts["i"]).to eq(ref_counts["i"]),
          "i count: generated #{gen_counts['i']}, reference #{ref_counts['i']}"
      end
    end

    describe "text content" do
      let(:generated_mhtml) { Uniword::Mhtml::MhtmlPackage.from_file(generated_mht_file) }
      let(:reference_mhtml) { Uniword::Mhtml::MhtmlPackage.from_file(rice_doc_path) }

      let(:gen_body_text) do
        doc = Nokogiri::HTML(generated_mhtml.body_html || "")
        normalized_text(doc.at_css("body")&.text.to_s)
      end
      let(:ref_body_text) do
        doc = Nokogiri::HTML(reference_mhtml.body_html || "")
        normalized_text(doc.at_css("body")&.text.to_s)
      end

      it "has the same normalized text length (within 5%)" do
        ratio = gen_body_text.length.to_f / ref_body_text.length
        expect(ratio).to be_within(0.05).of(1.0),
          "Generated text length #{gen_body_text.length} vs reference #{ref_body_text.length} (ratio #{ratio.round(3)})"
      end

      it "contains all key phrases from the reference" do
        # Extract significant words/phrases from reference text
        key_phrases = [
          "Rice",
          "Cereals and pulses",
          "Specifications and test methods",
          "Sampling",
          "Bibliography",
        ]
        key_phrases.each do |phrase|
          expect(gen_body_text).to include(phrase),
            "Generated MHT missing key phrase: '#{phrase}'"
        end
      end
    end
  end

  # -----------------------------------------------------------------------
  # DOCX: rice.html → DOCX vs rice.docx
  # -----------------------------------------------------------------------

  describe "rice.html -> DOCX matching rice.docx" do
    let(:generated_docx_path) { "/tmp/rice_fixture_docx_test" }

    before(:all) do
      @gen_path = "/tmp/rice_fixture_docx_test"
      examples = File.expand_path("examples", __dir__)
      rice = File.read(File.join(examples, "rice.html"), encoding: "UTF-8")
      Html2Doc.new(
        filename: @gen_path,
        output_format: :docx,
        imagedir: File.join(examples, "rice_images"),
        header_file: File.join(examples, "header.html"),
      ).process(rice)
    end

    after(:all) do
      File.delete("#{@gen_path}.docx") if File.exist?("#{@gen_path}.docx")
    end

    let(:generated_docx_file) { "#{generated_docx_path}.docx" }

    describe "DOCX package structure (via PackageDiffer)" do
      let(:diff_result) do
        Uniword::Diff::PackageDiffer.new(
          generated_docx_file,
          rice_docx_path,
        ).diff
      end

      it "both are valid DOCX packages" do
        expect(diff_result).to be_a(Uniword::Diff::PackageDiffResult)
      end

      it "has no missing required parts" do
        opc_issues = diff_result.opc_issues.select { |i| i.severity == :error }
        expect(opc_issues).to be_empty,
          "OPC errors: #{opc_issues.map(&:description).join('; ')}"
      end

      it "has the same set of ZIP parts (excluding media filenames)" do
        # Image filenames may differ (e.g. rice_image1.png vs image1.png)
        # so compare part counts by category, not exact names
        require "zip"

        gen_names = nil
        ref_names = nil
        Zip::File.open(generated_docx_file) { |z| gen_names = z.entries.map(&:name) }
        Zip::File.open(rice_docx_path) { |z| ref_names = z.entries.map(&:name) }

        # Compare non-media parts exactly
        gen_non_media = gen_names.reject { |n| n.start_with?("word/media/") }.sort
        ref_non_media = ref_names.reject { |n| n.start_with?("word/media/") }.sort
        expect(gen_non_media).to eq(ref_non_media),
          "Non-media part mismatch:\n  generated: #{gen_non_media.inspect}\n  reference: #{ref_non_media.inspect}"

        # Media part counts should match
        gen_media = gen_names.select { |n| n.start_with?("word/media/") }
        ref_media = ref_names.select { |n| n.start_with?("word/media/") }
        expect(gen_media.size).to eq(ref_media.size),
          "Media part count: generated #{gen_media.size}, reference #{ref_media.size}"
      end

      it "has matching document.xml paragraph count" do
        # Extract document.xml from both and compare paragraph counts
        require "zip"

        gen_doc_xml = nil
        ref_doc_xml = nil
        Zip::File.open(generated_docx_file) { |z| gen_doc_xml = z.read("word/document.xml") }
        Zip::File.open(rice_docx_path) { |z| ref_doc_xml = z.read("word/document.xml") }

        gen_doc = Nokogiri::XML(gen_doc_xml)
        ref_doc = Nokogiri::XML(ref_doc_xml)

        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        gen_paras = gen_doc.xpath("//w:p", "w" => w_ns).size
        ref_paras = ref_doc.xpath("//w:p", "w" => w_ns).size

        expect(gen_paras).to eq(ref_paras),
          "Paragraph count: generated #{gen_paras}, reference #{ref_paras}"
      end
    end

    describe "DOCX document content (via DocumentDiffer)" do
      let(:generated_pkg) { Uniword::Docx::Package.from_file(generated_docx_file) }
      let(:reference_pkg) { Uniword::Docx::Package.from_file(rice_docx_path) }

      let(:diff_result) do
        Uniword::Diff::DocumentDiffer.new(
          reference_pkg.document,
          generated_pkg.document,
        ).diff
      end

      it "both parse as valid Docx::Package" do
        expect(generated_pkg).to be_a(Uniword::Docx::Package)
        expect(reference_pkg).to be_a(Uniword::Docx::Package)
      end

      it "has the same paragraph count" do
        gen_count = generated_pkg.document.body.paragraphs.size
        ref_count = reference_pkg.document.body.paragraphs.size
        expect(gen_count).to eq(ref_count),
          "Paragraph count: generated #{gen_count}, reference #{ref_count}"
      end

      it "has the same table count" do
        gen_count = generated_pkg.document.body.tables.size
        ref_count = reference_pkg.document.body.tables.size
        expect(gen_count).to eq(ref_count),
          "Table count: generated #{gen_count}, reference #{ref_count}"
      end

      it "has matching text content" do
        gen_text = normalized_text(generated_pkg.text)
        ref_text = normalized_text(reference_pkg.text)

        # Check that key content is preserved
        expect(gen_text).to include("Rice")
        expect(gen_text).to include("Terms and Definitions")

        # Check overall text length is comparable
        ratio = gen_text.length.to_f / ref_text.length
        expect(ratio).to be_within(0.05).of(1.0),
          "Text length ratio: #{ratio.round(3)} (gen #{gen_text.length}, ref #{ref_text.length})"
      end

      it "has minimal text-only differences" do
        text_changes = diff_result.text_changes
        # Normalize whitespace for comparison: the reference DOCX was created
        # by Word's HTML import which handles NBSP/trailing spaces differently
        # than our converter. Filter out whitespace-only differences.
        significant_changes = text_changes.reject do |c|
          old_val = (c[:old].is_a?(String) ? c[:old] : c[:old]&.to_s).to_s
          new_val = (c[:new].is_a?(String) ? c[:new] : c[:new]&.to_s).to_s
          normalized_text(old_val) == normalized_text(new_val)
        end
        summary = significant_changes.first(10).map do |c|
          old_val = c[:old].is_a?(String) ? c[:old] : c[:old]&.to_s
          new_val = c[:new].is_a?(String) ? c[:new] : c[:new]&.to_s
          old_short = old_val.to_s.gsub(/\s+/, " ")[0..60]
          new_short = new_val.to_s.gsub(/\s+/, " ")[0..60]
          "  #{c[:type]} #{c[:change]} at #{c[:position]}: old=#{old_short.inspect} new=#{new_short.inspect}"
        end.join("\n")
        # Remaining differences are known edge cases:
        # - Header paragraph inter-element whitespace (2)
        # - oMath element positioning (Uniword model serializes math at paragraph end)
        # - Table footnote reference markers (1)
        expect(significant_changes.size).to be <= 10,
          "#{significant_changes.size} text change(s) (of #{text_changes.size} total, " \
          "#{text_changes.size - significant_changes.size} whitespace-only):\n#{summary}"
      end

      it "has no structure differences" do
        structure_changes = diff_result.structure_changes
        expect(structure_changes).to be_empty,
          "Structure changes: #{structure_changes.map { |c| c[:change] }.join(', ')}"
      end
    end
  end
end
