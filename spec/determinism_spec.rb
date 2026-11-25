require "spec_helper"
require "fileutils"

RSpec.describe "Html2Doc output determinism" do
  let(:input) do
    <<~HTML
      <html xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">
      <head><title>blank</title></head>
      <body>
      <h1>Hello word!</h1>
      <img src="19160-6.png">
      <div>
      <ol id="1"><li>A</li></ol>
      <ol id="2" start="3"><li>B</li></ol>
      </div>
      </body></html>
    HTML
  end

  def compile
    FileUtils.rm_rf("spec/det_test_files")
    FileUtils.rm_f(%w[spec/det_test.doc spec/det_test.htm])
    Html2Doc.new(filename: "spec/det_test", imagedir: "spec",
                 stylesheet: "spec/wordstyle-nopagesize.css",
                 liststyles: { ul: "l0", ol: "l1" })
      .process(input)
    File.binread("spec/det_test.doc")
  end

  it "produces byte-identical .doc MHTML output across runs" do
    expect(compile).to eq(compile)
  end

  it "mints a stable, content-derived MIME boundary" do
    first = compile
    second = compile
    boundary = first[/boundary="([^"]+)"/, 1]

    expect(boundary).to start_with("----=_NextPart_")
    expect(boundary).to eq(second[/boundary="([^"]+)"/, 1])
    expect(first.scan(/#{Regexp.escape(boundary)}/).count)
      .to eq(first.scan(/----=_NextPart_/).count)
  end

  it "mints image part names deterministically from image content" do
    uuid = compile[%r{cid:([0-9a-f-]{36}\.png)}, 1]

    expect(uuid).not_to be_nil
    expect(compile).to include("cid:#{uuid}")
  end

  it "mints list ids deterministically" do
    doc = compile
    ids = doc.scan(/mso-list-id:(\d{9})/).flatten

    expect(ids).not_to be_empty
    expect(ids).to eq(compile.scan(/mso-list-id:(\d{9})/).flatten)
    expect(ids.uniq.length).to eq(ids.length)
  end
end
