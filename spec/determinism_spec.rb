require "spec_helper"
require "digest"
require "fileutils"

RSpec.describe "Html2Doc output determinism" do
  let(:input) do
    <<~HTML
      <html xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">
      <head><title>blank</title></head>
      <body>
      <h1>Hello word!</h1>
      <img src="19160-6.png">
      <img src="19160-6.png">
      <img src="19160-7.gif">
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

  it "names image parts by their SHA-256 content address" do
    addresses = compile.scan(/cid:([0-9a-f]{64}\.(?:png|gif))/).flatten.uniq
    png = "#{Digest::SHA256.file('spec/19160-6.png').hexdigest}.png"
    gif = "#{Digest::SHA256.file('spec/19160-7.gif').hexdigest}.gif"

    expect(addresses.sort).to eq([gif, png].sort)
  end

  it "collapses identical images to a single content-addressed part" do
    cids = compile.scan(/cid:[0-9a-f]{64}\.png/).uniq

    expect(cids.length).to eq(1)
  end

  it "mints list ids deterministically" do
    doc = compile
    ids = doc.scan(/mso-list-id:(\d{9})/).flatten

    expect(ids).not_to be_empty
    expect(ids).to eq(compile.scan(/mso-list-id:(\d{9})/).flatten)
    expect(ids.uniq.length).to eq(ids.length)
  end
end
