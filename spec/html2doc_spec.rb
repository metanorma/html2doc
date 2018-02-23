require 'rspec/match_fuzzy'

BLANK_HTML = <<~HTML.freeze
    <html><head><title>blank</title></head>
    <body></body></html>
HTML

WORD_HDR = <<~HDR
MIME-Version: 1.0
Content-Type: multipart/related; boundary="----=_NextPart_"

------=_NextPart_
Content-Location: file:///C:/Doc/test.htm
Content-Type: text/html; charset="utf-8"

<?xml version="1.0"?>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]>
<xml>
<w:WordDocument>
<w:View>Print</w:View>
<w:Zoom>100</w:Zoom>
<w:DoNotOptimizeForBrowser/>
</w:WordDocument>
</xml>
<![endif]-->
<meta http-equiv=Content-Type content="text/html; charset=utf-8"/>

  <link rel=File-List href="test_files/filelist.xml"/>
<title>blank</title><style><![CDATA[
  <!--
HDR

WORD_BODY = <<~BODY

-->
]]></style></head>
<body><div style="mso-element:footnote-list"/></body></html>
BODY

WORD_FTR = <<~FTR
  ------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: application/xml

PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
bGVsaXN0LnhtbCIvPgo8L3htbD4K

------=_NextPart_--
FTR

DEFAULT_STYLESHEET = File.read("lib/html2doc/wordstyle.css").freeze

def guid_clean(x)
  x.gsub(/NextPart_[0-9a-f.]+/, "NextPart_")
end

RSpec.describe Html2Doc do
  it "has a version number" do
    expect(Html2Doc::VERSION).not_to be nil
  end

  it "processes a blank document" do
    Html2Doc.process(BLANK_HTML, "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR}
    #{DEFAULT_STYLESHEET}
    #{WORD_BODY}
    #{WORD_FTR}
    OUTPUT
  end
end
