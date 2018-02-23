def html_input(x)
  <<~HTML
    <html><head><title>blank</title>
    <meta name="Originator" content="Me"/>
    </head>
    <body>
    #{x}
    </body></html>
  HTML
end

def html_input_no_title(x)
  <<~HTML
    <html><head>
    <meta name="Originator" content="Me"/>
    </head>
    <body>
    #{x}
    </body></html>
  HTML
end

def html_input_empty_head(x)
  <<~HTML
    <html><head></head>
    <body>
    #{x}
    </body></html>
  HTML
end

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

WORD_HDR_END = <<~HDR
-->
]]></style>
<meta name="Originator" content="Me"/>
</head>
HDR

def word_body(x, fn)
  <<~BODY
<body>
  #{x}
#{fn}</body></html>
  BODY
end

WORD_FTR1 = <<~FTR
  ------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: application/xml

PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
bGVsaXN0LnhtbCIvPgo8L3htbD4K

------=_NextPart_--
FTR

WORD_FTR2 = <<~FTR
  ------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: application/xml
PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
bGVsaXN0LnhtbCIvPgogIDxvOkZpbGUgSFJlZj0iaGVhZGVyLmh0bWwiLz4KPC94bWw+Cg==
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/header.html
Content-Transfer-Encoding: base64
Content-Type: text/html charset="utf-8"
PGh0bWwgeG1sbnM6dj0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTp2bWwiDQp4bWxuczpvPSJ1
cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiDQp4bWxuczp3PSJ1cm46c2No
ZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTp3b3JkIg0KeG1sbnM6bT0iaHR0cDovL3NjaGVtYXMu
bWljcm9zb2Z0LmNvbS9vZmZpY2UvMjAwNC8xMi9vbW1sIg0KeG1sbnM9Imh0dHA6Ly93d3cudzMu
b3JnL1RSL1JFQy1odG1sNDAiPg0KDQo8aGVhZD4NCjxtZXRhIGh0dHAtZXF1aXY9Q29udGVudC1U
eXBlIGNvbnRlbnQ9InRleHQvaHRtbDsgY2hhcnNldD11dGYtOCI+DQo8bWV0YSBuYW1lPVByb2dJ
ZCBjb250ZW50PVdvcmQuRG9jdW1lbnQ+DQo8bWV0YSBuYW1lPUdlbmVyYXRvciBjb250ZW50PSJN
aWNyb3NvZnQgV29yZCAxNSI+DQo8bWV0YSBuYW1lPU9yaWdpbmF0b3IgY29udGVudD0iTWljcm9z
b2Z0IFdvcmQgMTUiPg0KPGxpbmsgaWQ9TWFpbi1GaWxlIHJlbD1NYWluLUZpbGUgaHJlZj0iLi4v
cmljZS5nYi5odG1sIj4NCjwhLS1baWYgZ3RlIG1zbyA5XT48eG1sPg0KIDxvOnNoYXBlZGVmYXVs
dHMgdjpleHQ9ImVkaXQiIHNwaWRtYXg9IjIwNDkiLz4NCjwveG1sPjwhW2VuZGlmXS0tPg0KPC9o
ZWFkPg0KDQo8Ym9keSBsYW5nPVpIIGxpbms9Ymx1ZSB2bGluaz1wdXJwbGU+DQoNCjxkaXYgc3R5
bGU9J21zby1lbGVtZW50OmZvb3Rub3RlLXNlcGFyYXRvcicgaWQ9ZnM+DQoNCjxwIGNsYXNzPU1z
b05vcm1hbD48c3BhbiBsYW5nPUVOLVVTPjxzcGFuIHN0eWxlPSdtc28tc3BlY2lhbC1jaGFyYWN0
ZXI6Zm9vdG5vdGUtc2VwYXJhdG9yJz48IVtpZiAhc3VwcG9ydEZvb3Rub3Rlc10+DQoNCjxociBh
bGlnbj1sZWZ0IHNpemU9MSB3aWR0aD0iMzMlIj4NCg0KPCFbZW5kaWZdPjwvc3Bhbj48L3NwYW4+
PC9wPg0KDQo8L2Rpdj4NCg0KPGRpdiBzdHlsZT0nbXNvLWVsZW1lbnQ6Zm9vdG5vdGUtY29udGlu
dWF0aW9uLXNlcGFyYXRvcicgaWQ9ZmNzPg0KDQo8cCBjbGFzcz1Nc29Ob3JtYWw+PHNwYW4gbGFu
Zz1FTi1VUz48c3BhbiBzdHlsZT0nbXNvLXNwZWNpYWwtY2hhcmFjdGVyOmZvb3Rub3RlLWNvbnRp
bnVhdGlvbi1zZXBhcmF0b3InPjwhW2lmICFzdXBwb3J0Rm9vdG5vdGVzXT4NCg0KPGhyIGFsaWdu
PWxlZnQgc2l6ZT0xPg0KDQo8IVtlbmRpZl0+PC9zcGFuPjwvc3Bhbj48L3A+DQoNCjwvZGl2Pg0K
DQo8ZGl2IHN0eWxlPSdtc28tZWxlbWVudDplbmRub3RlLXNlcGFyYXRvcicgaWQ9ZXM+DQoNCjxw
IGNsYXNzPU1zb05vcm1hbD48c3BhbiBsYW5nPUVOLVVTPjxzcGFuIHN0eWxlPSdtc28tc3BlY2lh
bC1jaGFyYWN0ZXI6Zm9vdG5vdGUtc2VwYXJhdG9yJz48IVtpZiAhc3VwcG9ydEZvb3Rub3Rlc10+
DQoNCjxociBhbGlnbj1sZWZ0IHNpemU9MSB3aWR0aD0iMzMlIj4NCg0KPCFbZW5kaWZdPjwvc3Bh
bj48L3NwYW4+PC9wPg0KDQo8L2Rpdj4NCg0KPGRpdiBzdHlsZT0nbXNvLWVsZW1lbnQ6ZW5kbm90
ZS1jb250aW51YXRpb24tc2VwYXJhdG9yJyBpZD1lY3M+DQoNCjxwIGNsYXNzPU1zb05vcm1hbD48
c3BhbiBsYW5nPUVOLVVTPjxzcGFuIHN0eWxlPSdtc28tc3BlY2lhbC1jaGFyYWN0ZXI6Zm9vdG5v
dGUtY29udGludWF0aW9uLXNlcGFyYXRvcic+PCFbaWYgIXN1cHBvcnRGb290bm90ZXNdPg0KDQo8
aHIgYWxpZ249bGVmdCBzaXplPTE+DQoNCjwhW2VuZGlmXT48L3NwYW4+PC9zcGFuPjwvcD4NCg0K
PC9kaXY+DQoNCjxkaXYgc3R5bGU9J21zby1lbGVtZW50OmhlYWRlcicgaWQ9aDI+DQoNCjxwIGNs
YXNzPU1zb0hlYWRlcj48c3BhbiBsYW5nPUVOLVVTPkRCMTEvQ0QgMTczMDEtMTwvc3Bhbj48c3Bh
biBsYW5nPUVOLVVTDQpzdHlsZT0nZm9udC1mYW1pbHk6IlRpbWVzIE5ldyBSb21hbiIsc2VyaWY7
bXNvLWFzY2lpLWZvbnQtZmFtaWx5OlNpbUhlaSc+4oCUPC9zcGFuPjxzcGFuDQpsYW5nPUVOLVVT
PjIwMTY8L3NwYW4+PC9wPg0KDQo8L2Rpdj4NCg0KPGRpdiBzdHlsZT0nbXNvLWVsZW1lbnQ6Zm9v
dGVyJyBpZD1mMj4NCg0KPHAgY2xhc3M9TXNvRm9vdGVyPjwhLS1baWYgc3VwcG9ydEZpZWxkc10+
PHNwYW4gbGFuZz1FTi1VUz48c3BhbiBzdHlsZT0nbXNvLWVsZW1lbnQ6DQpmaWVsZC1iZWdpbic+
PC9zcGFuPjxzcGFuIHN0eWxlPSdtc28tc3BhY2VydW46eWVzJz7CoDwvc3Bhbj5QQUdFPHNwYW4N
CnN0eWxlPSdtc28tc3BhY2VydW46eWVzJz7CoCA8L3NwYW4+XCogTUVSR0VGT1JNQVQgPHNwYW4g
c3R5bGU9J21zby1lbGVtZW50OmZpZWxkLXNlcGFyYXRvcic+PC9zcGFuPjwvc3Bhbj48IVtlbmRp
Zl0tLT48c3Bhbg0KbGFuZz1lbCBzdHlsZT0nbXNvLWFuc2ktbGFuZ3VhZ2U6IzA0MDA7bXNvLWZh
cmVhc3QtbGFuZ3VhZ2U6IzA0MDA7bXNvLW5vLXByb29mOg0KeWVzJz40Mjwvc3Bhbj48IS0tW2lm
IHN1cHBvcnRGaWVsZHNdPjxzcGFuIGxhbmc9RU4tVVM+PHNwYW4gc3R5bGU9J21zby1lbGVtZW50
Og0KZmllbGQtZW5kJz48L3NwYW4+PC9zcGFuPjwhW2VuZGlmXS0tPjwvcD4NCg0KPC9kaXY+DQoN
CjwvYm9keT4NCg0KPC9odG1sPg0K

------=_NextPart_--
FTR

WORD_FTR3 = <<~FTR
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/609e8807-c2d0-450c-b60b-d995a0f8dcaf.png
Content-Transfer-Encoding: base64
Content-Type: image/png
FTR

WORD_FTR3 = <<~FTR
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: application/xml

PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9IjFh
YzIwNjVmLTAzZjAtNGM3YS1iOWE2LTkyZTgyMDU5MWJmMC5wbmciLz4KICA8bzpGaWxlIEhSZWY9
ImZpbGVsaXN0LnhtbCIvPgo8L3htbD4K
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/cb7b0d19-891e-4634-815a-570d019d454c.png
Content-Transfer-Encoding: base64
Content-Type: image/png
------=_NextPart_--
FTR

DEFAULT_STYLESHEET = File.read("lib/html2doc/wordstyle.css", encoding: "utf-8").freeze

def guid_clean(x)
  x.gsub(/NextPart_[0-9a-f.]+/, "NextPart_")
end

def image_clean(x)
  x.gsub(%r{[0-9a-f-]+\.png}, "image.png").
    gsub(%r{[0-9a-f-]+\.gif}, "image.gif").
    gsub(%r{[0-9a-f-]+\.(jpeg|jpg)}, "image.jpg").
    gsub(%r{------=_NextPart_\s+Content-Location: file:///C:/Doc/test_files/image\.(png|gif).*?\s-----=_NextPart_}m, "------=_NextPart_").
    gsub(%r{Content-Type: image/(png|gif|jpeg)[^-]*------=_NextPart_-?-?}m, "").
    gsub(%r{ICAgICAg[^-]*-----}m, "-----").
    gsub(%r{\s*</img>\s*}m, "</img>").
    gsub(%r{</body>\s*</html>}m, "</body></html>")
end

RSpec.describe Html2Doc do
  it "has a version number" do
    expect(Html2Doc::VERSION).not_to be nil
  end

  it "processes a blank document" do
    Html2Doc.process(html_input(""), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "removes any temp files" do
    File.delete("test.doc")
    Html2Doc.process(html_input(""), "test", nil, nil, nil, nil)
    expect(File.exist?("test.doc")).to be true
    expect(File.exist?("test.htm")).to be false
    expect(File.exist?("test_files")).to be false
  end

  it "processes a stylesheet in an HTML document with a title" do
    Html2Doc.process(html_input(""), "test", "lib/html2doc/wordstyle.css", nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "processes a stylesheet in an HTML document without a title" do
    Html2Doc.process(html_input_no_title(""), "test", "lib/html2doc/wordstyle.css", nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR.sub("<title>blank</title>", "")} 
    #{DEFAULT_STYLESHEET} #{WORD_HDR_END} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "processes a stylesheet in an HTML document with an empty head" do
    Html2Doc.process(html_input_empty_head(""), "test", "lib/html2doc/wordstyle.css", nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR.sub("<title>blank</title>", "")}
    #{DEFAULT_STYLESHEET} 
    #{WORD_HDR_END.sub('<meta name="Originator" content="Me"/>'+"\n", "").sub("</style>\n</head>", "</style></head>")} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "processes a header" do
    Html2Doc.process(html_input(""), "test", nil, "header.html", nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET.gsub(/FILENAME/, "test")} 
    #{WORD_HDR_END} #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR2}
    OUTPUT
  end

  it "processes a populated document" do
    simple_body = "<h1>Hello word!</h1>
    <div>This is a very simple document</div>"
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body, '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "processes AsciiMath" do
    Html2Doc.process(html_input("<div>{{sum_(i=1)^n i^3=((n(n+1))/2)^2}}</div>"), "test", nil, nil, nil, ["{{", "}}"])
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div><m:oMath><m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:grow m:val="1"></m:grow><m:subHide m:val="off"></m:subHide><m:supHide m:val="off"></m:supHide></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e></m:e></m:nary><m:sSup><m:e><m:r><m:t>i</m:t></m:r></m:e><m:sup><m:r><m:t>3</m:t></m:r></m:sup></m:sSup><m:r><m:t>=</m:t></m:r><m:sSup><m:e><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val="bar"></m:type></m:fPr><m:num><m:r><m:t>n</m:t></m:r><m:r><m:t>(n+1)</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>
    </div>', '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "processes tabs" do
    simple_body = "<h1>Hello word!</h1>
    <div>This is a very &tab; simple document</div>"
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body.gsub(/\&tab;/, %[<span style="mso-tab-count:1">&#xA0; </span>]), '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "makes unstyled paragraphs be MsoNormal" do
    simple_body = '<h1>Hello word!</h1>
    <p>This is a very simple document</p>
    <p class="x">This style stays</p>'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body.gsub(/<p>/, %[<p class="MsoNormal">]), '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "makes unstyled list entries be MsoNormal" do
    simple_body = '<h1>Hello word!</h1>
    <ul>
    <li>This is a very simple document</li>
    <li class="x">This style stays</li>
    </ul>'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body.gsub(/<li>/, %[<li class="MsoNormal">]), '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "resizes images for height" do
    simple_body = '<img src="spec/19160-6.png">'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/png})
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="test_files/cb7b0d19-891e-4634-815a-570d019d454c.png" width="400" height="387"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR3)}
    OUTPUT
  end

  it "resizes images for width" do
    simple_body = '<img src="spec/19160-7.gif">'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/gif})
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="test_files/cb7b0d19-891e-4634-815a-570d019d454c.gif" width="400" height="118"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR3).gsub(/image\.png/, "image.gif")}
    OUTPUT
  end

  it "resizes images for height" do
    simple_body = '<img src="spec/19160-8.jpg">'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/jpeg})
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="test_files/cb7b0d19-891e-4634-815a-570d019d454c.jpg" width="208" height="680"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR3).gsub(/image\.png/, "image.jpg")}
    OUTPUT
  end

  it "processes epub:type footnotes" do
    simple_body = '<div>This is a very simple 
     document<a epub:type="footnote" href="#a1">1</a> allegedly<a epub:type="footnote" href="#a2">2</a></div>
     <aside id="a1">Footnote</aside>
     <aside id="a2">Other Footnote</aside>'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>This is a very simple
    document<a epub:type="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> allegedly<a epub:type="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
'<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Footnote</p></div>
<div style="mso-element:footnote" id="ftn2">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn2" href="#_ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Other Footnote</p></div>
</div>')}
    #{WORD_FTR1}
    OUTPUT
  end

    it "processes class footnotes" do
    simple_body = '<div>This is a very simple
     document<a class="footnote" href="#a1">1</a> allegedly<a class="footnote" href="#a2">2</a></div>
     <aside id="a1">Footnote</aside>
     <aside id="a2">Other Footnote</aside>'
    Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>This is a very simple
    document<a class="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> allegedly<a class="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
'<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Footnote</p></div>
<div style="mso-element:footnote" id="ftn2">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn2" href="#_ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Other Footnote</p></div>
</div>')}
    #{WORD_FTR1}
    OUTPUT
  end

    it "extracts paragraphs from footnotes" do
      simple_body = '<div>This is a very simple
     document<a class="footnote" href="#a1">1</a> allegedly<a class="footnote" href="#a2">2</a></div>
     <aside id="a1"><p>Footnote</p></aside>
     <div id="a2"><p>Other Footnote</p></div>'
      Html2Doc.process(html_input(simple_body), "test", nil, nil, nil, nil)
      expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
        to match_fuzzy(<<~OUTPUT)
      #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
      #{word_body('<div>This is a very simple
    document<a class="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> allegedly<a class="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
      '<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
<p class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Footnote</p></div>
<div style="mso-element:footnote" id="ftn2">
<p class="MsoFootnoteText"><a style="mso-footnote-id:ftn2" href="#_ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Other Footnote</p></div>
</div>')}
      #{WORD_FTR1}
      OUTPUT
    end
end
