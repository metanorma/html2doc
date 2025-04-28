require "base64"

def html_input(xml)
  <<~HTML
      <html xmlns:epub="http://www.idpf.org/2007/ops" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" lang="en" xml:lang="en"><head><title>blank</title>
      <meta name="Originator" content="Me"/>
      </head>
      <body>
    #{xml}
      </body></html>
  HTML
end

def html_input_no_title(xml)
  <<~HTML
      <html xmlns:epub="http://www.idpf.org/2007/ops" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" lang="en" xml:lang="en"><head>
      <meta name="Originator" content="Me"/>
      </head>
      <body>
    #{xml}
      </body></html>
  HTML
end

def html_input_empty_head(xml)
  <<~HTML
      <html xmlns:epub="http://www.idpf.org/2007/ops" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" lang="en" xml:lang="en"><head></head>
      <body>
    #{xml}
      </body></html>
  HTML
end

def mock_plurimath_error
  expect(Plurimath::Math).to receive(:parse) do
    raise "Syntax error"
  end.exactly(2).times
end

WORD_HDR = <<~HDR.freeze
  MIME-Version: 1.0
  Content-Type: multipart/related; boundary="----=_NextPart_"

  ------=_NextPart_
  Content-ID: <test.htm>
  Content-Disposition: inline; filename="test.htm"
  Content-Type: text/html; charset="utf-8"

  <?xml version="1.0"?>
  <html xmlns:epub="http://www.idpf.org/2007/ops" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40" lang="en" xml:lang="en"><head>
  <xml>
  <w:WordDocument>
  <w:View>Print</w:View>
  <w:Zoom>100</w:Zoom>
  <w:DoNotOptimizeForBrowser/>
  </w:WordDocument>
  </xml>
  <meta http-equiv=Content-Type content="text/html; charset=utf-8"/>

    <link rel=File-List href="cid:filelist.xml"/>
  <title>blank</title><style><![CDATA[
    <!--
HDR

WORD_HDR_END = <<~HDR.freeze
  -->
  ]]></style>
  <meta name="Originator" content="Me"/>
  </head>
HDR

def word_body(xml, footnote)
  <<~BODY
    <body>
      #{xml}
      #{footnote}</body></html>
  BODY
end

WORD_FTR1 = <<~FTR.freeze
  ------=_NextPart_
  Content-ID: <filelist.xml>
  Content-Disposition: inline; filename="filelist.xml"
  Content-Transfer-Encoding: base64
  Content-Type: #{Html2Doc.new({}).mime_type('filelist.xml')}

  PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
  ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
  bGVsaXN0LnhtbCIvPgo8L3htbD4K

  ------=_NextPart_--
FTR

WORD_FTR2 = <<~FTR.freeze
  ------=_NextPart_
  Content-ID: <filelist.xml>
  Content-Disposition: inline; filename="filelist.xml"
  Content-Transfer-Encoding: base64
  Content-Type: #{Html2Doc.new({}).mime_type('filelist.xml')}
  PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
  ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
  bGVsaXN0LnhtbCIvPgogIDxvOkZpbGUgSFJlZj0iaGVhZGVyLmh0bWwiLz4KPC94bWw+Cg==
  ------=_NextPart_
FTR

WORD_FTR3 = <<~FTR.freeze
  ------=_NextPart_
  Content-ID: <filelist.xml>
  Content-Disposition: inline; filename="filelist.xml"
  Content-Transfer-Encoding: base64
  Content-Type: #{Html2Doc.new({}).mime_type('filelist.xml')}

  PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
  ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9IjFh
  YzIwNjVmLTAzZjAtNGM3YS1iOWE2LTkyZTgyMDU5MWJmMC5wbmciLz4KICA8bzpGaWxlIEhSZWY9
  ImZpbGVsaXN0LnhtbCIvPgo8L3htbD4K
  ------=_NextPart_
  Content-ID: <cb7b0d19-891e-4634-815a-570d019d454c.png>
  Content-Disposition: inline; filename="cb7b0d19-891e-4634-815a-570d019d454c.png"
  Content-Transfer-Encoding: base64
  Content-Type: image/png
  ------=_NextPart_--
FTR

HEADERHTML = <<~FTR.freeze
  <html xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:w="urn:schemas-microsoft-com:office:word"
  xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
  xmlns:mv="http://macVmlSchemaUri" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
  <meta name=Title content="">
  <meta name=Keywords content="">
  <meta http-equiv=Content-Type content="text/html; charset=utf-8">
  <meta name=ProgId content=Word.Document>
  <meta name=Generator content="Microsoft Word 15">
  <meta name=Originator content="Microsoft Word 15">
  <link id=Main-File rel=Main-File href="FILENAME.html">
  <!--[if gte mso 9]><xml>
  <o:shapedefaults v:ext="edit" spidmax="2049"/>
  </xml><![endif]-->
  </head>
  <body lang=EN link=blue vlink="#954F72">
  <div style='mso-element:footnote-separator' id=fs>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  normal'><span lang=EN-GB><span style='mso-special-character:footnote-separator'><![if !supportFootnotes]>
  <hr align=left size=1 width="33%">
  <![endif]></span></span></p>
  </div>
  <div style='mso-element:footnote-continuation-separator' id=fcs>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  normal'><span lang=EN-GB><span style='mso-special-character:footnote-continuation-separator'><![if !supportFootnotes]>
  <hr align=left size=1>
  <![endif]></span></span></p>
  </div>
  <div style='mso-element:endnote-separator' id=es>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  normal'><span lang=EN-GB><span style='mso-special-character:footnote-separator'><![if !supportFootnotes]>
  <hr align=left size=1 width="33%">
  <![endif]></span></span></p>
  </div>
  <div style='mso-element:endnote-continuation-separator' id=ecs>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  normal'><span lang=EN-GB><span style='mso-special-character:footnote-continuation-separator'><![if !supportFootnotes]>
  <hr align=left size=1>
  <![endif]></span></span></p>
  </div>
  <div style='mso-element:header' id=eh1>
  <p class=MsoHeader align=left style='text-align:left;line-height:12.0pt;
  mso-line-height-rule:exactly'><span lang=EN-GB>ISO/IEC&#x26;nbsp;CD 17301-1:2016(E)</span></p>
  </div>
  <div style='mso-element:header' id=h1>
  <p class=MsoHeader style='margin-bottom:18.0pt'><span lang=EN-GB
  style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-weight:normal'>&#xa9;
  ISO/IEC&#x26;nbsp;2016&#x26;nbsp;&#x2013; All rights reserved</span><span lang=EN-GB
  style='font-weight:normal'><o:p></o:p></span></p>
  </div>
  <div style='mso-element:footer' id=ef1>
  <p class=MsoFooter style='margin-top:12.0pt;line-height:12.0pt;mso-line-height-rule:
  exactly'><!--[if supportFields]><b style='mso-bidi-font-weight:normal'><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-element:field-begin'></span><span
  style='mso-spacerun:yes'>&#xa0;</span>PAGE<span style='mso-spacerun:yes'>&#xa0;&#xa0;
  </span>\\* MERGEFORMAT <span style='mso-element:field-separator'></span></span></b><![endif]--><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-no-proof:yes'>2</span></span></b><!--[if supportFields]><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-element:field-end'></span></span></b><![endif]--><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-tab-count:1'>&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>&#xa9;
  ISO/IEC&#x26;nbsp;2016&#x26;nbsp;&#x2013; All rights reserved<o:p></o:p></span></p>
  </div>
  <div style='mso-element:header' id=eh2>
  <p class=MsoHeader align=left style='text-align:left;line-height:12.0pt;
  mso-line-height-rule:exactly'><span lang=EN-GB>ISO/IEC&#x26;nbsp;CD 17301-1:2016(E)</span></p>
  </div>
  <div style='mso-element:header' id=h2>
  <p class=MsoHeader align=right style='text-align:right;line-height:12.0pt;
  mso-line-height-rule:exactly'><span lang=EN-GB>ISO/IEC&#x26;nbsp;CD 17301-1:2016(E)</span></p>
  </div>
  <div style='mso-element:footer' id=ef2>
  <p class=MsoFooter style='line-height:12.0pt;mso-line-height-rule:exactly'><!--[if supportFields]><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-element:field-begin'></span><span
  style='mso-spacerun:yes'>&#xa0;</span>PAGE<span style='mso-spacerun:yes'>&#xa0;&#xa0;
  </span>\\* MERGEFORMAT <span style='mso-element:field-separator'></span></span><![endif]--><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-no-proof:yes'>ii</span></span><!--[if supportFields]><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-element:field-end'></span></span><![endif]--><span lang=EN-GB
  style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span style='mso-tab-count:
  1'>&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>&#xa9;
  ISO/IEC&#x26;nbsp;2016&#x26;nbsp;&#x2013; All rights reserved<o:p></o:p></span></p>
  </div>
  <div style='mso-element:footer' id=f2>
  <p class=MsoFooter style='line-height:12.0pt'><span lang=EN-GB
  style='font-size:10.0pt;mso-bidi-font-size:11.0pt'>&#xa9; ISO/IEC&#x26;nbsp;2016&#x26;nbsp;&#x2013; All
  rights reserved<span style='mso-tab-count:1'>&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span></span><!--[if supportFields]><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-element:field-begin'></span> PAGE<span style='mso-spacerun:yes'>&#xa0;&#xa0;
  </span>\\* MERGEFORMAT <span style='mso-element:field-separator'></span></span><![endif]--><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-no-proof:yes'>iii</span></span><!--[if supportFields]><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-element:field-end'></span></span><![endif]--><span lang=EN-GB
  style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><o:p></o:p></span></p>
  </div>
  <div style='mso-element:footer' id=ef3>
  <p class=MsoFooter style='margin-top:12.0pt;line-height:12.0pt;mso-line-height-rule:
  exactly'><!--[if supportFields]><b style='mso-bidi-font-weight:normal'><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-element:field-begin'></span><span
  style='mso-spacerun:yes'>&#xa0;</span>PAGE<span style='mso-spacerun:yes'>&#xa0;&#xa0;
  </span>\\* MERGEFORMAT <span style='mso-element:field-separator'></span></span></b><![endif]--><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-no-proof:yes'>2</span></span></b><!--[if supportFields]><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-element:field-end'></span></span></b><![endif]--><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><span
  style='mso-tab-count:1'>&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>&#xa9;
  ISO/IEC&#x26;nbsp;2016&#x26;nbsp;&#x2013; All rights reserved<o:p></o:p></span></p>
  </div>
  <div style='mso-element:footer' id=f3>
  <p class=MsoFooter style='line-height:12.0pt'><span lang=EN-GB
  style='font-size:10.0pt;mso-bidi-font-size:11.0pt'>&#xa9; ISO/IEC&#x26;nbsp;2016&#x26;nbsp;&#x2013; All
  rights reserved<span style='mso-tab-count:1'>&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span></span><!--[if supportFields]><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-element:field-begin'></span>
  PAGE<span style='mso-spacerun:yes'>&#xa0;&#xa0; </span>\\* MERGEFORMAT <span
  style='mso-element:field-separator'></span></span></b><![endif]--><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-no-proof:yes'>3</span></span></b><!--[if supportFields]><b
  style='mso-bidi-font-weight:normal'><span lang=EN-GB style='font-size:10.0pt;
  mso-bidi-font-size:11.0pt'><span style='mso-element:field-end'></span></span></b><![endif]--><span
  lang=EN-GB style='font-size:10.0pt;mso-bidi-font-size:11.0pt'><o:p></o:p></span></p>
  </div>
  </body>
  </html>
FTR

ASCII_MATH = '<m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:grow m:val="on"></m:grow><m:subHide m:val="off"></m:subHide><m:supHide m:val="off"></m:supHide></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSup><m:e><m:r><m:t>i</m:t></m:r></m:e><m:sup><m:r><m:t>3</m:t></m:r></m:sup></m:sSup></m:e></m:nary><m:r><m:t>=</m:t></m:r><m:sSup><m:e><m:d><m:dPr><m:sepChr m:val=","></m:sepChr></m:dPr><m:e><m:f><m:fPr><m:type m:val="bar"></m:type></m:fPr><m:num><m:r><m:t>n</m:t></m:r><m:d><m:dPr><m:sepChr m:val=","></m:sepChr></m:dPr><m:e><m:r><m:t>n+1</m:t></m:r></m:e></m:d></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:e></m:d></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'.freeze

DEFAULT_STYLESHEET = File.read("lib/html2doc/wordstyle.css",
                               encoding: "utf-8").freeze

def guid_clean(xml)
  xml.gsub(/NextPart_[0-9a-f.]+/, "NextPart_")
end

def image_clean(xml)
  xml.gsub(%r{[0-9a-f-]+\.png}, "image.png")
    .gsub(%r{[0-9a-f-]+\.gif}, "image.gif")
    .gsub(%r{[0-9a-f-]+\.(jpeg|jpg)}, "image.jpg")
    .gsub(%r{------=_NextPart_\s+Content-Location: file:///C:/Doc/test_files/image\.(png|gif).*?\s-----=_NextPart_}m, "------=_NextPart_")
    .gsub(%r{Content-Type: image/(png|gif|jpeg)[^-]*------=_NextPart_-?-?}m, "")
    .gsub(%r{ICAgICAg[^-]*-----}m, "-----")
    .gsub(%r{\s*</img>\s*}m, "</img>")
    .gsub(%r{</body>\s*</html>}m, "</body></html>")
end

RSpec.describe Html2Doc do
  it "has a version number" do
    expect(Html2Doc::VERSION).not_to be nil
  end

  it "preserves Word HTML directives" do
    Html2Doc.new(filename: "test").process(html_input(%[A<!--[if gte mso 9]>X<![endif]-->B]))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body(%{A<!--[if gte mso 9]>X<![endif]-->B},
                    '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "processes a blank document" do
    Html2Doc.new(filename: "test").process(html_input(""))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('', '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
      OUTPUT
  end

  it "removes any temp files" do
    File.delete("test.doc")
    Html2Doc.new(filename: "test").process(html_input(""))
    expect(File.exist?("test.doc")).to be true
    expect(File.exist?("test.htm")).to be false
    expect(File.exist?("test_files")).to be false
  end

  it "processes a stylesheet in an HTML document with a title" do
    Html2Doc.new(filename: "test", stylesheet: "lib/html2doc/wordstyle.css")
      .process(html_input(""))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('', '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
      OUTPUT
  end

  it "processes a stylesheet in an HTML document without a title" do
    Html2Doc.new(filename: "test",
                 stylesheet: "lib/html2doc/wordstyle.css")
      .process(html_input_no_title(""))
    expect(guid_clean(File.read("test.doc",
                                encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR.sub('<title>blank</title>', '')}
        #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('', '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
      OUTPUT
  end

  it "processes a stylesheet in an HTML document with an empty head" do
    Html2Doc.new(filename: "test",
                 stylesheet: "lib/html2doc/wordstyle.css")
      .process(html_input_empty_head(""))
    word_hdr_end = WORD_HDR_END
      .sub(%(<meta name="Originator" content="Me"/>\n), "")
      .sub("</style>\n</head>", "</style></head>")
    expect(guid_clean(File.read("test.doc",
                                encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR.sub('<title>blank</title>', '')}
        #{DEFAULT_STYLESHEET}
        #{word_hdr_end}
        #{word_body('', '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
      OUTPUT
  end

  it "processes a header" do
    Html2Doc.new(filename: "test",
                 header_file: "spec/header.html")
      .process(html_input(""))
    html = guid_clean(File.read("test.doc", encoding: "utf-8"))
    hdr = Base64.decode64(
      html
      .sub(%r{^.*Content-Location: file:///C:/Doc/test_files/header.html}, "")
      .sub(%r{^.*Content-Type: text/html charset="utf-8"}m, "")
      .sub(%r{------=_NextPart_--.*$}m, ""),
    ).force_encoding("UTF-8")
    # expect(hdr.gsub(/\xa0/, " ")).to match_fuzzy(HEADERHTML)
    expect(HTMLEntities.new.encode(hdr, :hexadecimal)
      .gsub("&#x3c;", "<").gsub("&#x3e;", ">")
      .gsub("&#x27;", "'").gsub("&#x22;", '"')
      .gsub("&#xd;", "&#xa;").gsub("&#xa;", "\n"))
      .to match_fuzzy(HEADERHTML)
    expect(html.sub(%r{Content-ID: <header.html>.*$}m, ""))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET.gsub(/url\("[^"]+"\)/, 'url(cid:header.html)')}
        #{WORD_HDR_END} #{word_body('', '<div style="mso-element:footnote-list"/>')} #{WORD_FTR2}
      OUTPUT
  end

  it "processes a header with an image" do
    Html2Doc.new(filename: "test",
                 header_file: "spec/header_img.html")
      .process(html_input(""))
    doc = guid_clean(File.read("test.doc", encoding: "utf-8"))
    expect(doc).to match(%r{Content-Type: image/png})
    expect(doc).to match(%r{iVBORw0KGgoAAAANSUhEUgAAA5cAAAN7CAYAAADRE24cAAAgAElEQVR4XuydB5gUxdaGC65gTogB})
  end

  it "processes a header with an image with absolute path" do
    doc = File.read("spec/header_img.html", encoding: "utf-8")
    File.open("spec/header_img1.html", "w:UTF-8") do |f|
      f.write(
        doc.sub(%r{spec/19160-6.png},
                File.expand_path(File.join(File.dirname(__FILE__),
                                           "19160-6.png"))),
      )
    end
    Html2Doc.new(filename: "test",
                 header_file: "spec/header_img1.html")
      .process(html_input(""))
    doc = guid_clean(File.read("test.doc", encoding: "utf-8"))
    expect(doc).to match(%r{Content-Type: image/png})
    expect(doc).to match(%r{iVBORw0KGgoAAAANSUhEUgAAA5cAAAN7CAYAAADRE24cAAAgAElEQVR4XuydB5gUxdaGC65gTogB})
  end

  it "processes a populated document" do
    simple_body = "<h1>Hello word!</h1>
    <div>This is a very simple document</div>"
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body(simple_body, '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "processes mstyle" do
    Html2Doc.new(filename: "test", asciimathdelims: ["{{", "}}"])
      .process(html_input(%[<div><math xmlns="http://www.w3.org/1998/Math/MathML" displaystyle="true">
  <mstyle displaystyle="true">
    <mstyle mathvariant="bold">
      <mrow>
        <mo>&#x2212;</mo>
        <msubsup>
          <mo>&#x222B;</mo>
          <mi>log</mi>
          <mn>2</mn>
        </msubsup>
        <mrow>
          <mo>(</mo>
          <msub>
            <mi>p</mi>
            <mi>u</mi>
          </msub>
          <mo>)</mo>
        </mrow>
      </mrow>
    </mstyle>
    <mstyle mathvariant="bold">
      <mtext>BB</mtext>
    </mstyle>
    <mstyle mathvariant="double-struck">
      <mtext>BBB</mtext>
    </mstyle>
    <mstyle mathvariant="script">
      <mtext>CC</mtext>
    </mstyle>
    <mstyle mathvariant="bold-script">
      <mtext>BCC</mtext>
    </mstyle>
    <mstyle mathvariant="monospace">
      <mtext>TT</mtext>
    </mstyle>
    <mstyle mathvariant="fraktur">
      <mtext>FR</mtext>
    </mstyle>
    <mstyle mathvariant="bold-fraktur">
      <mtext>BFR</mtext>
    </mstyle>
    <mstyle mathvariant="sans-serif">
      <mtext>SF</mtext>
    </mstyle>
    <mstyle mathvariant="bold-sans-serif">
      <mtext>BSFα</mtext>
    </mstyle>
    <mstyle mathvariant="sans-serif-italic">
      <mtext>SFI</mtext>
    </mstyle>
    <mstyle mathvariant="sans-serif-bold-italic">
      <mtext>SFBIα</mtext>
    </mstyle>
    <mstyle mathvariant="bold-italic">
      <mtext>BII</mtext>
    </mstyle>
    <mstyle mathvariant="italic">
      <mtext>II</mtext>
    </mstyle>
  </mstyle>
</math></div>]))
    doc = File.read("test.doc", encoding: "utf-8")
      .sub(%r{^.*<m:oMathPara>}m, "<m:oMathPara>")
      .sub(%r{</m:oMathPara>.*$}m, "</m:oMathPara>")
    expect(doc)
      .to be_equivalent_to(<<~OUTPUT)
        <m:oMathPara>
          <m:oMath>
          <m:r><m:rPr><m:sty m:val="b"></m:sty></m:rPr><m:r><m:t><span style="font-style:normal;font-weight:bold;">&#x2212;</span></m:t></m:r><m:nary><m:naryPr><m:chr m:val="&#x222B;"></m:chr><m:limLoc m:val="subSup"></m:limLoc><m:subHide m:val="0"></m:subHide><m:supHide m:val="0"></m:supHide></m:naryPr><m:sub><m:r><m:t><span style="font-style:normal;font-weight:bold;">log</span></m:t></m:r></m:sub><m:sup><m:r><m:t><span style="font-style:normal;font-weight:bold;">2</span></m:t></m:r></m:sup><m:e><m:r><m:t><span style="font-style:normal;font-weight:bold;">&#x200B;</span></m:t></m:r></m:e></m:nary><m:r><m:t><span style="font-style:normal;font-weight:bold;">(</span></m:t></m:r><m:sSub><m:sSubPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"></w:rFonts><w:i></w:i></w:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><m:t><span style="font-style:normal;font-weight:bold;">p</span></m:t></m:r></m:e><m:sub><m:r><m:t><span style="font-style:normal;font-weight:bold;">u</span></m:t></m:r></m:sub></m:sSub><m:d><m:dPr><m:begChr m:val=")"></m:begChr><m:sepChr m:val=""></m:sepChr></m:dPr><m:e></m:e></m:d></m:r><m:r><m:rPr><m:sty m:val="b"></m:sty></m:rPr><m:t><span style="font-style:normal;font-weight:bold;">BB</span></m:t></m:r><m:r><m:rPr><m:scr m:val="double-struck"></m:scr></m:rPr><m:t>&#x1D539;&#x1D539;&#x1D539;</m:t></m:r><m:r><m:rPr><m:scr m:val="script"></m:scr><m:sty m:val="p"></m:sty></m:rPr><m:t>&#x1D49E;&#x1D49E;</m:t></m:r><m:r><m:rPr><m:scr m:val="script"></m:scr><m:sty m:val="b"></m:sty></m:rPr><m:t>&#x1D4D1;&#x1D4D2;&#x1D4D2;</m:t></m:r><m:r><m:rPr><m:scr m:val="monospace"></m:scr></m:rPr><m:t>&#x1D683;&#x1D683;</m:t></m:r><m:r><m:rPr><m:scr m:val="fraktur"></m:scr><m:sty m:val="p"></m:sty></m:rPr><m:t>&#x1D509;&#x211C;</m:t></m:r><m:r><m:rPr><m:scr m:val="fraktur"></m:scr><m:sty m:val="b"></m:sty></m:rPr><m:t>&#x1D56D;&#x1D571;&#x1D57D;</m:t></m:r><m:r><m:rPr><m:scr m:val="sans-serif"></m:scr><m:sty m:val="p"></m:sty></m:rPr><m:t>&#x1D5B2;&#x1D5A5;</m:t></m:r><m:r><m:rPr><m:scr m:val="sans-serif"></m:scr><m:sty m:val="b"></m:sty></m:rPr><m:t>&#x1D5D5;&#x1D5E6;&#x1D5D9;&#x1D770;</m:t></m:r><m:r><m:rPr><m:scr m:val="sans-serif"></m:scr><m:sty m:val="i"></m:sty></m:rPr><m:t>&#x1D61A;&#x1D60D;&#x1D610;</m:t></m:r><m:r><m:rPr><m:scr m:val="sans-serif"></m:scr><m:sty m:val="bi"></m:sty></m:rPr><m:t>&#x1D64E;&#x1D641;&#x1D63D;&#x1D644;&#x1D7AA;</m:t></m:r><m:r><m:rPr><m:sty m:val="bi"></m:sty></m:rPr><m:t><span class="nostem" style="font-weight:bold;"><em></em>BII</span></m:t></m:r><m:r><m:rPr><m:sty m:val="i"></m:sty></m:rPr><m:t><span class="nostem"><em></em>II</span></m:t></m:r>
          </m:oMath>
        </m:oMathPara>
      OUTPUT
  end

  it "processes spaces in MathML mtext" do
    Html2Doc.new(filename: "test", asciimathdelims: ["{{", "}}"])
      .process(html_input("<div><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='true'>
                                <mrow><mi>H</mi><mtext>&#xa0;original&#xa0;</mtext><mi>J</mi></mrow>
                                </math></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div><m:oMathPara>
                <m:oMath>
                <m:r><m:t>H</m:t></m:r><m:r><m:rPr><m:sty m:val="p"></m:sty></m:rPr><m:t><span style="font-style:normal;">&#xA0;original&#xA0;</span></m:t></m:r><m:r><m:t>J</m:t></m:r>
                </m:oMath>
                </m:oMathPara></div>', '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "processes linebreaks in MathML mtext" do
    Html2Doc.new(filename: "test", asciimathdelims: ["{{", "}}"])
      .process(html_input("<div><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='true'>
                                <mstyle displaystyle='true'>
                                <mi>x</mi><mo>=</mo><mo linebreak='newline'/><mi>y</mi>
                                <mo>=</mo><mo linebreak='newline'/><mi>z</mi>
                                </mstyle>
                                </math></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div><m:oMathPara>
                <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
                <m:r><m:t>x</m:t></m:r><m:r><m:t>=</m:t></m:r>
                </m:oMath>
                <m:r>
                <br/>
                </m:r>
                <m:oMath>
                <m:r><m:t></m:t></m:r><m:r><m:t>y</m:t></m:r><m:r><m:t>=</m:t></m:r>
                </m:oMath>
                <m:r>
                <br/>
                </m:r>
                <m:oMath>
                <m:r><m:t></m:t></m:r><m:r><m:t>z</m:t></m:r>
                </m:oMath>
                </m:oMathPara></div>', '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "unwraps and converts accent in MathML" do
    Html2Doc.new(filename: "test", asciimathdelims: ["{{", "}}"])
      .process(html_input("<div><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='true'>
                                <mover accent='true'><mrow><mi>p</mi></mrow><mrow><mo>^</mo></mrow></mover>
</math></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div><m:oMathPara>
                <m:oMath>
                <m:acc><m:accPr><m:chr m:val="&#x302;"></m:chr></m:accPr><m:e><m:r><m:t>p</m:t></m:r></m:e></m:acc>
                </m:oMath>
                </m:oMathPara></div>', '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "wraps msup after munderover in MathML" do
    Html2Doc.new(filename: "test", asciimathdelims: ["{{", "}}"])
      .process(html_input("<div><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='true'>
<munderover><mo>&#x2211;</mo><mrow><mi>i</mi><mo>=</mo><mn>0</mn></mrow><mrow><mi>n</mi></mrow></munderover><msup><mn>2</mn><mrow><mi>i</mi></mrow></msup></math></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div><m:oMathPara>
                   <m:oMath>
                   <m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:subHide m:val="0"></m:subHide><m:supHide m:val="0"></m:supHide></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>&#x200B;</m:t></m:r></m:e></m:nary><m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"></w:rFonts><w:i></w:i></w:rPr></m:ctrlPr></m:sSupPr><m:e><m:r><m:t>2</m:t></m:r></m:e><m:sup><m:r><m:t>i</m:t></m:r></m:sup></m:sSup>
                  </m:oMath>
                  </m:oMathPara></div>', '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "uncenters paras with just inline maths" do
    Html2Doc.new(filename: "test")
      .process(html_input("<div><p><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p class="MsoNormal"><m:oMathPara>
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><p style='vertical-align:center;text-align:right;'><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p style="vertical-align:center;text-align:right;" class="MsoNormal"><m:oMathPara>
        <m:oMathParaPr><m:jc m:val="right"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><p style='vertical-align:center;text-align:left;'><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p style="vertical-align:center;text-align:left;" class="MsoNormal"><m:oMathPara>
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><table><tr><th><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></tr></th></table></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><table><tr><th><m:oMathPara>
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></th></tr></table></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><p><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='true'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p class="MsoNormal"><m:oMathPara>
        <m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><p><i>a</i><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p class="MsoNormal"><i>a</i><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><p><span>     </span><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p class="MsoNormal"><span> </span><m:oMathPara>
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT

    Html2Doc.new(filename: "test")
      .process(html_input("<div><p><math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math><br/>
                          <math xmlns='http://www.w3.org/1998/Math/MathML' displaystyle='false'><mi>i</mi><mo>=</mo><mn>0</mn></math></p></div>"))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
                #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
                #{word_body('<div><p class="MsoNormal"><m:oMathPara>
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara><br/>
        <m:oMathPara>
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
        <m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>0</m:t></m:r>
        </m:oMath>
        </m:oMathPara></p></div>',
                            '<div style="mso-element:footnote-list"/>')}
                #{WORD_FTR1}
      OUTPUT
  end

  it "processes tabs" do
    simple_body = "<h1>Hello word!</h1>
    <div>This is a very &tab; simple document</div>"
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body(simple_body.gsub('&tab;', %[<span style="mso-tab-count:1">&#xA0; </span>]), '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "makes unstyled paragraphs be MsoNormal" do
    simple_body = '<h1>Hello word!</h1>
    <p>This is a very simple document</p>
    <p class="x">This style stays</p>'
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body(simple_body.gsub('<p>', %[<p class="MsoNormal">]), '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "makes unstyled list entries be MsoNormal" do
    simple_body = '<h1>Hello word!</h1>
    <ul>
    <li>This is a very simple document</li>
    <li class="x">This style stays</li>
    </ul>'
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body(simple_body.gsub('<li>', %[<li class="MsoNormal">]), '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "resizes images for height, in a file in a subdirectory" do
    simple_body = '<img src="19160-6.png">'
    Html2Doc.new(filename: "spec/test", imagedir: "spec",
                 stylesheet: "spec/wordstyle-nopagesize.css")
      .process(html_input(simple_body))
    testdoc = File.read("spec/test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/png})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body('<img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.png" width="400" height="388"></img>', '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  it "resizes images for width" do
    simple_body = '<img src="spec/19160-7.gif">'
    Html2Doc.new(filename: "test", imagedir: ".",
                 stylesheet: "spec/wordstyle-nopagesize.css")
      .process(html_input(simple_body))
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/gif})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body('<img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.gif" width="400" height="118"></img>', '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  it "resizes images for height" do
    simple_body = '<img src="spec/19160-8.jpg">'
    Html2Doc.new(filename: "test", imagedir: ".",
                 stylesheet: "spec/wordstyle-nopagesize.css")
      .process(html_input(simple_body))
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/jpeg})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body('<img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.jpg" width="208" height="680"></img>', '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  it "resizes images for width with custom page size" do
    simple_body = '<img src="spec/19160-7.gif">'
    Html2Doc.new(filename: "test", imagedir: ".",
                 stylesheet: "spec/wordstyle-custom.css")
      .process(html_input(simple_body))
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/gif})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body('<img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.gif" width="400" height="118"></img>', '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  it "resizes images for height with custom page size" do
    simple_body = '<img src="spec/19160-8.jpg">'
    Html2Doc.new(filename: "test", imagedir: ".",
                 stylesheet: "spec/wordstyle-custom.css")
      .process(html_input(simple_body))
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/jpeg})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body('<img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.jpg" width="208" height="680"></img>', '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  it "resizes images for height with landscape page" do
    simple_body = <<~DIV
      <div class="WordSection1"><div class="figure"><img src="spec/19160-8.jpg"></div></div>
      <div class="WordSectionA"><div class="figure"><img src="spec/19160-8.jpg"></div></div>
    DIV
    Html2Doc.new(filename: "test", imagedir: ".",
                 stylesheet: "spec/wordstyle-custom.css")
      .process(html_input(simple_body))
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/jpeg})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body(%{<div class="WordSection1"><div class="figure"><img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.jpg" width="208" height="680"/></div>\n<div class="WordSectionA"><div class="figure"><img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.jpg" width="122" height="400"/></div>\n</div></div>}, '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  it "does not move images if they are external URLs" do
    simple_body = '<img src="https://example.com/19160-6.png">'
    Html2Doc.new(filename: "test", imagedir: ".")
      .process(html_input(simple_body))
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
      #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
      #{image_clean(word_body('<img src="https://example.com/19160-6.png"></img>', '<div style="mso-element:footnote-list"/>'))}
      #{image_clean(WORD_FTR1)}
    OUTPUT
  end

  it "deals with absolute image locations" do
    simple_body = %{<img src="#{__dir__}/19160-6.png">}
    Html2Doc.new(filename: "spec/test", imagedir: ".",
                 stylesheet: "spec/wordstyle-nopagesize.css")
      .process(html_input(simple_body))
    testdoc = File.read("spec/test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/png})
    expect(image_clean(guid_clean(testdoc))
      .sub(/^.*<body/m, "<body")
      .sub(%r{</body>.*$}m, "</body>")).to match_fuzzy(<<~OUTPUT)
        #{image_clean(word_body('<img src="cid:cb7b0d19-891e-4634-815a-570d019d454c.png" width="400" height="388"></img>', '<div style="mso-element:footnote-list"/>')).sub(%r{</html>}, '')}
      OUTPUT
  end

  #   it "warns about SVG" do
  #     simple_body = '<img src="https://example.com/19160-6.svg">'
  #     expect{ Html2Doc.process(html_input(simple_body), filename: "test") }
  #     .to output("https://example.com/19160-6.svg: SVG not supported\n").to_stderr
  #   end

  it "processes epub:type footnotes" do
    simple_body = '<div>This is a very simple
     document<a epub:type="footnote" href="#a_632">1</a> allegedly<a epub:type="footnote" href="#a_782">2</a></div>
     <aside id="a_632">Footnote</aside>
     <aside id="a_782">Other Footnote</aside>'
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
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
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
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

  it "processes footnotes with text wrapping the footnote reference" do
    simple_body = '<div>This is a very simple
     document<a class="footnote" href="#a1">(<span class="MsoFootnoteReference">1</span>)</a> allegedly<a class="footnote" href="#a2">2</a></div>
     <aside id="a1">Footnote</aside>
     <aside id="a2">Other Footnote</aside>'
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div>This is a very simple
                    document<a class="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference">(</span><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span><span class="MsoFootnoteReference">)</span></a> allegedly<a class="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
                    '<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
                <p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference">(</span><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span><span class="MsoFootnoteReference">)</span></a>Footnote</p></div>
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
    Html2Doc.new(filename: "test").process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
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

  it "labels lists with list styles" do
    simple_body = <<~BODY
      <div><ul id="0">
      <li><div><p><ol id="1"><li><ul id="2"><li><p><ol id="3"><li><ol id="4"><li>A</li><li><p>B</p><p>B2</p></li><li>C</li></ol></li></ol></p></li></ul></li></ol></p></div></li><div><ul id="5"><li>C</li></ul></div>
    BODY
    Html2Doc.new(filename: "test", liststyles: { ul: "l1", ol: "l2" })
      .process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div>
                <p style="mso-list:l1 level1 lfo1;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpFirst">A</p><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpMiddle">B<p class="MsoListParagraphCxSpMiddle">B2</p></p><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpLast">C</p></p></p></p></div></p><div><p style="mso-list:l1 level1 lfo2;" class="MsoListParagraphCxSpFirst">C</p></div>
                </div>',
                    '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "restarts numbering of lists with list styles" do
    simple_body = <<~BODY
      <div>
      <ol id="1"><li><div><p><ol id="2"><li><ul id="3"><li><p><ol id="4"><li><ol id="5"><li>A</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ol>
      <ol id="6"><li><div><p><ol id="7"><li><ul id="8"><li><p><ol id="9"><li><ol id="10"><li>A</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ol></div>
    BODY
    Html2Doc.new(filename: "test", liststyles: { ul: "l1", ol: "l2" })
      .process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div>
                  <p style="mso-list:l2 level1 lfo1;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpFirst">A</p></p></p></p></div></p>
                  <p style="mso-list:l2 level1 lfo2;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo2;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo2;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo2;" class="MsoListParagraphCxSpFirst">A</p></p></p></p></div></p></div>',
                    '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "labels lists with multiple list styles" do
    simple_body = <<~BODY
      <div><ul class="steps" id="0">
      <li><div><p><ol id="1"><li><ul id="2"><li><p><ol id="3"><li><ol id="4"><li>A</li><li><p>B</p><p>B2</p></li><li>C</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ul></div>
      <div><ul id="5">
      <li><div><p><ol id="6"><li><ul id="7"><li><p><ol id="8"><li><ol id="9"><li>A</li><li><p>B</p><p>B2</p></li><li>C</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ul></div>
      <div><ul class="other" id="10">
      <li><div><p><ol id="11"><li><ul id="12"><li><p><ol id="13"><li><ol id="14"><li>A</li><li><p>B</p><p>B2</p></li><li>C</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ul></div>
    BODY
    Html2Doc.new(filename: "test",
                 liststyles: { ul: "l1", ol: "l2",
                               steps: "l3" })
      .process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div>
                <p style="mso-list:l3 level1 lfo2;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l3 level2 lfo2;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l3 level4 lfo2;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l3 level5 lfo2;" class="MsoListParagraphCxSpFirst">A</p><p style="mso-list:l3 level5 lfo2;" class="MsoListParagraphCxSpMiddle">B<p class="MsoListParagraphCxSpMiddle">B2</p></p><p style="mso-list:l3 level5 lfo2;" class="MsoListParagraphCxSpLast">C</p></p></p></p></div></p></div>
                <div>
                <p style="mso-list:l1 level1 lfo1;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpFirst">A</p><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpMiddle">B<p class="MsoListParagraphCxSpMiddle">B2</p></p><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpLast">C</p></p></p></p></div></p></div>
                <div>
                <p style="mso-list:l1 level1 lfo3;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo3;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo3;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo3;" class="MsoListParagraphCxSpFirst">A</p><p style="mso-list:l2 level5 lfo3;" class="MsoListParagraphCxSpMiddle">B<p class="MsoListParagraphCxSpMiddle">B2</p></p><p style="mso-list:l2 level5 lfo3;" class="MsoListParagraphCxSpLast">C</p></p></p></p></div></p></div>',
                    '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "replaces id attributes with explicit a@name bookmarks" do
    simple_body = <<~BODY
      <div>
        <p id="a">Hello</p>
        <p id="b"/>
      </div>
    BODY
    Html2Doc.new(filename: "test", liststyles: { ul: "l1", ol: "l2" })
      .process(html_input(simple_body))
    expect(guid_clean(File.read("test.doc", encoding: "utf-8")))
      .to match_fuzzy(<<~OUTPUT)
        #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
        #{word_body('<div>
                    <p class="MsoNormal"><a name="a" id="a"></a>Hello</p>
                    <p class="MsoNormal"><a name="b" id="b"></a></p>
                  </div>',
                    '<div style="mso-element:footnote-list"/>')}
        #{WORD_FTR1}
      OUTPUT
  end

  it "test image base64 image encoding" do
    simple_body = '<img src="19160-6.png">'
    Html2Doc.new(filename: "spec/test", debug: true, imagedir: "spec")
      .process(html_input(simple_body))
    testdoc = File.read("spec/test.doc", encoding: "utf-8")
    base64_image = testdoc[/image\/png\n\n(.*?)\n\n----/m, 1].gsub!("\n", "")
    base64_image_basename = testdoc[%r{Content-ID: <([0-9a-z\-]+)\.png}m, 1]
    doc_bin_image = Base64.strict_decode64(base64_image)
    file_bin_image = IO
      .read("spec/test_files/#{base64_image_basename}.png", mode: "rb")
    expect(doc_bin_image).to eq file_bin_image
    FileUtils.rm_rf %w[spec/test_files spec/test.doc spec/test.htm]
  end
end
