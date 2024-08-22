class Html2Doc
  NOKOHEAD = <<~HERE.freeze
    <!DOCTYPE html SYSTEM
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head> <title></title> <meta charset="UTF-8" /> </head>
    <body> </body> </html>
  HERE

  def to_xhtml(xml)
    xml.gsub!(/<\?xml[^<>]*>/, "")
    unless /<!DOCTYPE /.match? xml
      xml = '<!DOCTYPE html SYSTEM
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">' + xml
    end
    xml = xml.gsub(/<!--\s*\[([^\<\]]+)\]>/, "<!-- MSWORD-COMMENT \\1 -->")
      .gsub(/<!\s*\[endif\]\s*-->/, "<!-- MSWORD-COMMENT-END -->")
    Nokogiri::XML.parse(xml)
  end

  DOCTYPE = <<~DOCTYPE.freeze
    <!DOCTYPE html SYSTEM "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
  DOCTYPE

  def from_xhtml(xml)
    xml.to_xml.sub(%{ xmlns="http://www.w3.org/1999/xhtml"}, "")
      .sub(DOCTYPE, "").gsub(%{ />}, "/>")
      .gsub(/<!-- MSWORD-COMMENT (.+?) -->/, "<!--[\\1]>")
      .gsub(/<!-- MSWORD-COMMENT-END -->/, "<![endif]-->")
      .gsub("\n--&gt;\n", "\n-->\n")
  end

  def msword_fix(doc)
    # brain damage in MSWord parser
    doc.gsub!(%r{<w:DoNotOptimizeForBrowser></w:DoNotOptimizeForBrowser>},
              "<w:DoNotOptimizeForBrowser/>")
    doc.gsub!(%r{<span style="mso-special-character:footnote"/>},
              '<span style="mso-special-character:footnote"></span>')
    doc.gsub!(%r{<div style="mso-element:footnote-list"></div>},
              '<div style="mso-element:footnote-list"/>')
    doc.gsub!(%r{(<a style="mso-comment-reference:[^<>/]+)/>}, "\\1></a>")
    doc.gsub!(%r{<link rel="File-List"}, "<link rel=File-List")
    doc.gsub!(%r{<meta http-equiv="Content-Type"},
              "<meta http-equiv=Content-Type")
    doc.gsub!(%r{></m:jc>}, "/>")
    doc.gsub!(%r{></v:stroke>}, "/>")
    doc.gsub!(%r{></v:f>}, "/>")
    doc.gsub!(%r{></v:path>}, "/>")
    doc.gsub!(%r{></o:lock>}, "/>")
    doc.gsub!(%r{></v:imagedata>}, "/>")
    doc.gsub!(%r{></w:wrap>}, "/>")
    doc.gsub!(%r{<(/)?m:(span|em)\b}, "<\\1\\2")
    doc.gsub!(%r{&tab;|&amp;tab;},
              '<span style="mso-tab-count:1">&#xA0; </span>')
    doc.split(%r{(<m:oMath>|</m:oMath>)}).each_slice(4).map do |a|
      a.size > 2 and a[2] = a[2].gsub(/>\s+</, "><")
      a
    end.join
  end

  PRINT_VIEW = <<~XML.freeze

    <xml>
    <w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
    </xml>
    <meta http-equiv='Content-Type' content="text/html; charset=utf-8"/>
  XML

  def namespace(root)
    { o: "urn:schemas-microsoft-com:office:office",
      w: "urn:schemas-microsoft-com:office:word",
      v: "urn:schemas-microsoft-com:vml",
      m: "http://schemas.microsoft.com/office/2004/12/omml" }.each { |k, v| root.add_namespace_definition(k.to_s, v) }
  end

  def rootnamespace(root)
    root.add_namespace(nil, "http://www.w3.org/TR/REC-html40")
  end
end
