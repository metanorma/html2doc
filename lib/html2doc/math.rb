require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"
require "plane1converter"

module Html2Doc
  @xsltemplate =
    Nokogiri::XSLT(File.read(File.join(File.dirname(__FILE__), "mml2omml.xsl"),
                             encoding: "utf-8"))

  def self.asciimath_to_mathml1(expr)
    AsciiMath::MathMLBuilder.new(msword: true).append_expression(
      AsciiMath.parse(HTMLEntities.new.decode(expr)).ast,
    ).to_s
      .gsub(/<math>/, "<math xmlns='http://www.w3.org/1998/Math/MathML'>")
  rescue StandardError => e
    puts "parsing: #{expr}"
    puts e.message
    raise e
  end

  def self.asciimath_to_mathml(doc, delims)
    return doc if delims.nil? || delims.size < 2

    m = doc.split(/(#{Regexp.escape(delims[0])}|#{Regexp.escape(delims[1])})/)
    m.each_slice(4).map.with_index do |(*a), i|
      i % 500 == 0 && m.size > 1000 && i > 0 and
        warn "MathML #{i} of #{(m.size / 4).floor}"
      a[2].nil? || a[2] = asciimath_to_mathml1(a[2])
      a.size > 1 ? a[0] + a[2] : a[0]
    end.join
  end

  def self.unwrap_accents(doc)
    doc.xpath("//*[@accent = 'true']").each do |x|
      x.elements.length > 1 or next
      x.elements[1].name == "mrow" and
        x.elements[1].replace(x.elements[1].children)
    end
    doc
  end

  # random fixes to MathML input that OOXML needs to render properly
  def self.ooxml_cleanup(math, docnamespaces)
    math = unwrap_accents(
      mathml_preserve_space(
        mathml_insert_rows(math, docnamespaces), docnamespaces
      ),
    )
    math.add_namespace(nil, "http://www.w3.org/1998/Math/MathML")
    math
  end

  def self.mathml_insert_rows(math, docnamespaces)
    math.xpath(%w(msup msub msubsup munder mover munderover)
            .map { |m| ".//xmlns:#{m}" }.join(" | "), docnamespaces).each do |x|
      next unless x.next_element && x.next_element != "mrow"

      x.next_element.wrap("<mrow/>")
    end
    math
  end

  def self.mathml_preserve_space(math, docnamespaces)
    math.xpath(".//xmlns:mtext", docnamespaces).each do |x|
      x.children = x.children.to_xml.gsub(/^\s/, "&#xA0;").gsub(/\s$/, "&#xA0;")
    end
    math
  end

  HTML_NS = 'xmlns="http://www.w3.org/1999/xhtml"'.freeze

  def self.unitalic(math)
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:scr)]/xmlns:sty[@m:val = 'p']]").each do |x|
      x.wrap("<span #{HTML_NS} style='font-style:normal;'></span>")
    end
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:scr)]/xmlns:sty[@m:val = 'bi']]").each do |x|
      x.wrap("<span #{HTML_NS} class='nostem' style='font-weight:bold;'><em></em></span>")
    end
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:scr)]/xmlns:sty[@m:val = 'i']]").each do |x|
      x.wrap("<span #{HTML_NS} class='nostem'><em></em></span>")
    end
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:scr)]/xmlns:sty[@m:val = 'b']]").each do |x|
      x.wrap("<span #{HTML_NS} style='font-style:normal;font-weight:bold;'></span>")
    end
    math.xpath(".//xmlns:r[xmlns:rPr/xmlns:scr[@m:val = 'monospace']]").each do |x|
      to_plane1(x, :monospace)
    end
    math.xpath(".//xmlns:r[xmlns:rPr/xmlns:scr[@m:val = 'double-struck']]").each do |x|
      to_plane1(x, :doublestruck)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:sty) or xmlns:sty/@m:val = 'p']/xmlns:scr[@m:val = 'script']]").each do |x|
      to_plane1(x, :script)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[xmlns:sty/@m:val = 'b']/xmlns:scr[@m:val = 'script']]").each do |x|
      to_plane1(x, :scriptbold)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:sty) or xmlns:sty/@m:val = 'p']/xmlns:scr[@m:val = 'fraktur']]").each do |x|
      to_plane1(x, :fraktur)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[xmlns:sty/@m:val = 'b']/xmlns:scr[@m:val = 'fraktur']]").each do |x|
      to_plane1(x, :frakturbold)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[not(xmlns:sty) or xmlns:sty/@m:val = 'p']/xmlns:scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sans)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[xmlns:sty/@m:val = 'b']/xmlns:scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansbold)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[xmlns:sty/@m:val = 'i']/xmlns:scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansitalic)
    end
    math.xpath(".//xmlns:r[xmlns:rPr[xmlns:sty/@m:val = 'bi']/xmlns:scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansbolditalic)
    end
    math
  end

  def self.to_plane1(xml, font)
    xml.traverse do |n|
      next unless n.text?

      n.replace(Plane1Converter.conv(HTMLEntities.new.decode(n.text), font))
    end
    xml
  end

  def self.mathml_to_ooml(docxml)
    docnamespaces = docxml.collect_namespaces
    m = docxml.xpath("//*[local-name() = 'math']")
    m.each_with_index do |x, i|
      i % 100 == 0 && m.size > 500 && i > 0 and
        warn "Math OOXML #{i} of #{m.size}"
      element = ooxml_cleanup(x, docnamespaces)
      doc = Nokogiri::XML::Document::new
      doc.root = element
      ooxml = ooml_clean(unitalic(esc_space(@xsltemplate.transform(doc))))
      ooxml = uncenter(x, ooxml)
      x.swap(ooxml)
    end
  end

  # We need span and em not to be namespaced. Word can't deal with explicit 
  # namespaces.
  # We will end up stripping them out again under Nokogiri 1.11, which correctly
  # insists on inheriting namespace from parent.
  def self.ooml_clean(xml)
    xml.to_s
      .gsub(/<\?[^>]+>\s*/, "")
      .gsub(/ xmlns(:[^=]+)?="[^"]+"/, "")
      .gsub(%r{<(/)?(?!span)(?!em)([a-z])}, "<\\1m:\\2")
  end

  # escape space as &#x32;; we are removing any spaces generated by
  # XML indentation
  def self.esc_space(xml)
    xml.traverse do |n|
      next unless n.text?

      n = n.text.gsub(/ /, "&#x32;")
    end
    xml
  end

  # if oomml has no siblings, by default it is centered; override this with
  # left/right if parent is so tagged
  def self.uncenter(math, ooxml)
    alignnode = math.at(".//ancestor::*[@style][local-name() = 'p' or "\
                        "local-name() = 'div' or local-name() = 'td']/@style")
    return ooxml unless alignnode && (math.next == nil && math.previous == nil)

    %w(left right).each do |dir|
      if alignnode.text.include? ("text-align:#{dir}")
        ooxml = "<m:oMathPara><m:oMathParaPr><m:jc "\
          "m:val='#{dir}'/></m:oMathParaPr>#{ooxml}</m:oMathPara>"
      end
    end
    ooxml
  end
end
