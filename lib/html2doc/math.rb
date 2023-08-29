require "uuidtools"
require "plurimath"
require "htmlentities"
require "nokogiri"
require "plane1converter"
require "metanorma-utils"

module Nokogiri
  module XML
    class Node
      OOXML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/".freeze

      def ooxml_xpath(path)
        p = Metanorma::Utils::ns(path).gsub("xmlns:", "m:")
        xpath(p, "m" => OOXML_NS)
      end
    end
  end
end

class Html2Doc
  def progress_conv(idx, step, total, threshold, msg)
    return unless (idx % step).zero? && total > threshold && idx.positive?

    warn "#{msg} #{idx} of #{total}"
  end

  def unwrap_accents(doc)
    doc.xpath("//*[@accent = 'true']").each do |x|
      x.elements.length > 1 or next
      x.elements[1].name == "mrow" and
        x.elements[1].replace(x.elements[1].children)
    end
    doc
  end

  MATHML_NS = "http://www.w3.org/1998/Math/MathML".freeze

  # random fixes to MathML input that OOXML needs to render properly
  def ooxml_cleanup(math, docnamespaces)
    unwrap_accents(
      mathml_preserve_space(
        mathml_insert_rows(math, docnamespaces), docnamespaces
      ),
    )
    math.add_namespace(nil, MATHML_NS)
    math
  end

  def mathml_insert_rows(math, docnamespaces)
    math.xpath(%w(msup msub msubsup munder mover munderover)
            .map { |m| ".//xmlns:#{m}" }.join(" | "), docnamespaces).each do |x|
      next unless x.next_element && x.next_element != "mrow"

      x.next_element.wrap("<mrow/>")
    end
    math
  end

  def mathml_preserve_space(math, docnamespaces)
    math.xpath(".//xmlns:mtext", docnamespaces).each do |x|
      x.children = x.children.to_xml.gsub(/^\s/, "&#xA0;").gsub(/\s$/, "&#xA0;")
    end
    math
  end

  HTML_NS = 'xmlns="http://www.w3.org/1999/xhtml"'.freeze

  def unitalic(math)
    math.ooxml_xpath(".//r[rPr[not(scr)]/sty[@m:val = 'p']]").each do |x|
      x.wrap("<span #{HTML_NS} style='font-style:normal;'></span>")
    end
    math.ooxml_xpath(".//r[rPr[not(scr)]/sty[@m:val = 'bi']]").each do |x|
      x.wrap("<span #{HTML_NS} class='nostem' style='font-weight:bold;'><em></em></span>")
    end
    math.ooxml_xpath(".//r[rPr[not(scr)]/sty[@m:val = 'i']]").each do |x|
      x.wrap("<span #{HTML_NS} class='nostem'><em></em></span>")
    end
    math.ooxml_xpath(".//r[rPr[not(scr)]/sty[@m:val = 'b']]").each do |x|
      x.wrap("<span #{HTML_NS} style='font-style:normal;font-weight:bold;'></span>")
    end
    math.ooxml_xpath(".//r[rPr/scr[@m:val = 'monospace']]").each do |x|
      to_plane1(x, :monospace)
    end
    math.ooxml_xpath(".//r[rPr/scr[@m:val = 'double-struck']]").each do |x|
      to_plane1(x, :doublestruck)
    end
    math.ooxml_xpath(".//r[rPr[not(sty) or sty/@m:val = 'p']/scr[@m:val = 'script']]").each do |x|
      to_plane1(x, :script)
    end
    math.ooxml_xpath(".//r[rPr[sty/@m:val = 'b']/scr[@m:val = 'script']]").each do |x|
      to_plane1(x, :scriptbold)
    end
    math.ooxml_xpath(".//r[rPr[not(sty) or sty/@m:val = 'p']/scr[@m:val = 'fraktur']]").each do |x|
      to_plane1(x, :fraktur)
    end
    math.ooxml_xpath(".//r[rPr[sty/@m:val = 'b']/scr[@m:val = 'fraktur']]").each do |x|
      to_plane1(x, :frakturbold)
    end
    math.ooxml_xpath(".//r[rPr[not(sty) or sty/@m:val = 'p']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sans)
    end
    math.ooxml_xpath(".//r[rPr[sty/@m:val = 'b']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansbold)
    end
    math.ooxml_xpath(".//r[rPr[sty/@m:val = 'i']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansitalic)
    end
    math.ooxml_xpath(".//r[rPr[sty/@m:val = 'bi']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansbolditalic)
    end
    math
  end

  def to_plane1(xml, font)
    xml.traverse do |n|
      next unless n.text?

      n.replace(Plane1Converter.conv(@c.decode(n.text), font))
    end
    xml
  end

  def mathml_to_ooml(docxml)
    docnamespaces = docxml.collect_namespaces
    m = docxml.xpath("//*[local-name() = 'math']")
    m.each_with_index do |x, i|
      progress_conv(i, 100, m.size, 500, "Math OOXML")
      mathml_to_ooml1(x, docnamespaces)
    end
  end

  # We need span and em not to be namespaced. Word can't deal with explicit
  # namespaces.
  # We will end up stripping them out again under Nokogiri 1.11, which correctly
  # insists on inheriting namespace from parent.
  def ooml_clean(xml)
    xml.to_s
      .gsub(/<\?[^>]+>\s*/, "")
      .gsub(/ xmlns(:[^=]+)?="[^"]+"/, "")
      .gsub(/<\/?m:oMathPara>/, "")
      #.gsub(%r{<(/)?(?!span)(?!em)([a-z])}, "<\\1m:\\2")
  end

  def mathml_to_ooml1(xml, docnamespaces)
    doc = Nokogiri::XML::Document::new
    doc.root = ooxml_cleanup(xml, docnamespaces)
    # ooxml = @xsltemplate.transform(doc)
    d = xml.parent["block"] != "false" # display_style
    ooxml = Nokogiri::XML(Plurimath::Math.parse(doc.to_xml, :mathml)
      .to_omml)
    ooxml = ooml_clean(unitalic(esc_space(accent_tr(ooxml))))
    ooxml = uncenter(xml, ooxml)
    xml.swap(ooxml)
  end

  def accent_tr(xml)
    xml.ooxml_xpath(".//accPr/chr").each do |x|
      x["m:val"] &&= accent_tr1(x["m:val"])
      x["val"] &&= accent_tr1(x["val"])
    end
    xml
  end

  def accent_tr1(accent)
    case accent
    when "\u2192" then "\u20D7"
    when "^" then "\u0302"
    when "~" then "\u0303"
    else accent
    end
  end

  # escape space as &#x32;; we are removing any spaces generated by
  # XML indentation
  def esc_space(xml)
    xml.traverse do |n|
      next unless n.text?

      n = n.text.gsub(/ /, "&#x32;")
    end
    xml
  end

  OOXML_NS = "http://schemas.microsoft.com/office/2004/12/omml".freeze

  def math_only_para?(node)
    x = node.dup
    x.xpath(".//m:math", "m" => MATHML_NS).each(&:remove)
    x.xpath(".//m:oMathPara | .//m:oMath", "m" => OOXML_NS).each(&:remove)
    x.text.strip.empty?
  end

  def math_block?(ooxml, mathml)
    ooxml.name == "oMathPara" || mathml["displaystyle"] == "true"
  end

  STYLE_BEARING_NODE =
    %w(p div td th li).map { |x| ".//ancestor::#{x}" }.join(" | ").freeze

  # if oomml has no siblings, by default it is centered; override this with
  # left/right if parent is so tagged
  # also if ooml has mathPara already, or is in para with only oMath content
  def uncenter(math, ooxml)
    alignnode = math.xpath(STYLE_BEARING_NODE).last
    ret = ooxml.root.to_xml(indent: 0)
    (math_block?(ooxml, math) ||
      !alignnode) || !math_only_para?(alignnode) and return ret
    dir = "left"
    alignnode["style"]&.include?("text-align:right") and dir = "right"
    "<oMathPara><oMathParaPr><jc " \
      "m:val='#{dir}'/></oMathParaPr>#{ret}</oMathPara>"
  end
end
