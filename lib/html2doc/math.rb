require "uuidtools"
require "plurimath"
require "htmlentities"
require "nokogiri"
require "plane1converter"
require "metanorma-utils"

module Nokogiri
  module XML
    class Node
      OOXML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math".freeze

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
    # encode_math(
    unwrap_accents(
      mathml_preserve_space(
        mathml_insert_rows(math, docnamespaces), docnamespaces
      ),
    )
    # )
    math.add_namespace(nil, MATHML_NS)
    math
  end

  def encode_math(elem)
    elem.traverse do |e|
      e.text? or next
      e.text.strip.empty? and next
      e.replace(@c.encode(e.text, :hexadecimal))
    end
    elem
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

  def wrap_text(elem, wrapper)
    elem.traverse do |e|
      e.text? or next
      e.text.strip.empty? and next
      e.wrap(wrapper)
    end
  end

  def unitalic(math)
    math.ooxml_xpath(".//r[rPr[not(m:scr)]/sty[@m:val = 'p']]").each do |x|
      wrap_text(x, "<span #{HTML_NS} style='font-style:normal;'></span>")
    end
    math.ooxml_xpath(".//r[rPr[not(m:scr)]/sty[@m:val = 'bi']]").each do |x|
      wrap_text(x,
                "<span #{HTML_NS} class='nostem' style='font-weight:bold;'><em></em></span>")
    end
    math.ooxml_xpath(".//r[rPr[not(m:scr)]/sty[@m:val = 'i']]").each do |x|
      wrap_text(x, "<span #{HTML_NS} class='nostem'><em></em></span>")
    end
    math.ooxml_xpath(".//r[rPr[not(m:scr)]/sty[@m:val = 'b']]").each do |x|
      wrap_text(x,
                "<span #{HTML_NS} style='font-style:normal;font-weight:bold;'></span>")
    end
    math.ooxml_xpath(".//r[rPr/scr[@m:val = 'monospace']]").each do |x|
      to_plane1(x, :monospace)
    end
    math.ooxml_xpath(".//r[rPr/scr[@m:val = 'double-struck']]").each do |x|
      to_plane1(x, :doublestruck)
    end
    math.ooxml_xpath(".//r[rPr[not(m:sty) or m:sty/@m:val = 'p']/scr[@m:val = 'script']]").each do |x|
      to_plane1(x, :script)
    end
    math.ooxml_xpath(".//r[rPr[m:sty/@m:val = 'b']/scr[@m:val = 'script']]").each do |x|
      to_plane1(x, :scriptbold)
    end
    math.ooxml_xpath(".//r[rPr[not(m:sty) or m:sty/@m:val = 'p']/scr[@m:val = 'fraktur']]").each do |x|
      to_plane1(x, :fraktur)
    end
    math.ooxml_xpath(".//r[rPr[m:sty/@m:val = 'b']/scr[@m:val = 'fraktur']]").each do |x|
      to_plane1(x, :frakturbold)
    end
    math.ooxml_xpath(".//r[rPr[not(m:sty) or m:sty/@m:val = 'p']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sans)
    end
    math.ooxml_xpath(".//r[rPr[m:sty/@m:val = 'b']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansbold)
    end
    math.ooxml_xpath(".//r[rPr[m:sty/@m:val = 'i']/scr[@m:val = 'sans-serif']]").each do |x|
      to_plane1(x, :sansitalic)
    end
    math.ooxml_xpath(".//r[rPr[m:sty/@m:val = 'bi']/scr[@m:val = 'sans-serif']]").each do |x|
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
    xml.to_xml(indent: 0)
      .gsub(/<\?[^>]+>\s*/, "")
      .gsub(/ xmlns(:[^=]+)?="[^"]+"/, "")
    # .gsub(%r{<(/)?(?!span)(?!em)([a-z])}, "<\\1m:\\2")
  end

  def mathml_to_ooml1(xml, docnamespaces)
    doc = Nokogiri::XML::Document::new
    doc.root = ooxml_cleanup(xml, docnamespaces)
    # d = xml.parent["block"] != "false" # display_style
    ooxml = Nokogiri::XML(Plurimath::Math
      .parse(doc.root.to_xml(indent: 0), :mathml)
      .to_omml(split_on_linebreak: true))
    ooxml = unitalic(accent_tr(ooxml))
    ooxml = ooml_clean(uncenter(xml, ooxml))
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

  OOXML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math".freeze

  def math_only_para?(node)
    x = node.dup
    x.xpath(".//m:math", "m" => MATHML_NS).each(&:remove)
    x.xpath(".//m:oMathPara | .//m:oMath", "m" => OOXML_NS).each(&:remove)
    x.xpath(".//m:oMathPara | .//m:oMath").each(&:remove)
    # namespace can go missing during processing
    x.text.strip.empty?
  end

  def math_block?(ooxml, mathml)
    # ooxml.name == "oMathPara" || mathml["displaystyle"] == "true"
    mathml["displaystyle"] == "true" &&
      ooxml.xpath("./m:oMath", "m" => OOXML_NS).size <= 1
  end

  STYLE_BEARING_NODE =
    %w(p div td th li).map { |x| ".//ancestor::#{x}" }.join(" | ").freeze

  # if oomml has no siblings, by default it is centered; override this with
  # left/right if parent is so tagged
  # also if ooml has mathPara already, or is in para with only oMath content
  def uncenter(math, ooxml)
    alignnode = math.xpath(STYLE_BEARING_NODE).last
    ooxml.document? and ooxml = ooxml.root
    ret = uncenter_unneeded(math, ooxml, alignnode) and return ret
    dir = ooxml_alignment(alignnode)
    ooxml.name == "oMathPara" or ooxml.wrap("<m:oMathPara></m:oMathPara>")
    ooxml.elements.first.previous =
      "<m:oMathParaPr><m:jc m:val='#{dir}'/></m:oMathParaPr>"
    ooxml
  end

  def ooxml_alignment(alignnode)
    dir = "left"
    /text-align:\s*right/.match?(alignnode["style"]) and dir = "right"
    /text-align:\s*center/.match?(alignnode["style"]) and dir = "center"
    dir
  end

  def uncenter_unneeded(math, ooxml, alignnode)
    (math_block?(ooxml, math) || !alignnode) and return ooxml
    math_only_para?(alignnode) and return nil
    ooxml.name == "oMathPara" and
      ooxml = ooxml.elements.select { |x| %w(oMath r).include?(x.name) }
    ooxml.size > 1 ? nil : Nokogiri::XML::NodeSet.new(math.document, ooxml)
  end
end
