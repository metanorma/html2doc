require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"
require "xml/xslt"
require "pp"

module Html2Doc
  @xslt = XML::XSLT.new
  #@xslt.xsl = File.read(File.join(File.dirname(__FILE__), "mathml2omml.xsl"))
  @xslt.xsl = File.read(File.join(File.dirname(__FILE__), "mml2omml.xsl"))

  def self.asciimath_to_mathml1(x)
    AsciiMath.parse(HTMLEntities.new.decode(x)).to_mathml.
        gsub(/<math>/, "<math xmlns='http://www.w3.org/1998/Math/MathML'>")
  end

  def self.asciimath_to_mathml(doc, delims)
    return doc if delims.nil? || delims.size < 2
    doc.split(/(#{Regexp.escape(delims[0])}|#{Regexp.escape(delims[1])})/).
      each_slice(4).map do |a|
      a[2].nil? || a[2] = asciimath_to_mathml1(a[2])
      a.size > 1 ? a[0] + a[2] : a[0]
    end.join
  end

  # random fixes to MathML input that OOXML needs to render properly
  def self.ooxml_cleanup(m)
    m.xpath(".//xmlns:msup[name(preceding-sibling::*[1])='munderover']",
            m.document.collect_namespaces).each do |x|
      x1 = x.replace("<mrow></mrow>").first
      x1.children = x
    end
    m.add_namespace(nil, "http://www.w3.org/1998/Math/MathML")
    m.to_s
  end

  def self.mathml_to_ooml(docxml)
    docxml.xpath("//*[local-name() = 'math']").each do |m|
      @xslt.xml = ooxml_cleanup(m)
      ooxml = @xslt.serve.gsub(/<\?[^>]+>\s*/, "").
        gsub(/ xmlns(:[^=]+)?="[^"]+"/, "").
        gsub(%r{<(/)?([a-z])}, "<\\1m:\\2")
      m.swap(ooxml)
    end
  end
end
