require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"
require "xml/xslt"
require "pp"

module Html2Doc
  @xslt = XML::XSLT.new
  @xslt.xsl = File.read(File.join(File.dirname(__FILE__), "mathml2omml.xsl"))

  def self.process(result, hash)
    hash[:dir1] = create_dir(hash[:filename], hash[:dir])
    result = process_html(result, hash)
    hash[:header_file].nil? ||
      system("cp #{hash[:header_file]} #{hash[:dir1]}/header.html")
    generate_filelist(hash[:filename], hash[:dir1])
    File.open("#{hash[:filename]}.htm", "w") { |f| f.write(result) }
    mime_package result, hash[:filename], hash[:dir1]
    rm_temp_files(hash[:filename], hash[:dir], hash[:dir1])
  end

  def self.create_dir(filename, dir)
    return dir if dir
    dir = "#{filename}_files"
    Dir.mkdir(dir) unless File.exists?(dir)
    dir
  end

  def self.process_html(result, hash)
    docxml = to_xhtml(asciimath_to_mathml(result, hash[:asciimathdelims]))
    define_head(cleanup(docxml, hash[:dir]), hash)
    msword_fix(from_xhtml(docxml))
  end

  def self.rm_temp_files(filename, dir, dir1)
    system "rm #{filename}.htm"
    system "rm -r #{dir1}" unless dir
  end

  def self.cleanup(docxml, dir)
    image_cleanup(docxml, dir)
    mathml_to_ooml(docxml)
    footnotes(docxml)
    msonormal(docxml)
    docxml
  end

  NOKOHEAD = <<~HERE.freeze
    <!DOCTYPE html SYSTEM
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head> <title></title> <meta charset="UTF-8" /> </head>
    <body> </body> </html>
  HERE

  def self.to_xhtml(xml)
    xml.gsub!(/<\?xml[^>]*>/, "")
    unless /<!DOCTYPE /.match? xml
      xml = '<!DOCTYPE html SYSTEM
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">' + xml
    end
    Nokogiri::XML.parse(xml)
  end

  DOCTYPE = <<~"DOCTYPE".freeze
    <!DOCTYPE html SYSTEM "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
  DOCTYPE

  def self.from_xhtml(xml)
    xml.to_xml.sub(%{ xmlns="http://www.w3.org/1999/xhtml"}, "").
      sub(DOCTYPE, "").
      gsub(%{ />}, "/>")
  end

  def self.msword_fix(r)
    # brain damage in MSWord parser
    r.gsub!(%r{<span style="mso-special-character:footnote"/>},
            '<span style="mso-special-character:footnote"></span>')
    r.gsub!(%r{<div style="mso-element:footnote-list"></div>},
            '<div style="mso-element:footnote-list"/>')
    r.gsub!(%r{(<a style="mso-comment-reference:[^>/]+)/>}, "\\1></a>")
    r.gsub!(%r{<link rel="File-List"}, "<link rel=File-List")
    r.gsub!(%r{<meta http-equiv="Content-Type"},
            "<meta http-equiv=Content-Type")
    r.gsub!(%r{&tab;|&amp;tab;}, '<span style="mso-tab-count:1">&#xA0; </span>')
    r
  end

  PRINT_VIEW = <<~XML.freeze
    <!--[if gte mso 9]>
    <xml>
    <w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
    </xml>
    <![endif]-->
    <meta http-equiv=Content-Type content="text/html; charset=utf-8"/>
  XML

  def self.define_head1(docxml, dir)
    docxml.xpath("//*[local-name() = 'head']").each do |h|
      h.children.first.add_previous_sibling <<~XML
      #{PRINT_VIEW}
        <link rel="File-List" href="#{dir}/filelist.xml"/>
      XML
    end
  end

  def self.filename_substitute(stylesheet, header_filename, filename)
    if header_filename.nil?
      stylesheet.gsub!(/\n[^\n]*FILENAME[^\n]*i\n/, "\n")
    else
      stylesheet.gsub!(/FILENAME/, filename)
    end
    stylesheet
  end

  def self.stylesheet(filename, header_filename, fn)
    (fn.nil? || fn.empty?) &&
      fn = File.join(File.dirname(__FILE__), "wordstyle.css")
    stylesheet = File.read(fn, encoding: "UTF-8")
    stylesheet = filename_substitute(stylesheet, header_filename, filename)
    xml = Nokogiri::XML("<style/>")
    xml.children.first << Nokogiri::XML::Comment.new(xml, "\n#{stylesheet}\n")
    xml.root.to_s
  end

  def self.define_head(docxml, hash)
    title = docxml.at("//*[local-name() = 'head']/*[local-name() = 'title']")
    head = docxml.at("//*[local-name() = 'head']")
    css = stylesheet(hash[:filename], hash[:header_file], hash[:stylesheet])
    add_stylesheet(head, title, css)
    define_head1(docxml, hash[:dir1])
    namespace(docxml.root)
  end

  def self.add_stylesheet(head, title, css)
    if head.children.empty?
      head.add_child css
    elsif title.nil?
      head.children.first.add_previous_sibling css
    else
      title.add_next_sibling css
    end
  end

  def self.namespace(root)
    {
      o: "urn:schemas-microsoft-com:office:office",
      w: "urn:schemas-microsoft-com:office:word",
      m: "http://schemas.microsoft.com/office/2004/12/omml",
    }.each { |k, v| root.add_namespace_definition(k.to_s, v) }
    root.add_namespace(nil, "http://www.w3.org/TR/REC-html40")
  end

  def self.msonormal(docxml)
    docxml.xpath("//*[local-name() = 'p'][not(self::*[@class])]").each do |p|
      p["class"] = "MsoNormal"
    end
    docxml.xpath("//*[local-name() = 'li'][not(self::*[@class])]").each do |p|
      p["class"] = "MsoNormal"
    end
  end
end
