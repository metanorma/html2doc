require "uuidtools"
require "htmlentities"
require "nokogiri"
require "fileutils"

class Html2Doc
  def initialize(hash)
    @filename = hash[:filename]
    @dir = hash[:dir]
    @dir1 = create_dir(@filename, @dir)
    @header_file = hash[:header_file]
    @asciimathdelims = hash[:asciimathdelims]
    @imagedir = hash[:imagedir]
    @debug = hash[:debug]
    @liststyles = hash[:liststyles]
    @stylesheet = hash[:stylesheet]
    @c = HTMLEntities.new
    @xsltemplate =
      Nokogiri::XSLT(File.read(File.join(File.dirname(__FILE__), "mml2omml.xsl"),
                               encoding: "utf-8"))
  end

  def process(result)
    result = process_html(result)
    process_header(@header_file)
    generate_filelist(@filename, @dir1)
    File.open("#{@filename}.htm", "w:UTF-8") { |f| f.write(result) }
    mime_package result, @filename, @dir1
    rm_temp_files(@filename, @dir, @dir1) unless @debug
  end

  def process_header(headerfile)
    return if headerfile.nil?

    doc = File.read(headerfile, encoding: "utf-8")
    doc = header_image_cleanup(doc, @dir1, @filename,
                               File.dirname(@filename))
    File.open("#{@dir1}/header.html", "w:UTF-8") { |f| f.write(doc) }
  end

  def clear_dir(dir)
    Dir.foreach(dir) do |f|
      fn = File.join(dir, f)
      File.delete(fn) if f != "." && f != ".."
    end
    dir
  end

  def create_dir(filename, dir)
    dir and return clear_dir(dir)
    dir = "#{filename}_files"
    FileUtils.mkdir_p(dir)
    clear_dir(dir)
  end

  def process_html(result)
    docxml = to_xhtml(asciimath_to_mathml(result, @asciimathdelims))
    define_head(cleanup(docxml))
    msword_fix(from_xhtml(docxml))
  end

  def rm_temp_files(filename, dir, dir1)
    FileUtils.rm "#{filename}.htm"
    FileUtils.rm_f "#{dir1}/header.html"
    FileUtils.rm_r dir1 unless dir
  end

  def cleanup(docxml)
    namespace(docxml.root)
    image_cleanup(docxml, @dir1, @imagedir)
    mathml_to_ooml(docxml)
    lists(docxml, @liststyles)
    footnotes(docxml)
    bookmarks(docxml)
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

  def to_xhtml(xml)
    xml.gsub!(/<\?xml[^>]*>/, "")
    unless /<!DOCTYPE /.match? xml
      xml = '<!DOCTYPE html SYSTEM
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">' + xml
    end
    xml = xml.gsub(/<!--\s*\[([^\]]+)\]>/, "<!-- MSWORD-COMMENT \\1 -->")
      .gsub(/<!\s*\[endif\]\s*-->/, "<!-- MSWORD-COMMENT-END -->")
    Nokogiri::XML.parse(xml)
  end

  DOCTYPE = <<~"DOCTYPE".freeze
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
    doc.gsub!(%r{(<a style="mso-comment-reference:[^>/]+)/>}, "\\1></a>")
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

  def define_head1(docxml, _dir)
    docxml.xpath("//*[local-name() = 'head']").each do |h|
      h.children.first.add_previous_sibling <<~XML
        #{PRINT_VIEW}
          <link rel="File-List" href="cid:filelist.xml"/>
      XML
    end
  end

  def filename_substitute(head, header_filename)
    return if header_filename.nil?

    head.xpath(".//*[local-name() = 'style']").each do |s|
      s1 = s.to_xml.gsub(/url\("[^"]+"\)/) do |m|
        /FILENAME/.match?(m) ? "url(cid:header.html)" : m
      end
      s.replace(s1)
    end
  end

  def stylesheet(_filename, _header_filename, cssname)
    (cssname.nil? || cssname.empty?) and
      cssname = File.join(File.dirname(__FILE__), "wordstyle.css")
    stylesheet = File.read(cssname, encoding: "UTF-8")
    xml = Nokogiri::XML("<style/>")
    # s = Nokogiri::XML::CDATA.new(xml, "\n#{stylesheet}\n")
    # xml.children.first << Nokogiri::XML::Comment.new(xml, s)
    xml.children.first << Nokogiri::XML::CDATA
      .new(xml, "\n<!--\n#{stylesheet}\n-->\n")

    xml.root.to_s
  end

  def define_head(docxml)
    title = docxml.at("//*[local-name() = 'head']/*[local-name() = 'title']")
    head = docxml.at("//*[local-name() = 'head']")
    css = stylesheet(@filename, @header_file, @stylesheet)
    add_stylesheet(head, title, css)
    filename_substitute(head, @header_file)
    define_head1(docxml, @dir1)
    rootnamespace(docxml.root)
  end

  def add_stylesheet(head, title, css)
    if head.children.empty?
      head.add_child css
    elsif title.nil?
      head.children.first.add_previous_sibling css
    else
      title.add_next_sibling css
    end
  end

  def namespace(root)
    {
      o: "urn:schemas-microsoft-com:office:office",
      w: "urn:schemas-microsoft-com:office:word",
      v: "urn:schemas-microsoft-com:vml",
      m: "http://schemas.microsoft.com/office/2004/12/omml",
    }.each { |k, v| root.add_namespace_definition(k.to_s, v) }
  end

  def rootnamespace(root)
    root.add_namespace(nil, "http://www.w3.org/TR/REC-html40")
  end

  def bookmarks(docxml)
    docxml.xpath("//*[@id][not(@name)][not(@style = 'mso-element:footnote')]")
      .each do |x|
      next if x["id"].empty? ||
        %w(shapetype v:shapetype shape v:shape).include?(x.name)

      if x.children.empty? then x.add_child("<a name='#{x['id']}'></a>")
      else x.children.first.previous = "<a name='#{x['id']}'></a>"
      end
      x.delete("id")
    end
  end

  def msonormal(docxml)
    docxml.xpath("//*[local-name() = 'p'][not(self::*[@class])]").each do |p|
      p["class"] = "MsoNormal"
    end
    docxml.xpath("//*[local-name() = 'li'][not(self::*[@class])]").each do |p|
      p["class"] = "MsoNormal"
    end
  end
end
