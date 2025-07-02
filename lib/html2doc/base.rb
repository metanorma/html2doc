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
    @stylesheet = read_stylesheet(hash[:stylesheet])
    @c = HTMLEntities.new
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
    headerfile.nil? and return
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
    docxml = to_xhtml(result)
    define_head(cleanup(docxml))
    msword_fix(from_xhtml(docxml))
  end

  def rm_temp_files(filename, dir, dir1)
    FileUtils.rm "#{filename}.htm"
    FileUtils.rm_f "#{dir1}/header.html"
    FileUtils.rm_r dir1 unless dir
  end

  def cleanup(docxml)
    locate_landscape(docxml)
    namespace(docxml.root)
    image_cleanup(docxml, @dir1, @imagedir)
    mathml_to_ooml(docxml)
    lists(docxml, @liststyles)
    footnotes(docxml)
    bookmarks(docxml)
    msonormal(docxml)
    docxml
  end

  def locate_landscape(_docxml)
    @landscape = @stylesheet.scan(/div\.\S+\s+\{\s*page:\s*[^;]+L;\s*\}/m)
      .map { |e| e.sub(/^div\.(\S+).*$/m, "\\1") }
  end

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

  def stylesheet(_filename, _header_filename, _cssname)
    stylesheet = "#{@stylesheet}\n#{@newliststyledefs}"
    xml = Nokogiri::XML("<style/>")
    xml.children.first << Nokogiri::XML::CDATA
      .new(xml, "\n<!--\n#{stylesheet}\n-->\n")
    xml.root.to_s
  end

  def read_stylesheet(cssname)
    (cssname.nil? || cssname.empty?) and
      cssname = File.join(File.dirname(__FILE__), "wordstyle.css")
    File.read(cssname, encoding: "UTF-8")
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
    else title.add_next_sibling css
    end
  end

  def bookmarks(docxml)
    docxml.xpath("//*[@id][not(@name)][not(@style = 'mso-element:footnote')]")
      .each do |x|
      (x["id"].empty? || x.namespace&.prefix == "v" &&
        %w(shapetype shape rect line group).include?(x.name)) and next
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
