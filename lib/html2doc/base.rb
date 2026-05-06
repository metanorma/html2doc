require "uuidtools"
require "htmlentities"
require "nokogiri"
require "fileutils"

class Html2Doc
  def initialize(hash)
    @filename = hash[:filename]
    @dir = hash[:dir]
    @header_file = hash[:header_file]
    @asciimathdelims = hash[:asciimathdelims]
    @imagedir = hash[:imagedir]
    @debug = hash[:debug]
    @liststyles = hash[:liststyles]
    @stylesheet = read_stylesheet(hash[:stylesheet])
    @c = HTMLEntities.new
    @output_format = hash[:output_format] || :mht
    # Legacy MHT and image-processing paths need a temp dir; DOCX does not
    @dir1 = create_dir(@filename, @dir) if @output_format != :docx
  end

  def process(result)
    case @output_format
    when :docx then process_docx(result)
    when :mht_legacy then process_mht_legacy(result)
    else process_mht(result)
    end
  end

  def process_docx(result)
    docxml = to_xhtml(result)
    # Always process AsciiMath in stem spans (default: backtick delimiters)
    old_delims = @asciimathdelims
    @asciimathdelims ||= ["`", "`"]
    stem_to_mathml(docxml)
    @asciimathdelims = old_delims
    cleanup(docxml, skip_footnotes: true, skip_images: true)
    unwrap_math_paras(docxml)
    # inject_stem_run_style(docxml)  # TODO: fix namespace handling

    converter = DocxConverter.new(
      filename: @filename,
      imagedir: @imagedir,
      stylesheet: @stylesheet,
      liststyles: @liststyles,
      header_file: @header_file,
    )
    package = converter.convert(docxml)
    converter.save_to_file(package, "#{@filename}.docx")
  end

  # Unwrap <m:oMathPara> → <m:oMath> children.
  # The reference output uses only <m:oMath> (no <m:oMathPara>).
  def unwrap_math_paras(docxml)
    docxml.xpath("//*[local-name() = 'oMathPara']").each do |omp|
      omp.replace(omp.children)
    end
  end

  # Add <w:rStyle w:val="stem"/> to all <m:r> elements in OMML.
  # The reference output applies the "stem" character style to all math runs.
  def inject_stem_run_style(docxml)
    docxml.xpath("//*[local-name() = 'r']").each do |mr|
      # Only process m:r elements (in math namespace), not w:r
      next unless mr.namespace&.href == OOXML_MATH_NS

      # Find or create w:rPr
      w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      rpr = mr.xpath("./rPr").find { |n| n.namespace&.href == w_ns }
      unless rpr
        rpr = Nokogiri::XML::Node.new("rPr", docxml)
        rpr.default_namespace = w_ns
        # Add before m:t or as last child
        mt = mr.xpath("./*").find { |n| n.name == "t" && n.namespace&.href == OOXML_MATH_NS }
        if mt
          mt.add_previous_sibling(rpr)
        else
          mr.add_child(rpr)
        end
      end
      # Add w:rStyle if not already present
      has_rstyle = rpr.xpath("./*").any? { |n| n.name == "rStyle" }
      unless has_rstyle
        rstyle = Nokogiri::XML::Node.new("rStyle", docxml)
        rstyle["val"] = "stem"
        rpr.add_child(rstyle)
      end
    end
  end

  OOXML_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math".freeze

  # Convert <span class="stem">`asciimath`</span> to <math> elements
  # so that mathml_to_ooml in cleanup can convert them to OMML.
  def stem_to_mathml(docxml)
    require "plurimath"
    docxml.css("span.stem").each do |span|
      text = span.text.to_s.strip
      next if text.empty?

      # Strip delimiters (e.g. backticks)
      delimiters = @asciimathdelims
      if delimiters && text.start_with?(delimiters[0]) && text.end_with?(delimiters[1])
        text = text[delimiters[0].length..-(delimiters[1].length + 1)]
      end

      begin
        formula = Plurimath::Math.parse(text, :asciimath)
        mathml = formula.to_mathml
        math_node = Nokogiri::XML.fragment(mathml).children.first
        if math_node
          math_node["displaystyle"] = "true"
          span.replace(math_node)
        end
      rescue StandardError => e
        warn "html2doc: failed to parse AsciiMath '#{text}': #{e.message}"
      end
    end
  end

  # Legacy MHT output: original Word-HTML + MIME packaging pipeline.
  def process_mht_legacy(result)
    result = process_html(result)
    process_header(@header_file)
    generate_filelist(@filename, @dir1)
    File.open("#{@filename}.htm", "w:UTF-8") { |f| f.write(result) }
    mime_package(result, @filename, @dir1)
    rm_temp_files(@filename, @dir, @dir1) unless @debug
  end

  # MHT output: convert HTML to OOXML model via DocxConverter, then
  # serialize to MHTML via Uniword's Transformer + MimePackager.
  def process_mht(result)
    docxml = to_xhtml(result)
    # Always process AsciiMath in stem spans (default: backtick delimiters)
    old_delims = @asciimathdelims
    @asciimathdelims ||= ["`", "`"]
    stem_to_mathml(docxml)
    @asciimathdelims = old_delims
    cleanup(docxml, skip_footnotes: true, skip_images: true)
    unwrap_math_paras(docxml)

    converter = DocxConverter.new(
      filename: @filename,
      imagedir: @imagedir,
      stylesheet: @stylesheet,
      liststyles: @liststyles,
      header_file: @header_file,
    )
    package = converter.convert(docxml)

    # Use Transformer to convert full DocxPackage (with core_properties
    # and relationships) to Mhtml::Document
    transformer = Uniword::Transformation::Transformer.new
    doc_name = File.basename(@filename)
    mhtml_doc = transformer.docx_package_to_mhtml(package, doc_name)

    # Add header.html as MIME part if header_file is provided
    if @header_file && File.exist?(@header_file)
      add_header_to_mhtml(mhtml_doc)
    end

    # Write MHTML to file via MimePackager
    packager = Uniword::Infrastructure::MimePackager.from_document(mhtml_doc)
    packager.package("#{@filename}.doc")
  end

  # Add header.html as a MIME part to the MHTML document
  def add_header_to_mhtml(mhtml_doc)
    header_content = File.read(@header_file, encoding: "utf-8")
    header_content = header_image_cleanup(header_content, @dir1, @filename,
                                          File.dirname(@filename))

    header_part = Uniword::Mhtml::MimePart.new
    header_part.content_location = "file:///C:/Doc/#{File.basename(@filename)}_files/header.html"
    header_part.content_transfer_encoding = "base64"
    header_part.content_type = 'text/html charset="utf-8"'
    header_part.raw_content = [header_content].pack("m").gsub(/\n/, "\r\n")

    mhtml_doc.parts << header_part
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

  def cleanup(docxml, skip_footnotes: false, skip_images: false)
    locate_landscape(docxml)
    namespace(docxml.root)
    image_cleanup(docxml, @dir1, @imagedir) unless skip_images
    mathml_to_ooml(docxml)
    lists(docxml, @liststyles)
    footnotes(docxml) unless skip_footnotes
    bookmarks(docxml)
    msonormal(docxml)
    docxml
  end

  def locate_landscape(_docxml)
    # Skip if stylesheet is a .docx path (not CSS text)
    return if @stylesheet.is_a?(String) && @stylesheet.end_with?(".docx")
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
    # For .docx files, store the path instead of reading binary content
    return cssname if cssname.end_with?(".docx")
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
