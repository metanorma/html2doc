require "uuidtools"
require "nokogiri"

module Html2Doc
  def self.process(result, filename, header_file, dir)
    # preserve HTML escapes
    unless /<!DOCTYPE html/.match? result
      result.gsub!(/<\?xml version="1.0"\?>/, "")
      result = "<!DOCTYPE html SYSTEM " +
        "'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>" + result
    end
    docxml = Nokogiri::XML(result)
    image_cleanup(docxml, dir)
    define_head(docxml, dir, filename, header_file)
    result = self.msword_fix(docxml.to_xml)
    system "cp #{header_file} #{dir}/header.html" unless header_file.nil?
    generate_filelist(filename, dir)
    File.open("#{filename}.htm", "w") { |f| f.write(result) }
    mime_package result, filename, dir
  end

  def self.msword_fix(r)
    # brain damage in MSWord parser
    r.gsub(%r{<span style="mso-special-character:footnote"/>},
           '<span style="mso-special-character:footnote"></span>').
           gsub(%r{(<a style="mso-comment-reference:[^>/]+)/>}, "\\1></a>").
           gsub(%r{<link rel="File-List"}, "<link rel=File-List").
           gsub(%r{<meta http-equiv="Content-Type"},
                "<meta http-equiv=Content-Type").
                gsub(%r{&tab;|&amp;tab;},
                     '<span style="mso-tab-count:1">&#xA0; </span>')
  end

  def self.image_resize(orig_filename)
    image_size = ImageSize.path(orig_filename).size
    # max width for Word document is 400, max height is 680
    if image_size[0] > 400
      image_size[1] = (image_size[1] * 400 / image_size[0]).ceil
      image_size[0] = 400
    end
    if image_size[1] > 680
      image_size[0] = (image_size[0] * 680 / image_size[1]).ceil
      image_size[1] = 680
    end
    image_size
  end

  def self.image_cleanup(docxml, dir)
    docxml.xpath("//*[local-name() = 'img']").each do |i|
      matched = /\.(?<suffix>\S+)$/.match i["src"]
      uuid = UUIDTools::UUID.random_create.to_s
      new_full_filename = File.join(dir, "#{uuid}.#{matched[:suffix]}")
      # presupposes that the image source is local
      system "cp #{i['src']} #{new_full_filename}"
      # image_size = image_resize(i["src"])
      i["width"], i["height"] = image_resize(i["src"])
      i["src"] = new_full_filename
      #i["height"] = image_size[1]
      #i["width"] = image_size[0]
    end
    docxml
  end

  def self.define_head1(docxml, dir)
    docxml.xpath("//*[local-name() = 'head']").each do |h|
      h.children.first.add_previous_sibling <<~XML
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
            <link rel="File-List" href="#{dir}/filelist.xml"/>
      XML
    end
  end

  def self.stylesheet(filename, header_filename)
    fn = File.join(File.dirname(__FILE__), "wordstyle.css")
    stylesheet = File.read(fn, encoding: "UTF-8")
    if header_filename.nil?
      stylesheet.gsub!(/\n[^\n]*FILENAME[^\n]*i\n/, "\n")
    else
      stylesheet.gsub!(/FILENAME/, filename)
    end
    xml = Nokogiri::XML("<style/>")
    xml.children.first << Nokogiri::XML::Comment.new(xml, "\n#{stylesheet}\n")
    xml.root.to_s
  end

  def self.define_head(docxml, dir, filename, header_file)
    title = docxml.at("//*[local-name() = 'head']/*[local-name() = 'title']")
    head = docxml.at("//*[local-name() = 'head']")
    if title.nil?
      head.children.first.add_previous_sibling stylesheet(filename, header_file)
    else
      title.add_next_sibling stylesheet(filename, header_file)
    end
    self.define_head1(docxml, dir)
  end

  def self.mime_preamble(boundary, filename, result)
    <<~"PREAMBLE"
    MIME-Version: 1.0
    Content-Type: multipart/related; boundary="#{boundary}"

    --#{boundary}
    Content-Location: file:///C:/Doc/#{filename}.htm
    Content-Type: text/html; charset="utf-8"

    #{result}

    PREAMBLE
  end

  def self.mime_attachment(boundary, filename, item, dir)
    encoded_file = Base64.strict_encode64(
      File.read("#{dir}/#{item}"),
    ).gsub(/(.{76})/, "\\1\n")
    <<~"FILE"
    --#{boundary}
    Content-Location: file:///C:/Doc/#{filename}_files/#{item}
    Content-Transfer-Encoding: base64
    Content-Type: #{mime_type(item)}

    #{encoded_file}

    FILE
  end

  def self.mime_type(item)
    types = MIME::Types.type_for(item)
    type = types ? types.first.to_s : 'text/plain; charset="utf-8"'
    type = type + ' charset="utf-8"' if /^text/.match?(type) && types
    type
  end

  def self.mime_boundary
    salt = UUIDTools::UUID.random_create.to_s.gsub(/-/, ".")[0..17]
    "----=_NextPart_#{salt}"
  end

  def self.mime_package(result, filename, dir)
    boundary = mime_boundary
    mhtml = mime_preamble(boundary, filename, result)
    Dir.foreach(dir) do |item|
      next if item == "." || item == ".." || /^\./.match(item)
      mhtml += mime_attachment(boundary, filename, item, dir)
    end
    mhtml += "--#{boundary}--"
    File.open("#{filename}.doc", "w") { |f| f.write mhtml }
  end

  def self.generate_filelist(filename, dir)
    File.open(File.join(dir, "filelist.xml"), "w") do |f|
      f.write(<<~"XML")
                <xml xmlns:o="urn:schemas-microsoft-com:office:office">
                 <o:MainFile HRef="../#{filename}.htm"/>
      XML
      Dir.foreach(dir) do |item|
        next if item == "." || item == ".." || /^\./.match(item)
        f.write %{  <o:File HRef="#{item}"/>\n}
      end
      f.write("</xml>\n")
    end
  end
end
