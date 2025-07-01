require "uuidtools"
require "base64"
require "mime/types"
require "fileutils"
require "vectory"

class Html2Doc
  def mime_preamble(boundary, filename, result)
    <<~"PREAMBLE"
      MIME-Version: 1.0
      Content-Type: multipart/related; boundary="#{boundary}"

      --#{boundary}
      Content-ID: <#{File.basename(filename)}>
      Content-Disposition: inline; filename="#{File.basename(filename)}"
      Content-Type: text/html; charset="utf-8"

      #{result}

    PREAMBLE
  end

  def mime_attachment(boundary, _filename, item, dir)
    content_type = mime_type(item)
    text_mode = %w[text application].any? { |p| content_type.start_with? p }

    path = File.join(dir, item)
    content = text_mode ? File.read(path, encoding: "utf-8") : IO.binread(path)

    encoded_file = Base64.strict_encode64(content).gsub(/(.{76})/, "\\1\n")
    <<~"FILE"
      --#{boundary}
      Content-ID: <#{File.basename(item)}>
      Content-Disposition: inline; filename="#{File.basename(item)}"
      Content-Transfer-Encoding: base64
      Content-Type: #{content_type}

      #{encoded_file}

    FILE
  end

  def mime_type(item)
    types = MIME::Types.type_for(item)
    type = types ? types.first.to_s : 'text/plain; charset="utf-8"'
    type = %(#{type} charset="utf-8") if /^text/.match(type) && types
    type
  end

  def mime_boundary
    salt = UUIDTools::UUID.random_create.to_s.tr("-", ".")[0..17]
    "----=_NextPart_#{salt}"
  end

  def mime_package(result, filename, dir)
    boundary = mime_boundary
    mhtml = mime_preamble(boundary, "#{filename}.htm", result)
    mhtml += mime_attachment(boundary, "#{filename}.htm", "filelist.xml", dir)
    Dir.foreach(dir) do |item|
      next if item == "." || item == ".." || /^\./.match(item) ||
        item == "filelist.xml"

      mhtml += mime_attachment(boundary, "#{filename}.htm", item, dir)
    end
    mhtml += "--#{boundary}--"
    File.open("#{filename}.doc", "w:UTF-8") { |f| f.write contentid(mhtml) }
  end

  def contentid(mhtml)
    mhtml.gsub %r{(<img[^<>]*?src=")([^"'<]+)(['"])}m do |m|
      repl = "#{$1}cid:#{File.basename($2)}#{$3}"
      /^data:|^https?:/ =~ $2 ? m : repl
    end.gsub %r{(<v:imagedata[^<>]*?src=")([^"'<]+)(['"])}m do |m|
      repl = "#{$1}cid:#{File.basename($2)}#{$3}"
      /^data:|^https?:/ =~ $2 ? m : repl
    end
  end

  IMAGE_PATH = "//*[local-name() = 'img' or local-name() = 'imagedata']".freeze

  def mkuuid
    UUIDTools::UUID.random_create.to_s
  end

  def warnsvg(src)
    warn "#{src}: SVG not supported" if /\.svg$/i.match?(src)
  end

  def localname(src, localdir)
    %r{^([A-Z]:)?/}.match?(src) ? src : File.join(localdir, src)
  end

  # only processes locally stored images
  def image_cleanup(docxml, dir, localdir)
    maxheight, maxwidth = page_dimensions(docxml)
    docxml.traverse do |i|
      skip_image_cleanup?(i) and next
      local_filename = rename_image(i, dir, localdir)
      i["width"], i["height"] =
        if landscape?(i)
          Vectory.image_resize(i, local_filename, maxwidth, maxheight)
        else
          Vectory.image_resize(i, local_filename, maxheight, maxwidth)
        end
    end
    docxml
  end

  def landscape?(img)
    img.ancestors.each do |a|
      a.name == "div" or next
      @landscape.include?(a["class"]) and return true
    end
    false
  end

  def rename_image(img, dir, localdir)
    local_filename = localname(img["src"], localdir)
    new_filename = "#{mkuuid}#{File.extname(img['src'])}"
    FileUtils.cp local_filename, File.join(dir, new_filename)
    img["src"] = File.join(File.basename(dir), new_filename)
    local_filename
  end

  def skip_image_cleanup?(img)
    src = img["src"]
    (img.element? && %w(img imagedata).include?(img.name)) or return true
    (src.nil? || src.empty? || /^http/.match?(src) ||
      %r{^data:(image|application)/[^;]+;base64}.match?(src)) and return true
    false
  end

  # we are going to use the 2nd instance of @page in the Word CSS,
  # skipping the cover page. Currently doesn't deal with Landscape.
  # Scan both @stylesheet and docxml.to_xml (where @standardstylesheet has ended up)
  # Allow 0.9 * height to fit caption
  def page_dimensions(docxml)
    page_size = find_page_size_in_doc(@stylesheet, docxml.to_xml) or
      return [680, 400]
    m_size = /size:\s*(\S+)\s+(\S+)\s*;/.match(page_size) or return [680, 400]
    m_marg = /margin:\s*(\S+)\s+(\S+)\s*(\S+)\s*(\S+)\s*;/.match(page_size) or
      return [680, 400]
    [0.9 * (units_to_px(m_size[2]) - units_to_px(m_marg[1]) - units_to_px(m_marg[3])),
     units_to_px(m_size[1]) - units_to_px(m_marg[2]) - units_to_px(m_marg[4])]
  rescue StandardError
    [680, 400]
  end

  def find_page_size_in_doc(stylesheet, doc)
    find_page_size(stylesheet, "WordSection2", false) ||
      find_page_size(stylesheet, "WordSection3", false) ||
      find_page_size(doc, "WordSection2", true) ||
      find_page_size(doc, "WordSection3", true) ||
      find_page_size(stylesheet, "", false) || find_page_size(doc, "", true)
  end

  # if in_xml, CSS is embedded in XML <style> tag
  def find_page_size(stylesheet, klass, in_xml)
    xml_found = false
    found = false
    ret = ""
    stylesheet&.lines&.each do |l|
      in_xml && l.include?("<style") and xml_found = true and found = false
      in_xml && l.include?("</style>") and xml_found = false
      /^\s*@page\s+#{klass}/.match?(l) and found = true
      found && /^\s*\{?size:/.match?(l) and ret += l
      found && /^\s*\{?margin:/.match?(l) and ret += l
      if found && /}/.match?(l)
        !ret.blank? && (!in_xml || xml_found) and return ret
        ret = ""
        found = false
      end
    end
    nil
  end

  def units_to_px(measure)
    m = /^(\S+)(pt|cm)/.match(measure)
    ret = case m[2]
          when "px" then (m[1].to_f * 0.75)
          when "pt" then m[1].to_f
          when "cm" then (m[1].to_f * 28.346456693)
          when "in" then (m[1].to_f * 72)
          end
    ret.to_i
  end

  # do not parse the header through Nokogiri, since it will contain
  # non-XML like <![if !supportFootnotes]>
  def header_image_cleanup(doc, dir, filename, localdir)
    doc.split(%r{(<img [^>]*>|<v:imagedata [^>]*>)}).each_slice(2).map do |a|
      header_image_cleanup1(a, dir, filename, localdir)
    end.join
  end

  def header_image_cleanup1(a, dir, _filename, localdir)
    if a.size == 2 && !(/ src="https?:/.match a[1]) &&
        !(%r{ src="data:(image|application)/[^;]+;base64}.match a[1])
      m = / src=['"](?<src>[^"']+)['"]/.match a[1]
      m2 = /\.(?<suffix>[a-zA-Z_0-9]+)$/.match m[:src]
      new_filename = "#{mkuuid}.#{m2[:suffix]}"
      FileUtils.cp localname(m[:src], localdir), File.join(dir, new_filename)
      a[1].sub!(%r{ src=['"](?<src>[^"']+)['"]}, " src='cid:#{new_filename}'")
    end
    a.join
  end

  def generate_filelist(filename, dir)
    File.open(File.join(dir, "filelist.xml"), "w") do |f|
      f.write %{<xml xmlns:o="urn:schemas-microsoft-com:office:office">
        <o:MainFile HRef="../#{filename}.htm"/>}
      Dir.entries(dir).sort.each do |item|
        (item == "." || item == ".." || /^\./.match(item)) and next
        f.write %{  <o:File HRef="#{item}"/>\n}
      end
      f.write("</xml>\n")
    end
  end
end
