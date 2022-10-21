require "uuidtools"
require "base64"
require "mime/types"
require "image_size"
require "fileutils"

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
    salt = UUIDTools::UUID.random_create.to_s.gsub(/-/, ".")[0..17]
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
    mhtml.gsub %r{(<img[^>]*?src=")([^"']+)(['"])}m do |m|
      repl = "#{$1}cid:#{File.basename($2)}#{$3}"
      /^data:|^https?:/ =~ $2 ? m : repl
    end.gsub %r{(<v:imagedata[^>]*?src=")([^"']+)(['"])}m do |m|
      repl = "#{$1}cid:#{File.basename($2)}#{$3}"
      /^data:|^https?:/ =~ $2 ? m : repl
    end
  end

  # max width for Word document is 400, max height is 680
  def image_resize(img, path, maxheight, maxwidth)
    s, realsize = get_image_size(img, path)
    return s if s[0] == nil && s[1] == nil

    if img.name == "svg" && !img["viewBox"]
      img["viewBox"] = "0 0 #{s[0]} #{s[1]}"
    end
    s[1] = s[0] * realsize[1] / realsize[0] if s[1].zero? && !s[0].zero?
    s[0] = s[1] * realsize[0] / realsize[1] if s[0].zero? && !s[1].zero?
    s = [(s[0] * maxheight / s[1]).ceil, maxheight] if s[1] > maxheight
    s = [maxwidth, (s[1] * maxwidth / s[0]).ceil] if s[0] > maxwidth
    s
  end

  def get_image_size(img, path)
    realsize = ImageSize.path(path).size # does not support emf
    s = [img["width"].to_i, img["height"].to_i]
    return [s, s] if /\.emf$/.match?(path)

    s = realsize if s[0].zero? && s[1].zero?
    s = [nil, nil]  if realsize.nil? || realsize[0].nil? || realsize[1].nil?
    [s, realsize]
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
    docxml.traverse do |i|
      src = i["src"]
      next unless i.element? && %w(img v:imagedata).include?(i.name)
      next if src.nil? || src.empty? || /^http/.match?(src)
      next if %r{^data:(image|application)/[^;]+;base64}.match? src

      local_filename = localname(src, localdir)
      new_filename = "#{mkuuid}#{File.extname(src)}"
      FileUtils.cp local_filename, File.join(dir, new_filename)
      i["width"], i["height"] = image_resize(i, local_filename, 680, 400)
      i["src"] = File.join(File.basename(dir), new_filename)
    end
    docxml
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
        next if item == "." || item == ".." || /^\./.match(item)

        f.write %{  <o:File HRef="#{item}"/>\n}
      end
      f.write("</xml>\n")
    end
  end
end
