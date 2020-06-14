require "uuidtools"
require "base64"
require "mime/types"
require "image_size"
require "fileutils"

module Html2Doc
  def self.mime_preamble(boundary, filename, result)
    <<~"PREAMBLE"
    MIME-Version: 1.0
    Content-Type: multipart/related; boundary="#{boundary}"

    --#{boundary}
    Content-Location: file:///C:/Doc/#{File.basename(filename)}.htm
    Content-Type: text/html; charset="utf-8"

    #{result}

    PREAMBLE
  end

  def self.mime_attachment(boundary, filename, item, dir)
    content_type = mime_type(item)
    text_mode = %w[text application].any? { |p| content_type.start_with? p }

    path = File.join(dir, item)
    content = text_mode ? File.read(path, encoding: "utf-8") : IO.binread(path)

    encoded_file = Base64.strict_encode64(content).gsub(/(.{76})/, "\\1\n")
    <<~"FILE"
    --#{boundary}
    Content-Location: file:///C:/Doc/#{File.basename(filename)}_files/#{item}
    Content-Transfer-Encoding: base64
    Content-Type: #{content_type}

    #{encoded_file}

    FILE
  end

  def self.mime_type(item)
    types = MIME::Types.type_for(item)
    type = types ? types.first.to_s : 'text/plain; charset="utf-8"'
    type = type + ' charset="utf-8"' if /^text/.match(type) && types
    type
  end

  def self.mime_boundary
    salt = UUIDTools::UUID.random_create.to_s.gsub(/-/, ".")[0..17]
    "----=_NextPart_#{salt}"
  end

  def self.mime_package(result, filename, dir)
    boundary = mime_boundary
    mhtml = mime_preamble(boundary, filename, result)
    mhtml += mime_attachment(boundary, filename, "filelist.xml", dir)
    Dir.foreach(dir) do |item|
      next if item == "." || item == ".." || /^\./.match(item) ||
        item == "filelist.xml"
      mhtml += mime_attachment(boundary, filename, item, dir)
    end
    mhtml += "--#{boundary}--"
    File.open("#{filename}.doc", "w:UTF-8") { |f| f.write mhtml }
  end

  # max width for Word document is 400, max height is 680
  def self.image_resize(i, path, maxheight, maxwidth)
    realSize = ImageSize.path(path).size
    s = [i["width"].to_i, i["height"].to_i]
    s = realSize if s[0].zero? && s[1].zero?
    return [nil, nil] if realSize[0].nil? || realSize[1].nil?
    s[1] = s[0] * realSize[1] / realSize[0] if s[1].zero? && !s[0].zero?
    s[0] = s[1] * realSize[0] / realSize[1] if s[0].zero? && !s[1].zero?
    s = [(s[0] * maxheight / s[1]).ceil, maxheight] if s[1] > maxheight
    s = [maxwidth, (s[1] * maxwidth / s[0]).ceil] if s[0] > maxwidth
    s
  end

  IMAGE_PATH = "//*[local-name() = 'img' or local-name() = 'imagedata']".freeze

  def self.mkuuid
    UUIDTools::UUID.random_create.to_s
  end

  def self.warnsvg(src)
    warn "#{src}: SVG not supported" if /\.svg$/i.match(src)
  end

  # only processes locally stored images
  def self.image_cleanup(docxml, dir, localdir)
    docxml.traverse do |i|
      next unless i.element? && %w(img v:imagedata).include?(i.name)
      #warnsvg(i["src"])
      next if /^http/.match i["src"]
      next if %r{^data:image/[^;]+;base64}.match i["src"]
      local_filename = %r{^([A-Z]:)?/}.match(i["src"]) ? i["src"] :
        File.join(localdir, i["src"])
      new_filename = "#{mkuuid}#{File.extname(i["src"])}"
      FileUtils.cp local_filename, File.join(dir, new_filename)
      i["width"], i["height"] = image_resize(i, local_filename, 680, 400)
      i["src"] = File.join(File.basename(dir), new_filename)
    end
    docxml
  end

  # do not parse the header through Nokogiri, since it will contain 
  # non-XML like <![if !supportFootnotes]>
  def self.header_image_cleanup(doc, dir, filename, localdir)
    doc.split(%r{(<img [^>]*>|<v:imagedata [^>]*>)}).each_slice(2).map do |a|
      header_image_cleanup1(a, dir, filename, localdir)
    end.join
  end

  def self.header_image_cleanup1(a, dir, filename, localdir)
    if a.size == 2 && !(/ src="https?:/.match a[1]) &&
        !(%r{ src="data:image/[^;]+;base64}.match a[1])
      m = / src=['"](?<src>[^"']+)['"]/.match a[1]
      #warnsvg(m[:src])
      m2 = /\.(?<suffix>[a-zA-Z_0-9]+)$/.match m[:src]
      new_filename = "#{mkuuid}.#{m2[:suffix]}"
      old_filename = %r{^([A-Z]:)?/}.match(m[:src]) ? m[:src] : File.join(localdir, m[:src])
      FileUtils.cp old_filename, File.join(dir, new_filename)
      a[1].sub!(%r{ src=['"](?<src>[^"']+)['"]}, " src='file:///C:/Doc/#{filename}_files/#{new_filename}'")
    end
    a.join
  end

  def self.generate_filelist(filename, dir)
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
