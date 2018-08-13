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
    Content-Location: file:///C:/Doc/#{filename}.htm
    Content-Type: text/html; charset="utf-8"

    #{result}

    PREAMBLE
  end

  def self.mime_attachment(boundary, filename, item, dir)
    encoded_file = Base64.strict_encode64(
      File.read("#{dir}/#{item}", encoding: "utf-8"),
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

  # max height for Word document is 400, max width is 680
  def self.image_resize(i, maxheight, maxwidth)
    realSize = ImageSize.path(i["src"]).size
    s = [i["width"].to_i, i["height"].to_i]
    s = realSize if s[0].zero? && s[1].zero?
    s[1] = s[0] * realSize[1] / realSize[0] if s[1].zero? && !s[0].zero?
    s[0] = s[1] * realSize[0] / realSize[1] if s[0].zero? && !s[1].zero?
    s = [(s[0] * maxheight / s[1]).ceil, maxheight] if s[1] > maxheight
    s = [maxwidth, (s[1] * maxwidth / s[0]).ceil] if s[0] > maxwidth
    s
  end

  IMAGE_PATH = "//*[local-name() = 'img' or local-name() = 'imagedata']".freeze

  def self.image_cleanup(docxml, dir)
    docxml.xpath(IMAGE_PATH).each do |i|
      matched = /\.(?<suffix>\S+)$/.match i["src"]
      uuid = UUIDTools::UUID.random_create.to_s
      new_full_filename = File.join(dir, "#{uuid}.#{matched[:suffix]}")
      # presupposes that the image source is local
      #system "cp #{i['src']} #{new_full_filename}"
      FileUtils.cp i["src"], new_full_filename
      i["width"], i["height"] = image_resize(i, 400, 680)
      i["src"] = new_full_filename
    end
    docxml
  end

  # do not parse the header through Nokogiri, since it will contain 
  # non-XML like <![if !supportFootnotes]>
  def self.header_image_cleanup(doc, dir, filename)
    doc.split(%r{(<img [^>]*>|<v:imagedata [^>]*>)}).each_slice(2).map do |a|
      header_image_cleanup1(a, dir, filename)
    end.join
  end

  def self.header_image_cleanup1(a, dir, filename)
    if a.size == 2
      matched = / src=['"](?<src>[^"']+)['"]/.match a[1]
      matched2 = /\.(?<suffix>\S+)$/.match matched[:src]
      uuid = UUIDTools::UUID.random_create.to_s
      new_full_filename = "file:///C:/Doc/#{filename}_files/#{uuid}.#{matched2[:suffix]}"
      dest_filename = File.join(dir, "#{uuid}.#{matched2[:suffix]}")
      #system "cp #{matched[:src]} #{dest_filename}"
      FileUtils.cp matched[:src], dest_filename
      a[1].sub!(%r{ src=['"](?<src>[^"']+)['"]}, " src='#{new_full_filename}'")
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
