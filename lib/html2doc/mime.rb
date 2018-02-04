require "uuidtools"
require "nokogiri"

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
end
