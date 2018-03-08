require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"
require "xml/xslt"
require "pp"

module Html2Doc
  def self.style_list(li, level, listno)
    return unless listno
    if li["style"]
      li["style"] += ";"
    else
      li["style"] = ""
    end
    # I don't know what the lfo-n attribute is. I doubt Micro$oft now does either.
    li["style"] += "mso-list:#{listno} level#{level} lfo1;"
  end

  def self.list_add(xpath, liststyles, listtype, level)
    xpath.each do |list|
      (list.xpath(".//li") - list.xpath(".//ol//li | .//ul//li")).each do |li|
        style_list(li, level, liststyles[listtype])
        list_add(li.xpath(".//ul") - li.xpath(".//ul//ul | .//ol//ul"), liststyles, :ul, level + 1)
        list_add(li.xpath(".//ol") - li.xpath(".//ul//ol | .//ol//ol"), liststyles, :ol, level + 1)
      end
    end
  end

  def self.lists(docxml, liststyles)
    return if liststyles.nil?
    if liststyles.has_key?(:ul)
      list_add(docxml.xpath("//ul[not(ancestor::ul) and not(ancestor::ol)]"), liststyles, :ul, 1)
    end
    if liststyles.has_key?(:ol)
      list_add(docxml.xpath("//ol[not(ancestor::ul) and not(ancestor::ol)]"), liststyles, :ol, 1)
    end
  end
end
