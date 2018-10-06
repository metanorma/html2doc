require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"
require "xml/xslt"
require "pp"

module Html2Doc
  def self.style_list(li, level, liststyle, listnumber)
    return unless liststyle
    if li["style"]
      li["style"] += ";"
    else
      li["style"] = ""
    end
    li["style"] += "mso-list:#{liststyle} level#{level} lfo#{listnumber};"
  end

  def self.list_add(xpath, liststyles, listtype, level, listnumber)
    xpath.each_with_index do |list, i|
      listnumber = i + 1 if level == 1
      (list.xpath(".//li") - list.xpath(".//ol//li | .//ul//li")).each do |li|
        style_list(li, level, liststyles[listtype], listnumber)
        list_add(li.xpath(".//ul") - li.xpath(".//ul//ul | .//ol//ul"), liststyles, :ul, level + 1, listnumber)
        list_add(li.xpath(".//ol") - li.xpath(".//ul//ol | .//ol//ol"), liststyles, :ol, level + 1, listnumber)
      end
    end
  end

  def self.list2para(u)
    return if u.xpath("./li").empty?
    u.xpath("./li").last["class"] = "MsoListParagraphCxSpLast"
    u.xpath("./li").first["class"] = "MsoListParagraphCxSpFirst"
    u.xpath("./li").each do |l|
      l.name = "p"
      l["class"] ||= "MsoListParagraphCxSpMiddle"
      l&.first_element_child&.name == "p" and l.first_element_child.replace(l.first_element_child.children)
    end
    u.replace(u.children)
  end

  def self.lists(docxml, liststyles)
    return if liststyles.nil?
    liststyles.has_key?(:ul) and
      list_add(docxml.xpath("//ul[not(ancestor::ul) and not(ancestor::ol)]"), liststyles, :ul, 1, nil)
    liststyles.has_key?(:ol) and
      list_add(docxml.xpath("//ol[not(ancestor::ul) and not(ancestor::ol)]"), liststyles, :ol, 1, nil)
    liststyles.has_key?(:ul) and docxml.xpath("//ul").each { |u| list2para(u) }
    liststyles.has_key?(:ol) and docxml.xpath("//ol").each { |u| list2para(u) }
  end
end
