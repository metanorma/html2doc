require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"

class Html2Doc
  def style_list(elem, level, liststyle, listnumber)
    return unless liststyle

    if elem["style"]
      elem["style"] += ";"
    else
      elem["style"] = ""
    end
    elem["style"] += "mso-list:#{liststyle} level#{level} lfo#{listnumber};"
  end

  def list_add1(elem, liststyles, listtype, level)
    if %i[ul ol].include? listtype
      list_add(elem.xpath(".//ul") - elem.xpath(".//ul//ul | .//ol//ul"),
               liststyles, :ul, level + 1)
      list_add(elem.xpath(".//ol") - elem.xpath(".//ul//ol | .//ol//ol"),
               liststyles, :ol, level + 1)
    else
      list_add(elem.xpath(".//ul") - elem.xpath(".//ul//ul | .//ol//ul"),
               liststyles, listtype, level + 1)
      list_add(elem.xpath(".//ol") - elem.xpath(".//ul//ol | .//ol//ol"),
               liststyles, listtype, level + 1)
    end
  end

  def list_add(xpath, liststyles, listtype, level)
    xpath.each do |l|
      level == 1 and l["seen"] = true and @listnumber += 1
      l["id"] ||= UUIDTools::UUID.random_create
      (l.xpath(".//li") - l.xpath(".//ol//li | .//ul//li")).each do |li|
        style_list(li, level, liststyles[listtype], @listnumber)
        list_add1(li, liststyles, listtype, level)
      end
      list_add_tail(l, liststyles, listtype, level)
    end
  end

  def list_add_tail(list, liststyles, listtype, level)
    list.xpath(".//ul[not(ancestor::li/ancestor::*/@id = '#{list['id']}')] | "\
               ".//ol[not(ancestor::li/ancestor::*/@id = '#{list['id']}')]")
      .each do |li|
      list_add1(li.parent, liststyles, listtype, level - 1)
    end
  end

  def list2para(list)
    return if list.xpath("./li").empty?

    list2para_position(list)
    list.xpath("./li").each do |l|
      l.name = "p"
      l["class"] ||= "MsoListParagraphCxSpMiddle"
      next unless l.first_element_child&.name == "p"

      l["style"] += (l.first_element_child["style"] || "")
      l.first_element_child.replace(l.first_element_child.children)
    end
    list.replace(list.children)
  end

  def list2para_position(list)
    list.xpath("./li").first["class"] ||= "MsoListParagraphCxSpFirst"
    list.xpath("./li").last["class"] ||= "MsoListParagraphCxSpLast"
    list.xpath("./li/p").each do |p|
      p["class"] ||= "MsoListParagraphCxSpMiddle"
    end
  end

  TOPLIST = "[not(ancestor::ul) and not(ancestor::ol)]".freeze

  def lists1(docxml, liststyles, style)
    case style
    when :ul then list_add(docxml.xpath("//ul[not(@class)]#{TOPLIST}"),
                           liststyles, :ul, 1)
    when :ol then list_add(docxml.xpath("//ol[not(@class)]#{TOPLIST}"),
                           liststyles, :ol, 1)
    else
      list_add(docxml.xpath("//ol[@class = '#{style}']#{TOPLIST} | "\
                            "//ul[@class = '#{style}']#{TOPLIST}"),
               liststyles, style, 1)
    end
  end

  def lists_unstyled(docxml, liststyles)
    liststyles.has_key?(:ul) and
      list_add(docxml.xpath("//ul#{TOPLIST}[not(@seen)]"),
               liststyles, :ul, 1)
    liststyles.has_key?(:ol) and
      list_add(docxml.xpath("//ol#{TOPLIST}[not(@seen)]"),
               liststyles, :ul, 1)
    docxml.xpath("//ul[@seen] | //ol[@seen]").each do |l|
      l.delete("seen")
    end
  end

  def lists(docxml, liststyles)
    return if liststyles.nil?

    @listnumber = 0
    liststyles.each_key { |k| lists1(docxml, liststyles, k) }
    lists_unstyled(docxml, liststyles)
    liststyles.has_key?(:ul) and docxml.xpath("//ul").each { |u| list2para(u) }
    liststyles.has_key?(:ol) and docxml.xpath("//ol").each { |u| list2para(u) }
  end
end
