require "uuidtools"
require "asciimath"
require "htmlentities"
require "nokogiri"
require "uuidtools"

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

  def self.list_add1(li, liststyles, listtype, level)
    if [:ul, :ol].include? listtype
          list_add(li.xpath(".//ul") - li.xpath(".//ul//ul | .//ol//ul"),
                   liststyles, :ul, level + 1)
          list_add(li.xpath(".//ol") - li.xpath(".//ul//ol | .//ol//ol"),
                   liststyles, :ol, level + 1)
        else
          list_add(li.xpath(".//ul") - li.xpath(".//ul//ul | .//ol//ul"),
                   liststyles, listtype, level + 1)
          list_add(li.xpath(".//ol") - li.xpath(".//ul//ol | .//ol//ol"),
                   liststyles, listtype, level + 1)
        end
  end

  def self.list_add(xpath, liststyles, listtype, level)
    xpath.each_with_index do |list, i|
      @listnumber += 1 if level == 1
      list["seen"] = true if level == 1
      list["id"] ||= UUIDTools::UUID.random_create
      (list.xpath(".//li") - list.xpath(".//ol//li | .//ul//li")).each do |li|
        style_list(li, level, liststyles[listtype], @listnumber)
        list_add1(li, liststyles, listtype, level)
      end
      list.xpath(".//ul[not(ancestor::li/ancestor::*/@id = '#{list['id']}')] | "\
                 ".//ol[not(ancestor::li/ancestor::*/@id = '#{list['id']}')]").each do |li|
        list_add1(li.parent, liststyles, listtype, level-1)
      end
    end
  end

  def self.list2para(u)
    return if u.xpath("./li").empty?
    u.xpath("./li").first["class"] ||= "MsoListParagraphCxSpFirst"
    u.xpath("./li").last["class"] ||= "MsoListParagraphCxSpLast"
    u.xpath("./li/p").each { |p| p["class"] ||= "MsoListParagraphCxSpMiddle" }
    u.xpath("./li").each do |l|
      l.name = "p"
      l["class"] ||= "MsoListParagraphCxSpMiddle"
      l&.first_element_child&.name == "p" and
        l.first_element_child.replace(l.first_element_child.children)
    end
    u.replace(u.children)
  end

  TOPLIST = "[not(ancestor::ul) and not(ancestor::ol)]".freeze

  def self.lists1(docxml, liststyles, k)
    case k
    when :ul then list_add(docxml.xpath("//ul[not(@class)]#{TOPLIST}"),
                            liststyles, :ul, 1)
    when :ol then list_add(docxml.xpath("//ol[not(@class)]#{TOPLIST}"),
                           liststyles, :ol, 1)
    else
      list_add(docxml.xpath("//ol[@class = '#{k.to_s}']#{TOPLIST} | "\
                            "//ul[@class = '#{k.to_s}']#{TOPLIST}"),
      liststyles, k, 1)
    end
  end

  def self.lists_unstyled(docxml, liststyles)
    list_add(docxml.xpath("//ul#{TOPLIST}[not(@seen)]"),
             liststyles, :ul, 1) if liststyles.has_key?(:ul)
    list_add(docxml.xpath("//ol#{TOPLIST}[not(@seen)]"),
             liststyles, :ul, 1) if liststyles.has_key?(:ol)
    docxml.xpath("//ul[@seen] | //ol[@seen]").each do |l|
      l.delete("seen")
    end
  end

  def self.lists(docxml, liststyles)
    return if liststyles.nil?
    @listnumber = 0
    liststyles.each_key { |k| lists1(docxml, liststyles, k) }
    lists_unstyled(docxml, liststyles)
    liststyles.has_key?(:ul) and docxml.xpath("//ul").each { |u| list2para(u) }
    liststyles.has_key?(:ol) and docxml.xpath("//ol").each { |u| list2para(u) }
  end
end
