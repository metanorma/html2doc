require "uuidtools"
require "htmlentities"
require "nokogiri"

class Html2Doc
  def style_list(elem, level, liststyle, listnumber)
    liststyle or return
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
      level == 1 && l["seen"] = true and @listnumber += 1
      l["id"] ||= UUIDTools::UUID.random_create
      liststyle = derive_liststyle(l, liststyles[listtype], level)
      (l.xpath(".//li") - l.xpath(".//ol//li | .//ul//li")).each do |li|
        style_list(li, level, liststyle, @listnumber)
        list_add1(li, liststyles, listtype, level)
      end
      list_add_tail(l, liststyles, listtype, level)
    end
  end

  def derive_liststyle(list, liststyle, level)
    list["start"] && list["start"] != "1" or return liststyle
    @liststyledefsidx += 1
    ret = "l#{@liststyledefsidx}"
    @newliststyledefs += newliststyle(list["start"], liststyle, ret, level)
    ret
  end

  def newliststyle(start, liststyle, newstylename, level)
    s = @liststyledefs[liststyle]
      .gsub(/@list\s+#{liststyle}/, "@list #{newstylename}")
      .sub(/@list\s+#{newstylename}\s+\{[^}]*\}/m, <<~LISTSTYLE)
        @list #{newstylename}\n{mso-list-id:#{rand(100_000_000..999_999_999)};
        mso-list-template-ids:#{rand(100_000_000..999_999_999)};}
      LISTSTYLE
      .sub(/@list\s+#{newstylename}:level#{level}\s+\{/m,
           "\\0mso-level-start-at:#{start};\n")
    "#{s}\n"
  end

  def list_add_tail(list, liststyles, listtype, level)
    list.xpath(".//ul[not(ancestor::li/ancestor::*/@id = '#{list['id']}')] | "\
               ".//ol[not(ancestor::li/ancestor::*/@id = '#{list['id']}')]")
      .each do |li|
      list_add1(li.parent, liststyles, listtype, level - 1)
    end
  end

  def list2para(list)
    list.xpath("./li").empty? and return
    list2para_position(list)
    list.xpath("./li").each do |l|
      l.name = "p"
      l["class"] ||= "MsoListParagraphCxSpMiddle"
      l.first_element_child&.name == "p" or next
      l["style"] ||= ""
      l["style"] += l.first_element_child["style"]
        &.sub(/mso-list[^;]+;/, "") || ""
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
    liststyles.nil? and return
    parse_stylesheet_line_styles
    liststyles.each_key { |k| lists1(docxml, liststyles, k) }
    lists_unstyled(docxml, liststyles)
    liststyles.has_key?(:ul) and docxml.xpath("//ul").each { |u| list2para(u) }
    liststyles.has_key?(:ol) and docxml.xpath("//ol").each { |u| list2para(u) }
  end

  def parse_stylesheet_line_styles
    @listnumber = 0
    result = process_stylesheet_lines(@stylesheet.split("\n"))
    @liststyledefs = clean_result_content(result)
    @newliststyledefs = ""
    @liststyledefsidx = @liststyledefs.keys.map do |k|
      k.sub(/^.*(\d+)$/, "\\1").to_i
    end.max
  end

  private

  def extract_list_name(line)
    match = line.match(/^\s*@list\s+([^:\s]+)(?::.*)?/)
    match ? match[1] : nil
  end

  def list_declaration?(line)
    !extract_list_name(line).nil?
  end

  def save_current_list(result, current_base, current_content)
    current_base.nil? || current_content.empty? and return result
    if result[current_base]
      result[current_base] += current_content
    else
      result[current_base] = current_content
    end
    result
  end

  def process_stylesheet_lines(lines)
    result = {}
    current_base = nil
    current_content = ""
    parsing_active = false

    lines.each do |line|
      if list_declaration?(line)
        base_name = extract_list_name(line)
        if current_base == base_name
          current_content += "#{line}\n"
        else
          # save accumulated list style definition, new list style
          save_current_list(result, current_base, current_content)
          current_base = base_name
          current_content = "#{line}\n"
        end
        parsing_active = true

      elsif parsing_active && line.include?("}")
        # End of current block - add this line and stop parsing
        current_content += "#{line}\n"
        parsing_active = false

      elsif parsing_active
        # Continue adding content while parsing is active
        current_content += "#{line}\n"
      end
      # If parsing_active is false and no @list declaration, skip the line
    end
    # Save the last list if we were still parsing
    save_current_list(result, current_base, current_content)
    result
  end

  def clean_result_content(result)
    result.each { |k, v| result[k] = v.rstrip }
    result
  end
end
