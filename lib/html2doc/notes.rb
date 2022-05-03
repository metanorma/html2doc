require "uuidtools"

class Html2Doc
  def footnotes(docxml)
    i = 1
    fn = []
    docxml.xpath("//a").each do |a|
      next unless process_footnote_link(docxml, a, i, fn)

      i += 1
    end
    process_footnote_texts(docxml, fn)
  end

  def process_footnote_texts(docxml, footnotes)
    body = docxml.at("//body")
    list = body.add_child("<div style='mso-element:footnote-list'/>")
    footnotes.each_with_index do |f, i|
      fn = list.first.add_child(footnote_container(docxml, i + 1))
      f.parent = fn.first
      footnote_div_to_p(f)
    end
    footnote_cleanup(docxml)
  end

  def footnote_div_to_p(elem)
    if %w{div aside}.include? elem.name
      if elem.at(".//p")
        elem.replace(elem.children)
      else
        elem.name = "p"
        elem["class"] = "MsoFootnoteText"
      end
    end
  end

  FN = "<span class='MsoFootnoteReference'>"\
    "<span style='mso-special-character:footnote'/></span>".freeze

  def footnote_container(docxml, idx)
    ref = docxml&.at("//a[@href='#_ftn#{idx}']")&.children&.to_xml(indent: 0)
      &.gsub(/>\n</, "><") || FN
    <<~DIV
      <div style='mso-element:footnote' id='ftn#{idx}'>
        <a style='mso-footnote-id:ftn#{idx}' href='#_ftn#{idx}'
           name='_ftnref#{idx}' title='' id='_ftnref#{idx}'>#{ref.strip}</a></div>
    DIV
  end

  def process_footnote_link(docxml, elem, idx, footnote)
    return false unless footnote?(elem)

    href = elem["href"].gsub(/^#/, "")
    note = docxml.at("//*[@name = '#{href}' or @id = '#{href}']")
    return false if note.nil?

    set_footnote_link_attrs(elem, idx)
    if elem.at("./span[@class = 'MsoFootnoteReference']")
      process_footnote_link1(elem)
    else elem.children = FN
    end
    footnote << transform_footnote_text(note)
  end

  def process_footnote_link1(elem)
    elem.children.each do |c|
      if c.name == "span" && c["class"] == "MsoFootnoteReference"
        c.replace(FN)
      else
        c.wrap("<span class='MsoFootnoteReference'></span>")
      end
    end
  end

  def transform_footnote_text(note)
    note["id"] = ""
    note.xpath(".//div").each { |div| div.replace(div.children) }
    note.xpath(".//aside | .//p").each do |p|
      p.name = "p"
      p["class"] = "MsoFootnoteText"
    end
    note.remove
  end

  def footnote?(elem)
    elem["epub:type"]&.casecmp("footnote")&.zero? ||
      elem["class"]&.casecmp("footnote")&.zero?
  end

  def set_footnote_link_attrs(elem, idx)
    elem["style"] = "mso-footnote-id:ftn#{idx}"
    elem["href"] = "#_ftn#{idx}"
    elem["name"] = "_ftnref#{idx}"
    elem["title"] = ""
  end

  # We expect that the content of the footnote text received is one or
  # more text containers, p or aside or div (which we have already converted
  # to p). We do not expect any <a name> or links back to text; if they
  # are present in the HTML, they need to have been cleaned out before
  # passing to this gem
  def footnote_cleanup(docxml)
    docxml.xpath('//div[@style="mso-element:footnote"]/a')
      .each do |x|
      n = x.next_element
      n&.children&.first&.add_previous_sibling(x.remove)
    end
    docxml
  end
end
