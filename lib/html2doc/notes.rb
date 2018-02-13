require "uuidtools"
require "nokogiri"

module Html2Doc

  def self.footnotes(docxml)
    i, fn = 1, []
    docxml.xpath("//a").each do |a|
      next unless process_footnote_link(docxml, a, i, fn)
      i += 1
    end
    process_footnote_texts(docxml, fn)
  end

  def self.process_footnote_texts(docxml, footnotes)
    body = docxml.at("//body")
    list = body.add_child("<div style='mso-element:footnote-list'/>")
    footnotes.each_with_index do |f, i|
      fn = list.first.add_child(footnote_container(i+1))
      f.parent = fn.first
      footnote_div_to_p(f)
    end
    footnote_cleanup(docxml)
  end

  def self.footnote_div_to_p(f)
    if %w{div aside}.include? f.name
      if f.at(".//p")
        f = f.replace(f.children)
      else
        f.name = "p"
        f["class"] = "MsoFootnoteText"
      end
    end
  end

  def self.footnote_container(i)
    <<~DIV
      <div style='mso-element:footnote' id='ftn#{i}'>
        <a style='mso-footnote-id:ftn#{i}' href=#_ftn#{i}' 
           name='_ftnref#{i}' title='' id='_ftnref#{i}'><span 
           class='MsoFootnoteReference'><span 
           style='mso-special-character:footnote'></span></span></div>
    DIV
  end

  def self.process_footnote_link(docxml, a, i, fn)
    return false unless is_footnote(a)
    href = a["href"].gsub(/^#/, "")
    note = docxml.at("//*[@name = '#{href}' or @id = '#{href}']")
    return false if note.nil?
    set_footnote_link_attrs(a, i)
    a.children = "<span class='MsoFootnoteReference'>"\
      "<span style='mso-special-character:footnote'/></span>"
    fn << transform_footnote_text(note)
  end

  def self.transform_footnote_text(note)
    note["id"] = ""
    note.xpath(".//div").each { |div| div = div.replace(div.children) }
    note.xpath(".//aside | .//p").each do |p|
      p.name = "p"
      p["class"] = "MsoFootnoteText"
    end
    note.remove
  end

  def self.is_footnote(a)
    a["epub:type"]&.casecmp("footnote") == 0 ||
      a["class"]&.casecmp("footnote") == 0
  end

  def self.set_footnote_link_attrs(a, i)
    a["style"] = "mso-footnote-id:ftn#{i}"
    a["href"] = "#_ftn#{i}"
    a["name"] = "_ftnref#{i}"
    a["title"] = ""
  end

  # We expect that the content of the footnote text received is one or
  # more text containers, p or aside or div (which we have already converted
  # to p). We do not expect any <a name> or links back to text; if they
  # are present in the HTML, they need to have been cleaned out before
  # passing to this gem
  def self.footnote_cleanup(docxml)
    docxml.xpath('//div[@style="mso-element:footnote"]/a').
      each do |x|
      n = x.next_element
      n&.children&.first&.add_previous_sibling(x.remove)
    end
    docxml
  end




end
