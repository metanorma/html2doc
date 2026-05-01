require "uniword"
require "set"

class Html2Doc
  class DocxConverter
    # XML namespaces used in html2doc's cleaned output
    XHTML_NS = "http://www.w3.org/1999/xhtml"
    MATHML_NS = "http://www.w3.org/1998/Math/MathML"
    OOXML_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def initialize(options = {})
      @options = options
      @stylesheet = options[:stylesheet]
      @filename = options[:filename]
      @imagedir = options[:imagedir]
      @liststyles = options[:liststyles]
      @header_file = options[:header_file]
      @footnotes = nil
      @endnotes = nil
      @bookmark_id = 0
      @hyperlink_id = 0
      @hyperlinks = []
      @header_footer_rid = 0
      @body_content = []
      @section_index = -1  # -1 = not yet in any section; 0 = WordSection1 (cover page)
    end

    def convert(docxml)
      @docxml = docxml

      # Create document root early (needed for image registration)
      @document = Uniword::Wordprocessingml::DocumentRoot.new

      # Extract footnotes before body conversion (Phase 5)
      @footnotes = extract_footnotes(docxml)

      # Convert body content
      body = convert_body(docxml)
      @document.body = body

      # Load ISO styles and numbering
      styles = load_styles
      numbering = convert_numbering(docxml)

      # Parse headers/footers from header.html (Phase 7)
      header_sections = parse_header_file

      # Apply section properties (uses @section_boundaries from convert_body,
      # plus header_sections for header/footer references)
      apply_sections(body, header_sections)

      # Store header/footer parts for serialization and wire references
      if header_sections && !header_sections.empty?
        parts = build_header_footer_parts(header_sections)
        @document.header_footer_parts = parts
        wire_header_footer_refs(body, parts)
      end

      # Assemble and return package
      assemble_package(@document, styles, numbering)
    end

    def save_to_file(package, filename)
      package.to_file(filename)
    end

    private

    # ============================================================
    # Phase 3: Body, Paragraph, and Run conversion
    # ============================================================

    def convert_body(docxml)
      body = Uniword::Wordprocessingml::Body.new
      body_nodes = docxml.at_css("body")
      return body unless body_nodes

      # Build ordered content list: each entry is [:p, para] or [:tbl, table]
      @body_content = []
      @section_boundaries = []  # indices into @body_content for start of each section

      body_nodes.children.each do |node|
        next if node.text? && node.text.strip.empty?

        case node.name
        when "p"
          add_body_paragraph(node)
        when "h1", "h2", "h3", "h4", "h5", "h6"
          add_body_paragraph(convert_heading(node))
        when "table"
          add_body_table(convert_table(node))
        when "div"
          process_div_with_sections(node)
        when "ul", "ol"
          convert_list_items(node).each { |p| add_body_paragraph(p) }
        when "oMathPara"
          add_body_paragraph(convert_math_block(node))
        when "oMath"
          add_body_paragraph(convert_math_block(node))
        when "br"
          handle_body_br(node)
        else
          add_body_paragraph(convert_paragraph(node)) if node.element?
        end
      end

      # Post-process: merge page-break paragraphs with the next paragraph
      merge_page_break_paragraphs

      # Populate body from ordered content
      @body_content.each do |type, obj|
        case type
        when :p  then body.paragraphs << obj
        when :tbl then body.tables << obj
        end
      end

      # Build element_order for correct interleaved serialization
      unless @body_content.empty?
        body.element_order = @body_content.map do |type, _|
          Lutaml::Xml::Element.new("Element", type == :p ? "p" : "tbl")
        end
      end

      body
    end

    # Process a div, detecting WordSection boundaries
    def process_div_with_sections(element)
      css_class = element["class"].to_s
      if (m = css_class.match(/WordSection(\d+)/))
        @section_boundaries << @body_content.size
        @section_index = m[1].to_i - 1  # 0-based: WordSection1 = 0 (cover page)
      end

      convert_div(element).each { |item| @body_content << item }
    end

    def handle_body_br(element)
      if element["class"] == "section"
        # Section break between WordSection divs: create an empty paragraph
        # (Word adds a paragraph mark at section boundaries to carry sectPr)
        add_body_paragraph(Uniword::Wordprocessingml::Paragraph.new)
        return
      end

      # Regular line break: create a standalone paragraph
      add_body_paragraph(create_break_paragraph(element))
    end

    def add_body_paragraph(para_or_node)
      if para_or_node.is_a?(Nokogiri::XML::Node)
        para = if toc_paragraph?(para_or_node)
                 convert_toc_paragraph(para_or_node)
               else
                 convert_paragraph(para_or_node)
               end
      else
        para = para_or_node
      end
      @body_content << [:p, para]
    end

    def toc_paragraph?(element)
      css_class = element["class"].to_s
      css_class.match?(/^MsoToc\d+$/)
    end

    def empty_paragraph?(para)
      return false if para.properties
      return false if para.runs.any?
      return false if para.hyperlinks.any?
      return false if para.o_maths.any?
      return false if para.o_math_paras.any?
      true
    end

    def add_body_table(table)
      @body_content << [:tbl, table]
    end

    # Post-process: merge standalone page-break paragraphs into the next paragraph.
    # In Word, page breaks are runs within a paragraph, not separate paragraphs.
    def merge_page_break_paragraphs
      i = 0
      while i < @body_content.size - 1
        item = @body_content[i]
        next_item = @body_content[i + 1]

        # Only process paragraph items that contain ONLY a page break run
        if item.first == :p && page_break_only_paragraph?(item.last) && next_item.first == :p
          br_run = item.last.runs.first
          target_para = next_item.last
          target_para.runs.unshift(br_run)
          @body_content.delete_at(i)
          # Adjust section boundary indices
          @section_boundaries.map! { |idx| idx > i ? idx - 1 : idx }
          # Don't increment i — check the same position again
        else
          i += 1
        end
      end
    end

    def page_break_only_paragraph?(para)
      return false if para.hyperlinks.any? || para.o_maths.any? || para.o_math_paras.any?
      return false if para.runs.size != 1
      run = para.runs.first
      run.break && run.break.type == "page"
    end

    def convert_paragraph(element)
      # Delegate to TOC-specific converter for TOC paragraphs
      return convert_toc_paragraph(element) if toc_paragraph?(element)

      para = Uniword::Wordprocessingml::Paragraph.new

      # Set style from class attribute (only for block-level elements, not cover page)
      # Inline elements like <a class="TableFootnoteRef"> should not set paragraph style
      css_class = element["class"]
      if css_class && @section_index != 0 && block_level_element?(element)
        style_id = resolve_style(css_class)
        # Skip "Normal" style — reference output relies on default (no explicit pStyle)
        if style_id && style_id != "Normal"
          para.properties = Uniword::Wordprocessingml::ParagraphProperties.new(
            style: style_id
          )
        end
      end

      # Parse inline style for paragraph properties
      parse_paragraph_style(para, element["style"]) if element["style"]

      # Convert inline content (runs, bookmarks, hyperlinks, math)
      convert_inline_content(para, element)

      # Strip explicit bold from paragraphs with bold-included styles
      if para.properties&.style && BOLD_STYLES.include?(para.properties.style.value)
        strip_bold_from_runs(para)
      end

      para
    end

    # Block-level HTML elements that should set paragraph styles.
    # Inline elements (a, span, strong, etc.) should not.
    BLOCK_ELEMENTS = %w[p div h1 h2 h3 h4 h5 h6 li dt dd blockquote pre section article aside].freeze

    # Styles that include bold formatting — explicit bold runs are redundant.
    BOLD_STYLES = %w[Heading1 Heading2 Heading3 Heading4 Heading5 Heading6
                     TOC1 TOC2 TOC3 TOC4 TOC5 TOC6 TOC7 TOC8 TOC9].freeze

    def block_level_element?(element)
      BLOCK_ELEMENTS.include?(element.name)
    end

    def convert_heading(element)
      level = element.name.sub("h", "").to_i
      para = convert_paragraph(element)
      para.properties ||= Uniword::Wordprocessingml::ParagraphProperties.new

      # Use CSS class style if present, otherwise default to HeadingN.
      css_class = element["class"]
      style_id = css_class ? resolve_style(css_class) : nil
      style_id ||= "Heading#{level}"

      para.properties.style = Uniword::Properties::StyleReference.new(value: style_id)

      # Strip explicit bold only for styles that include bold formatting.
      # Annex and other document-specific styles may not include bold.
      if BOLD_STYLES.include?(style_id)
        strip_bold_from_runs(para)
      end

      para
    end

    def strip_bold_from_runs(para)
      para.runs.each do |run|
        if run.properties && run.properties.bold
          run.properties = Uniword::Wordprocessingml::RunProperties.new(
            style: run.properties.style,
            italic: run.properties.italic,
            underline: run.properties.underline,
            vertical_align: run.properties.vertical_align,
            fonts: run.properties.fonts,
            color: run.properties.color,
            size: run.properties.size,
          )
        end
      end
      para.hyperlinks.each do |hl|
        hl.runs.each do |run|
          if run.properties && run.properties.bold
            run.properties = Uniword::Wordprocessingml::RunProperties.new(
              style: run.properties.style,
              italic: run.properties.italic,
              underline: run.properties.underline,
              vertical_align: run.properties.vertical_align,
              fonts: run.properties.fonts,
              color: run.properties.color,
              size: run.properties.size,
            )
          end
        end
      end
    end

    # ============================================================
    # TOC paragraph conversion (proper field code model)
    # ============================================================

    # Convert a TOC paragraph into a proper OOXML Paragraph with field codes.
    #
    # TOC paragraphs have a specific HTML structure with Word field code markup:
    #   <p class="MsoToc1">
    #     [<span><span style="mso-element:field-begin">...</span>  (TOC[0] only)
    #      <span>TOC \o "1-2" \h \z \u</span>
    #      <span style="mso-element:field-separator"></span></span>]
    #     <span class="MsoHyperlink"><a href="#_Toc...">
    #       Entry text
    #       <span class="MsoTocTextSpan">[tab/field/instr/separator/end]</span>
    #       ...
    #     </a></span>
    #   </p>
    #
    # The OOXML model uses Run objects with field_char, instr_text, and tab
    # attributes to represent this structure.
    def convert_toc_paragraph(element)
      para = Uniword::Wordprocessingml::Paragraph.new

      css_class = element["class"]
      style_id = resolve_style(css_class)
      if style_id
        para.properties = Uniword::Wordprocessingml::ParagraphProperties.new(
          style: style_id
        )
      end

      element.children.each do |child|
        next if child.text? && child.text.strip.empty?

        if child.element?
          css_cls = child["class"].to_s

          if css_cls == "MsoHyperlink"
            # The main TOC entry: MsoHyperlink span wrapping an <a> element
            process_toc_hyperlink(para, child)
          elsif field_wrapper_span?(child)
            # Top-level field code wrapper (TOC[0] only): contains field-begin,
            # instruction text, field-separator as child spans
            process_toc_field_wrapper(para, child)
          else
            # Skip elements containing nested <p> (e.g., trailing NBSP paragraphs
            # in the last TOC entry's field-end span) — handled at div level
            next if child.css("p").any?
            convert_inline_content(para, child)
          end
        end
      end

      para
    end

    # Process the TOC field code wrapper (field-begin → instrText → field-separate).
    # This only appears in the first TOC paragraph (TOC[0]).
    def process_toc_field_wrapper(para, element)
      element.children.each do |child|
        span_style = child["style"].to_s if child.element?

        if child.element? && field_begin_style?(span_style)
          run = Uniword::Wordprocessingml::Run.new
          run.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "begin")
          para.runs << run
        elsif child.element? && field_sep_style?(span_style)
          run = Uniword::Wordprocessingml::Run.new
          run.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "separate")
          para.runs << run
        elsif child.text? && !child.text.strip.empty?
          # Field instruction text (e.g., "TOC \o \"1-2\" \h \z \u")
          text = child.text.strip
          run = Uniword::Wordprocessingml::Run.new
          run.instr_text = Uniword::Wordprocessingml::InstrText.new(text: " #{text} ")
          para.runs << run
        elsif child.element? && span_style.include?("mso-spacerun")
          # Spacerun span — skip (just whitespace formatting)
        end
      end
    end

    # Process the MsoHyperlink span containing the TOC entry.
    # Finds the <a> element inside and builds a Hyperlink with properly
    # structured runs (text, tab, field chars, instr text).
    def process_toc_hyperlink(para, element)
      a_el = element.at_css("a")
      return unless a_el

      href = a_el["href"]
      return unless href

      runs = build_toc_hyperlink_runs(a_el)

      anchor = href.sub(/^#/, "")
      hl = Uniword::Wordprocessingml::Hyperlink.new(
        anchor: anchor,
        runs: runs
      )
      store_hyperlink_position(para, hl)
      para.hyperlinks << hl
    end

    # Build the runs for a TOC hyperlink from the <a> element's children.
    #
    # Maps the HTML structure to OOXML runs:
    # - Text nodes → Run with Hyperlink style
    # - <span class="MsoTocTextSpan"> with <span style="mso-tab-count:*"> → Run with msotoctextspan1, tab
    # - <span class="MsoTocTextSpan"> with <span style="mso-element:field-begin"> → Run with msotoctextspan1, field_char
    # - <span class="MsoTocTextSpan"> with field instruction text → Run with msotoctextspan1, instr_text
    # - <span class="MsoTocTextSpan"> with <span style="mso-element:field-separator"> → Run with msotoctextspan1, field_char
    # - <span class="MsoTocTextSpan"> with page number text → Run with msotoctextspan1, text
    # - <span class="MsoTocTextSpan"> with <span style="mso-element:field-end"> → Run with msotoctextspan1, field_char
    def build_toc_hyperlink_runs(a_element)
      runs = []

      a_element.children.each do |child|
        next if child.text? && child.text.strip.empty?

        if child.text?
          # Direct text (e.g., "Introduction")
          text = child.text
          next if text.empty?
          runs << create_toc_text_run(text, "Hyperlink")
        elsif child.element?
          css_cls = child["class"].to_s

          if css_cls == "MsoTocTextSpan"
            runs << build_toc_text_span_run(child)
          elsif child.name == "b" || child.name == "strong"
            # Bold text in TOC hyperlink (e.g., annex titles)
            child.children.each do |grandchild|
              next if grandchild.text? && grandchild.text.strip.empty?
              if grandchild.text?
                runs << create_toc_text_run(grandchild.text, "Hyperlink", bold: true)
              end
            end
          else
            # Other spans (e.g., <span style="mso-no-proof:yes"> wrapper)
            # Process their children recursively
            child.children.each do |grandchild|
              next if grandchild.text? && grandchild.text.strip.empty?

              if grandchild.text?
                runs << create_toc_text_run(grandchild.text, "Hyperlink")
              elsif grandchild.element?
                gc_cls = grandchild["class"].to_s
                if gc_cls == "MsoTocTextSpan"
                  runs << build_toc_text_span_run(grandchild)
                elsif grandchild.name == "a"
                  # Nested <a> inside wrapper span — process its children
                  grandchild.children.each do |inner|
                    next if inner.text? && inner.text.strip.empty?
                    if inner.text?
                      runs << create_toc_text_run(inner.text, "Hyperlink")
                    elsif inner.element? && inner["class"].to_s == "MsoTocTextSpan"
                      runs << build_toc_text_span_run(inner)
                    end
                  end
                end
              end
            end
          end
        end
      end

      runs
    end

    # Build a single Run from a <span class="MsoTocTextSpan"> element.
    # Inspects children to determine the run type (tab, field_char, instr_text, text).
    def build_toc_text_span_run(span)
      run = Uniword::Wordprocessingml::Run.new
      apply_run_style(run, "msotoctextspan1")

      span.children.each do |child|
        next if child.text? && child.text.strip.empty?

        if child.element?
          child_style = child["style"].to_s

          if field_begin_style?(child_style)
            run.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "begin")
            return run
          elsif field_sep_style?(child_style)
            run.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "separate")
            return run
          elsif field_end_style?(child_style)
            run.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "end")
            return run
          elsif tab_count_style?(child_style)
            run.tab = Uniword::Wordprocessingml::Tab.new
            return run
          end
        elsif child.text?
          text = child.text.strip
          if pagerref_instruction?(text)
            run.instr_text = Uniword::Wordprocessingml::InstrText.new(text: " #{text} ")
          else
            run.text = text
          end
          return run
        end
      end

      # Empty MsoTocTextSpan
      run
    end

    def create_toc_text_run(text, style_id, bold: false)
      run = Uniword::Wordprocessingml::Run.new
      run.text = text
      apply_run_style(run, style_id)
      if bold
        run.properties.bold = Uniword::Properties::Bold.new
      end
      run
    end

    def field_begin_style?(style)
      style.include?("mso-element:field-begin")
    end

    def field_sep_style?(style)
      style.include?("mso-element:field-separator")
    end

    def field_end_style?(style)
      style.include?("mso-element:field-end")
    end

    def tab_count_style?(style)
      style&.include?("mso-tab-count")
    end

    def pagerref_instruction?(text)
      text.match?(/\APAGEREF\s/)
    end

    # Check if a span contains field-begin or field-separator children,
    # indicating it's the TOC field code wrapper (only in the first TOC paragraph).
    def field_wrapper_span?(element)
      return false unless element.element?

      element.children.any? do |child|
        next false unless child.element?
        style = child["style"].to_s
        field_begin_style?(style) || field_sep_style?(style)
      end
    end

    def convert_div(element)
      items = []
      children = element.children.to_a
      i = 0

      while i < children.size
        child = children[i]
        if child.text? && child.text.strip.empty?
          i += 1
          next
        end

        case child.name
        when "p"
          # Skip <p> elements with no children at all (empty placeholder <p></p>)
          unless child.children.empty?
            items << [:p, convert_paragraph(child)]
            # TOC paragraphs may contain nested <p> elements (e.g., NBSP
            # paragraph inside field-end span) — extract them separately
            if toc_paragraph?(child)
              child.css("p").each do |nested_p|
                items << [:p, convert_paragraph(nested_p)]
              end
            end
          end
        when "h1", "h2", "h3", "h4", "h5", "h6"
          items << [:p, convert_heading(child)]
        when "table"
          items << [:tbl, convert_table(child)]
        when "div"
          items.concat(convert_div(child))
        when "ul", "ol"
          convert_list_items(child).each { |p| items << [:p, p] }
        when "oMathPara"
          items << [:p, convert_math_block(child)]
        when "oMath"
          # Collect oMath and subsequent inline content into one paragraph
          para = Uniword::Wordprocessingml::Paragraph.new
          convert_math_inline(para, child)
          prev_was_omath = true
          i += 1
          while i < children.size
            sib = children[i]
            if sib.text?
              text = sib.text
              if text.strip.empty?
                if prev_was_omath
                  run = create_run(text)
                  apply_run_style(run, "stem")
                  para.runs << run
                end
                i += 1
                next
              end
              para.runs << create_run(text)
              prev_was_omath = false
            elsif sib.element?
              case sib.name
              when "span"
                convert_span(para, sib)
              when "a"
                convert_link(para, sib)
              when "br"
                break
              else
                break # block element — stop collecting inline content
              end
              prev_was_omath = false
            end
            i += 1
          end
          items << [:p, para]
          next # don't increment i again
        when "br"
          if child["class"] == "section"
            i += 1
            next
          end
          items << [:p, create_break_paragraph(child)]
        when "img"
          # Standalone image wrapped in a paragraph
          para = Uniword::Wordprocessingml::Paragraph.new
          convert_image(para, child)
          items << [:p, para]
        when "dl"
          items.concat(convert_dl(child))
        when "a"
          # Bookmark anchor — skip at block level (no visible content)
          if child.text.strip.empty?
            i += 1
            next
          end
          # Non-empty link at block level — treat as paragraph
          items << [:p, convert_paragraph(child)]
        when "title"
          # Admonition title — redundant (content is already in the following <p>)
          i += 1
          next
        when "span", "aside"
          # Inline/structural elements at block level — skip if no visible content
          if child.text.strip.empty? && child.css("img").empty?
            i += 1
            next
          end
          items << [:p, convert_paragraph(child)]
        else
          if child.element?
            items << [:p, convert_paragraph(child)]
          end
        end
        i += 1
      end
      items
    end

    # Convert definition list (dl/dt/dd) to paragraphs
    def convert_dl(dl_element)
      items = []
      dl_element.children.each do |child|
        next if child.text? && child.text.strip.empty?
        next unless child.element?
        # Skip empty anchor elements inside dl
        next if child.name == "a" && child.text.strip.empty?

        items << [:p, convert_paragraph(child)]
      end
      items
    end

    def convert_list_items(element)
      paragraphs = []
      element.css("li").each do |li|
        para = convert_paragraph(li)
        paragraphs << para
      end
      paragraphs
    end

    # ============================================================
    # Inline content conversion (runs, bookmarks, hyperlinks, math)
    # ============================================================

    def convert_inline_content(para, element)
      prev_was_omath = false
      after_br = false
      element.children.each do |child|
        case
        when child.text?
          text = child.text
          next if text.empty?
          # Skip whitespace-only text nodes that are indentation (contain newlines)
          next if text.strip.empty? && text.include?("\n")
          # Text nodes following <br> may have leading newlines (HTML formatting)
          # that should be stripped — the <br> already provides the line break
          text = text.lstrip if after_br && text.start_with?("\n")
          next if text.empty?
          run = create_run(text)
          # Whitespace-only text right after oMath is a spacer run that
          # carries the "stem" character style in the reference output.
          if prev_was_omath && text.strip.empty?
            apply_run_style(run, "stem")
          end
          para.runs << run
          prev_was_omath = false
          after_br = false
        when child.element?
          after_br = child.name == "br"
          prev_was_omath = child.name == "oMath"
          convert_element(para, child)
        end
      end
    end

    def convert_element(para, element)
      case element.name
      when "strong", "b"
        convert_formatted_element(para, element, bold: true)
      when "em", "i"
        convert_formatted_element(para, element, italic: true)
      when "u"
        convert_formatted_element(para, element, underline: "single")
      when "s", "strike", "del"
        convert_formatted_element(para, element, strike: true)
      when "sub"
        convert_formatted_element(para, element, vertical_align: "subscript")
      when "sup"
        convert_formatted_element(para, element, vertical_align: "superscript")
      when "span"
        convert_span(para, element)
      when "a"
        convert_link(para, element)
      when "br"
        para.runs << create_break_run(element)
      when "img"
        convert_image(para, element)
      when "oMath"
        convert_math_inline(para, element)
      when "table"
        # Nested table - handled at body level
        nil
      else
        # Unknown element - recurse into children
        convert_inline_content(para, element)
      end
    end

    def convert_formatted_element(para, element, **formatting)
      # Collect any nested formatting and apply it all at once
      merged = collect_formatting(element, formatting)

      run_count_before = para.runs.size

      element.children.each do |child|
        case
        when child.text?
          text = child.text
          next if text.empty?
          # Skip whitespace-only text nodes that are indentation (contain newlines)
          next if text.strip.empty? && text.include?("\n")
          para.runs << create_run(text, merged)
        when child.element?
          case child.name
          when "strong", "b"
            convert_formatted_element(para, child, **merged.merge(bold: true))
          when "em", "i"
            convert_formatted_element(para, child, **merged.merge(italic: true))
          when "u"
            convert_formatted_element(para, child, **merged.merge(underline: "single"))
          when "span"
            # Parse span's CSS and merge with existing formatting
            span_fmt = parse_span_style(child["style"])
            convert_formatted_element(para, child, **merged.merge(span_fmt))
          else
            convert_element(para, child)
          end
        end
      end

      # If no runs were created but formatting is present, emit an empty run
      # to preserve empty formatting tags (e.g., <b><span><p></p></span></b>)
      if para.runs.size == run_count_before && !merged.empty?
        para.runs << create_run("", merged)
      end
    end

    def convert_span(para, element)
      # mso-tab-count spans: emit <w:tab/> run (Word uses these as tab stops)
      if tab_count_style?(element["style"])
        run = Uniword::Wordprocessingml::Run.new
        run.tab = Uniword::Wordprocessingml::Tab.new
        para.runs << run
        return
      end

      # Check for CSS class → run style mapping (e.g., MsoTocTextSpan → msotoctextspan1)
      css_class = element["class"]
      if css_class
        style_id = resolve_style(css_class)
        if style_id
          has_content = false
          # Apply this run style to all text content within this span
          element.children.each do |child|
            case
            when child.text?
              text = child.text
              next if text.empty?
              has_content = true
              run = create_run(text)
              apply_run_style(run, style_id)
              para.runs << run
            when child.element?
              has_content = true
              convert_element_with_style(para, child, style_id)
            end
          end
          # Empty styled spans still produce a run (e.g., empty MsoTocTextSpan
          # for field markers in TOC entries)
          unless has_content
            run = create_run("")
            apply_run_style(run, style_id)
            para.runs << run
          end
          return
        end
      end

      formatting = parse_span_style(element["style"])
      return convert_inline_content(para, element) if formatting.empty?

      element.children.each do |child|
        case
        when child.text?
          text = child.text
          next if text.empty?
          para.runs << create_run(text, formatting)
        when child.element?
          # Nested element inside styled span
          convert_formatted_element(para, child, **formatting)
        end
      end
    end

    def convert_element_with_style(para, element, style_id, formatting = {})
      case
      when element.text?
        text = element.text
        return if text.empty?
        merged = formatting.empty? ? {} : formatting
        run = create_run(text, merged)
        apply_run_style(run, style_id)
        para.runs << run
      when element.element?
        # Accumulate formatting from wrapper elements
        case element.name
        when "strong", "b"
          element.children.each { |child| convert_element_with_style(para, child, style_id, formatting.merge(bold: true)) }
        when "em", "i"
          element.children.each { |child| convert_element_with_style(para, child, style_id, formatting.merge(italic: true)) }
        when "u"
          element.children.each { |child| convert_element_with_style(para, child, style_id, formatting.merge(underline: "single")) }
        when "span"
          span_fmt = parse_span_style(element["style"])
          element.children.each { |child| convert_element_with_style(para, child, style_id, formatting.merge(span_fmt)) }
        else
          # For other nested elements, recurse with current formatting
          element.children.each do |child|
            convert_element_with_style(para, child, style_id, formatting)
          end
        end
      end
    end

    def apply_run_style(run, style_id)
      run.properties ||= Uniword::Wordprocessingml::RunProperties.new
      run.properties.style = Uniword::Properties::RunStyleReference.new(value: style_id)
    end

    def convert_link(para, element)
      href = element["href"]

      # Bookmark anchors: <a name="..."> — skip bookmark creation.
      # The reference output has minimal bookmarks (1 total), not one per HTML id.
      if element["name"] && !href
        convert_inline_content(para, element)
        return
      end

      return unless href

      # Footnote reference markers (Phase 5)
      if element["data-footnote-ref"]
        ref_id = element["data-footnote-ref"]
        run = Uniword::Wordprocessingml::Run.new
        run.footnote_reference = Uniword::Wordprocessingml::FootnoteReference.new(
          id: ref_id
        )
        run.properties = Uniword::Wordprocessingml::RunProperties.new(
          style: Uniword::Properties::RunStyleReference.new(value: "FootnoteReference")
        )
        para.runs << run
        return
      end

      # Regular hyperlink
      runs = []
      element.children.each do |child|
        case
        when child.text?
          text = child.text
          next if text.empty?
          # Skip whitespace-only text nodes — these are indentation artifacts
          # between styled spans (e.g., MsoTocTextSpan in TOC entries).
          next if text.strip.empty?
          r = create_run(text)
          apply_run_style(r, "Hyperlink")
          runs << r
        when child.element?
          # Create a temp paragraph to collect runs
          temp = Uniword::Wordprocessingml::Paragraph.new
          convert_element(temp, child)
          temp.runs.each do |r|
            # Only apply Hyperlink style if no explicit style was set
            # (e.g., MsoTocTextSpan spans already have msotoctextspan1)
            unless r.properties&.style
              apply_run_style(r, "Hyperlink")
            end
            runs << r
          end
        end
      end

      if href.start_with?("#")
        # Internal anchor link
        hl = Uniword::Wordprocessingml::Hyperlink.new(
          anchor: href.sub(/^#/, ""),
          runs: runs
        )
        store_hyperlink_position(para, hl)
        para.hyperlinks << hl
      else
        # External URL link
        r_id = "rIdLink#{@hyperlink_id += 1}"
        @hyperlinks << { id: r_id, url: href }
        hl = Uniword::Wordprocessingml::Hyperlink.new(
          id: r_id,
          runs: runs
        )
        store_hyperlink_position(para, hl)
        para.hyperlinks << hl
      end
    end

    # Store the current run count as the hyperlink's position marker.
    # This allows the MHTML renderer to interleave runs and hyperlinks
    # in the correct order (hyperlink follows after N runs).
    def store_hyperlink_position(para, hyperlink)
      pos = para.runs.size
      hyperlink.instance_variable_set(:@_run_position, pos)
    end

    # ============================================================
    # Table conversion (basic)
    # ============================================================

    def convert_table(element)
      table = Uniword::Wordprocessingml::Table.new
      rows_data = []

      # First pass: collect cell data including colspan/rowspan
      element.css("tr").each do |tr|
        row_cells = []
        tr.css("td, th").each do |cell|
          row_cells << {
            element: cell,
            colspan: cell["colspan"]&.to_i || 1,
            rowspan: cell["rowspan"]&.to_i || 1,
            is_header: cell.name == "th",
          }
        end
        rows_data << row_cells
      end

      # Track vertical merge state: col_index → remaining rowspan count
      vmerge_state = {}

      # Second pass: build OOXML table rows
      rows_data.each_with_index do |row_cells, _row_idx|
        row = Uniword::Wordprocessingml::TableRow.new
        col_offset = 0

        row_cells.each do |cell_data|
          # Skip columns occupied by ongoing vertical merges
          while vmerge_state[col_offset]
            # Emit a continuation cell for this vertically-merged column
            tc = Uniword::Wordprocessingml::TableCell.new
            apply_cell_merge(tc, :continue)
            tc.paragraphs = [Uniword::Wordprocessingml::Paragraph.new]
            row.cells << tc

            vmerge_state[col_offset] -= 1
            vmerge_state.delete(col_offset) if vmerge_state[col_offset] <= 0
            col_offset += 1
          end

          cell_el = cell_data[:element]
          tc = Uniword::Wordprocessingml::TableCell.new
          tc.header = cell_data[:is_header]
          tc.paragraphs = convert_cell_content(cell_el, skip_bold: cell_data[:is_header])

          # Apply gridSpan for colspan
          if cell_data[:colspan] > 1
            apply_cell_merge(tc, :grid_span, cell_data[:colspan])
          end

          # Apply vMerge restart for rowspan
          if cell_data[:rowspan] > 1
            apply_cell_merge(tc, :restart)
            # Register vertical merge for subsequent rows
            cell_data[:colspan].times do |i|
              vmerge_state[col_offset + i] = cell_data[:rowspan] - 1
            end
          end

          row.cells << tc
          col_offset += cell_data[:colspan]
        end

        # Fill remaining vmerge continuation cells at the end of the row
        while vmerge_state[col_offset]
          tc = Uniword::Wordprocessingml::TableCell.new
          apply_cell_merge(tc, :continue)
          tc.paragraphs = [Uniword::Wordprocessingml::Paragraph.new]
          row.cells << tc

          vmerge_state[col_offset] -= 1
          vmerge_state.delete(col_offset) if vmerge_state[col_offset] <= 0
          col_offset += 1
        end

        table.rows << row
      end

      table
    end

    # Apply merge properties (gridSpan or vMerge) to a TableCell.
    def apply_cell_merge(tc, type, span = nil)
      tc.properties ||= Uniword::Wordprocessingml::TableCellProperties.new
      case type
      when :grid_span
        tc.properties.grid_span = Uniword::Wordprocessingml::ValInt.new(value: span)
      when :restart
        tc.properties.v_merge = Uniword::Wordprocessingml::ValInt.new(value: "restart")
      when :continue
        tc.properties.v_merge = Uniword::Wordprocessingml::ValInt.new(value: "continue")
      end
    end

    # Convert cell children into paragraphs.
    # Bare text/inline elements get wrapped in a single paragraph.
    # Block elements (p, div, oMath, etc.) each produce their own paragraph.
    def convert_cell_content(cell, skip_bold: false)
      paras = []
      current_para = nil
      prev_was_omath = false

      # Check cell-level formatting
      cell_style = cell["style"]
      cell_bold = !skip_bold && (cell_style&.include?("font-weight:bold") || cell_style&.include?("font-weight: bold"))
      cell_align = cell["align"]

      cell.children.each do |child|
        next if child.text? && child.text.strip.empty? && !prev_was_omath

        case child.name
        when "p"
          prev_was_omath = false
          flush_inline_para(paras, current_para)
          current_para = nil
          para = convert_paragraph(child)
          paras << para unless empty_paragraph?(para)
        when "oMathPara", "oMath"
          prev_was_omath = true
          current_para ||= Uniword::Wordprocessingml::Paragraph.new
          convert_math_inline(current_para, child)
        when "div"
          prev_was_omath = false
          flush_inline_para(paras, current_para)
          current_para = nil
          convert_div(child).each { |_, p| paras << p }
        when "br"
          prev_was_omath = false
          current_para ||= Uniword::Wordprocessingml::Paragraph.new
          current_para.runs << create_break_run(child)
        else
          if child.element?
            prev_was_omath = false
            if block_level_element?(child)
              flush_inline_para(paras, current_para)
              current_para = nil
              para = convert_paragraph(child)
              paras << para unless empty_paragraph?(para)
            else
              current_para ||= Uniword::Wordprocessingml::Paragraph.new
              convert_element(current_para, child)
            end
          elsif child.text?
            text = child.text
            if text.strip.empty? && !prev_was_omath
              next
            end
            current_para ||= Uniword::Wordprocessingml::Paragraph.new
            run = create_run(text)
            if prev_was_omath && text.strip.empty?
              apply_run_style(run, "stem")
            end
            current_para.runs << run
            prev_was_omath = false
          end
        end
      end

      flush_inline_para(paras, current_para)

      # Apply cell-level formatting to all paragraphs
      if cell_bold || cell_align
        paras.each { |p| apply_cell_formatting(p, cell_bold, cell_align) }
      end

      paras << Uniword::Wordprocessingml::Paragraph.new if paras.empty?
      paras
    end

    def apply_cell_formatting(para, bold, align)
      if bold
        para.runs.each do |r|
          existing = r.properties
          fmt = {}
          fmt[:bold] = true if bold
          if existing
            # Merge with existing properties
            r.properties = Uniword::Wordprocessingml::RunProperties.new(
              style: existing.style,
              bold: true,
              italic: existing.italic,
              underline: existing.underline,
              vertical_align: existing.vertical_align,
              fonts: existing.fonts,
              color: existing.color,
              size: existing.size,
            )
          else
            r.properties = Uniword::Wordprocessingml::RunProperties.new(bold: true)
          end
        end
      end
      if align && %w[left center right justify].include?(align)
        para.properties ||= Uniword::Wordprocessingml::ParagraphProperties.new
        para.properties.alignment = Uniword::Properties::Alignment.new(value: align == "justify" ? "both" : align)
      end
    end

    def flush_inline_para(paras, para)
      return unless para
      paras << para unless empty_paragraph?(para)
    end

    # ============================================================
    # Phase 7: Math (OMML passthrough)
    # ============================================================

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def inject_stem_style(xml)
      doc = Nokogiri::XML(xml)
      # Ensure namespace declarations exist on root for Uniword parsing
      root = doc.root
      unless root.namespaces.key?("xmlns:m")
        root.add_namespace("m", M_NS)
      end
      unless root.namespaces.key?("xmlns:w")
        root.add_namespace("w", W_NS)
      end

      # Find all <m:r> elements (Nokogiri uses full name "m:r" without namespace decl)
      doc.css("m\\:r").each do |mr|
        # Check if w:rPr already exists
        has_word_rpr = mr.children.any? { |c| c.name == "w:rPr" || (c.name == "rPr" && c.namespace&.href&.include?("wordprocessingml")) }
        next if has_word_rpr

        # Create w:rPr with w:rStyle using proper namespace
        rpr = Nokogiri::XML::Node.new("w:rPr", doc)
        rstyle = Nokogiri::XML::Node.new("w:rStyle", doc)
        rstyle["w:val"] = "stem"
        rpr.add_child(rstyle)
        # Find <m:t>
        mt = mr.children.find { |c| c.name == "m:t" || (c.name == "t" && c.namespace&.href&.include?("math")) }
        if mt
          mt.add_previous_sibling(rpr)
        else
          mr.add_child(rpr)
        end
      end
      # Remove <w:i/> from inside <m:ctrlPr> — reference output doesn't include
      # italic in math control properties (Plurimath adds it but reference strips it)
      doc.css("m\\:ctrlPr w\\:i").each(&:remove)

      doc.root.to_xml
    end

    def convert_math_inline(para, element)
      # Inline math: parse OMML element and add as standalone oMath
      # (not wrapped in oMathPara — reference output has only oMath)
      xml = inject_stem_style(element.to_xml)
      begin
        o_math = Uniword::Math::OMath.from_xml(xml)
        o_math.instance_variable_set(:@_run_position, para.runs.size)
        para.o_maths << o_math
      rescue StandardError => e
        warn "html2doc: failed to parse inline math: #{e.message}"
        para.runs << create_run("[math]")
      end
    end

    def convert_math_block(element)
      para = Uniword::Wordprocessingml::Paragraph.new

      xml = inject_stem_style(element.to_xml)
      begin
        o_math = Uniword::Math::OMath.from_xml(xml)
        para.o_maths << o_math
      rescue StandardError => e
        warn "html2doc: failed to parse block math: #{e.message}"
        para.runs << create_run("[math block]")
      end

      para
    end
    # ============================================================

    def create_run(text, formatting = {})
      run = Uniword::Wordprocessingml::Run.new
      # Normalize whitespace: HTML source whitespace (indentation, newlines)
      # should not appear verbatim in Word output.
      run.text = text.gsub(/\s+/, " ")
      unless formatting.empty?
        run.properties = Uniword::Wordprocessingml::RunProperties.new(**formatting)
      end
      run
    end

    def create_break_run(element)
      run = Uniword::Wordprocessingml::Run.new
      if page_break_element?(element)
        run.break = Uniword::Wordprocessingml::Break.new(type: "page")
      else
        run.break = Uniword::Wordprocessingml::Break.new
      end
      run
    end

    def create_break_paragraph(element)
      para = Uniword::Wordprocessingml::Paragraph.new
      para.runs << create_break_run(element)
      para
    end

    def page_break_element?(element)
      style = element["style"].to_s
      return true if style.include?("page-break-before")
      return true if style.include?("page-break-after")
      return true if element["clear"] == "all"
      return true if element["class"] == "section"

      false
    end

    # Recursively collect formatting from nested elements
    def collect_formatting(element, base = {})
      result = base.dup
      case element.name
      when "strong", "b" then result[:bold] = true
      when "em", "i" then result[:italic] = true
      when "u" then result[:underline] = "single"
      when "s", "strike", "del" then result[:strike] = true
      when "sub" then result[:vertical_align] = "subscript"
      when "sup" then result[:vertical_align] = "superscript"
      when "span"
        result.merge!(parse_span_style(element["style"]))
      end
      result
    end

    def parse_span_style(style_str)
      return {} unless style_str

      formatting = {}
      style_str.split(";").each do |decl|
        prop, value = decl.split(":", 2).map(&:strip)
        next unless prop && value

        case prop
        when "font-weight"
          formatting[:bold] = true if value == "bold" || value.to_i >= 700
        when "font-style"
          formatting[:italic] = true if value == "italic"
        when "text-decoration"
          formatting[:underline] = "single" if value.include?("underline")
          formatting[:strike] = true if value.include?("line-through")
        when "color"
          formatting[:color] = parse_color(value)
        when "font-size"
          formatting[:size] = parse_font_size(value)
        when "font-family"
          formatting[:font] = value.split(",").first.strip.gsub(/['"]/, "")
        when "background", "background-color"
          formatting[:shading_fill] = parse_color(value)
        when "vertical-align"
          if value == "super"
            formatting[:vertical_align] = "superscript"
          elsif value == "sub"
            formatting[:vertical_align] = "subscript"
          end
        end
      end
      formatting
    end

    def parse_color(value)
      value = value.strip
      if value.start_with?("#")
        value.sub(/^#/, "").upcase
      elsif m = value.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/)
        "%02X%02X%02X" % [m[1].to_i, m[2].to_i, m[3].to_i]
      else
        value
      end
    end

    def parse_font_size(value)
      value = value.strip
      if m = value.match(/^([\d.]+)\s*pt/)
        (m[1].to_f * 2).to_i  # pt to half-points
      elsif m = value.match(/^([\d.]+)\s*px/)
        (m[1].to_f * 1.5).to_i  # approximate px to half-points
      else
        nil
      end
    end

    def parse_paragraph_style(para, style_str)
      return unless style_str

      para.properties ||= Uniword::Wordprocessingml::ParagraphProperties.new
      props = para.properties

      style_str.split(";").each do |decl|
        prop, value = decl.split(":", 2).map(&:strip)
        next unless prop && value

        case prop
        when "text-align"
          case value
          when "left" then props.alignment = Uniword::Properties::Alignment.new(value: "left")
          when "center" then props.alignment = Uniword::Properties::Alignment.new(value: "center")
          when "right" then props.alignment = Uniword::Properties::Alignment.new(value: "right")
          when "justify" then props.alignment = Uniword::Properties::Alignment.new(value: "both")
          end
        when "margin-left"
          props.indent_left = css_length_to_twips(value) if value =~ /\d/
        when "margin-right"
          props.indent_right = css_length_to_twips(value) if value =~ /\d/
        when "text-indent"
          props.indent_first_line = css_length_to_twips(value) if value =~ /\d/
        when "mso-list"
          # Parse mso-list:stylename levelN lfoN (Phase 4)
          if m = value.match(/(\w+)\s+level(\d+)\s+lfo(\d+)/)
            props.numbering_properties = Uniword::Properties::NumberingProperties.new(
              num_id: Uniword::Properties::NumberingId.new(value: m[3].to_i),
              ilvl: Uniword::Properties::NumberingLevel.new(value: m[2].to_i - 1)
            )
          end
        end
      end
    end

    def css_length_to_twips(value)
      if m = value.match(/^([\d.]+)\s*(pt|px|cm|in|mm)/)
        case m[2]
        when "pt" then (m[1].to_f * 20).to_i
        when "px" then (m[1].to_f * 15).to_i
        when "cm" then (m[1].to_f * 567).to_i
        when "in" then (m[1].to_f * 1440).to_i
        when "mm" then (m[1].to_f * 56.7).to_i
        else m[1].to_i
        end
      else
        value.to_i
      end
    end

    def resolve_style(css_class)
      # Try each class in the space-separated list
      css_class.to_s.split(/\s+/).each do |cls|
        style_id = StyleLoader.style_for_class(cls)
        return style_id if style_id
      end
      nil
    end

    # ============================================================
    # Phase 2: Style loading (delegates to StyleLoader)
    # ============================================================

    def load_styles
      styles = StyleLoader.iso_styles
      add_extra_styles(styles) if styles
      styles
    end

    # Styles used in the reference output that aren't in the ISO template
    def add_extra_styles(styles)
      extra_xml = <<~XML
        <w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:customStyle="1" w:styleId="msotoctextspan1">
          <w:name w:val="msotoctextspan1"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr>
            <w:strike w:val="0"/>
            <w:dstrike w:val="0"/>
            <w:vanish w:val="0"/>
            <w:webHidden/>
            <w:color w:val="auto"/>
            <w:u w:val="none"/>
            <w:effect w:val="none"/>
            <w:lang w:val="en-GB"/>
            <w:specVanish w:val="0"/>
          </w:rPr>
        </w:style>
        <w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:customStyle="1" w:styleId="zzmovetofollowing">
          <w:name w:val="zzmovetofollowing"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
        </w:style>
      XML
      Nokogiri::XML.fragment(extra_xml).children.each do |style_node|
        next unless style_node.element?
        style = Uniword::Wordprocessingml::Style.from_xml(style_node.to_xml)
        styles.styles << style
      end
    end

    # ============================================================
    # Phase 4: Numbering (delegates to StyleLoader)
    # ============================================================

    def convert_numbering(docxml)
      StyleLoader.iso_numbering
    end

    # ============================================================
    # Phase 5: Footnotes and endnotes
    # ============================================================

    def extract_footnotes(docxml)
      footnotes = Uniword::Wordprocessingml::Footnotes.new
      idx = 1

      # Find all footnote links
      docxml.css("a").each do |link|
        next unless footnote_link?(link)
        href = link["href"]&.sub(/^#/, "")
        next unless href

        # Find the footnote target
        # Bookmarks step may have moved id to <a name="..."> inside the container
        target = docxml.at_css("##{CSS.escape(href)}")
        if target.nil?
          named_anchor = docxml.at_css("a[name='#{CSS.escape(href)}']")
          # The real target is the parent container (aside/div), not the anchor
          target = named_anchor&.parent if named_anchor &&
            %w[aside div].include?(named_anchor.parent&.name)
        end
        next unless target

        # Create Footnote model with content
        fn = Uniword::Wordprocessingml::Footnote.new(id: idx.to_s)
        target_paras = convert_footnote_content(target)
        target_paras.each { |p| fn.paragraphs << p }
        footnotes.footnote_entries << fn

        # Replace link with placeholder for footnote reference
        link["data-footnote-ref"] = idx.to_s
        link.inner_html = ""

        # Remove the footnote target
        target.remove
        idx += 1
      end

      footnotes.footnote_entries.empty? ? nil : footnotes
    end

    def footnote_link?(elem)
      elem["epub:type"]&.casecmp("footnote")&.zero? ||
        elem["class"]&.casecmp("footnote")&.zero?
    end

    def convert_footnote_content(target)
      paragraphs = []
      # Get direct children (p, aside, div elements)
      target.children.each do |child|
        next if child.text? && child.text.strip.empty?
        # Skip bookmark anchors added by cleanup
        next if child.name == "a" && child["name"]

        case child.name
        when "p"
          para = convert_paragraph(child)
          apply_footnote_style(para)
          paragraphs << para
        when "aside", "div"
          # Recurse into nested containers
          child.children.each do |inner|
            next if inner.text? && inner.text.strip.empty?
            if inner.name == "p"
              para = convert_paragraph(inner)
              apply_footnote_style(para)
              paragraphs << para
            elsif inner.element?
              para = convert_paragraph(inner)
              apply_footnote_style(para)
              paragraphs << para
            end
          end
        else
          if child.element?
            para = convert_paragraph(child)
            apply_footnote_style(para)
            paragraphs << para
          end
        end
      end

      # Ensure at least one paragraph
      if paragraphs.empty?
        para = Uniword::Wordprocessingml::Paragraph.new
        apply_footnote_style(para)
        paragraphs << para
      end

      paragraphs
    end

    def apply_footnote_style(para)
      para.properties ||= Uniword::Wordprocessingml::ParagraphProperties.new
      para.properties.style = Uniword::Properties::StyleReference.new(
        value: "FootnoteText"
      )
    end

    # ============================================================
    # Phase 6: Images
    # ============================================================

    def convert_image(para, element)
      src = element["src"]
      return unless src && !src.empty?

      # Skip external URLs and data URIs
      if src.start_with?("http://", "https://")
        return
      end

      if src.start_with?("data:")
        return
      end

      # Strip any leading path traversal or _files directory prefix
      clean_src = src.sub(%r{^\.\./}, "")

      # Resolve local image path
      image_path = resolve_image_path(clean_src)
      unless image_path && File.exist?(image_path)
        warn "html2doc: image not found: #{src} (tried #{image_path})"
        return
      end

      # Get dimensions from HTML attributes
      width_px = element["width"]&.to_i
      height_px = element["height"]&.to_i

      # Convert px to EMU (1 px = 9525 EMU at 96 dpi)
      width_emu = width_px&.positive? ? width_px * 9525 : nil
      height_emu = height_px&.positive? ? height_px * 9525 : nil

      begin
        run = Uniword::Builder::ImageBuilder.create_run(
          @document, image_path,
          width: width_emu, height: height_emu,
          alt_text: element["alt"]
        )
        para.runs << run
      rescue StandardError => e
        warn "html2doc: failed to embed image #{src}: #{e.message}"
      end
    end

    def resolve_image_path(src)
      candidates = []

      # Try relative to imagedir
      if @imagedir
        candidates << File.join(@imagedir, src)
        # Also try just the basename (src may include subdirectory)
        candidates << File.join(@imagedir, File.basename(src))
      end

      # Try relative to filename directory
      if @filename
        base_dir = File.dirname(@filename)
        candidates << File.join(base_dir, src)
        candidates << File.join(base_dir, File.basename(src))
      end

      # Try the src as-is (absolute or relative to CWD)
      candidates << src

      candidates.compact.find { |path| File.exist?(path) }
    end

    # Rename image parts in word/media/ to sequential imageN.ext format.
    # The reference DOCX uses image1.png, image2.png, etc.
    def normalize_image_filenames(document)
      return unless document.image_parts && !document.image_parts.empty?

      img_idx = 0
      document.image_parts.each do |_r_id, data|
        img_idx += 1
        ext = File.extname(data[:target])
        data[:target] = "media/image#{img_idx}#{ext}"
      end
    end

    # ============================================================
    # Phase 8: Package assembly
    # ============================================================

    def assemble_package(document, styles, numbering)
      package = Uniword::Docx::Package.new
      package.document = document
      package.styles = styles if styles
      package.numbering = numbering if numbering
      package.footnotes = @footnotes if @footnotes
      package.font_table = StyleLoader.iso_font_table
      package.settings = StyleLoader.iso_settings
      package.theme = StyleLoader.iso_theme
      package.core_properties = Uniword::Ooxml::CoreProperties.new
      package.app_properties = Uniword::Ooxml::AppProperties.new

      # Normalize image filenames to sequential imageN.ext format
      normalize_image_filenames(document)

      # Set up document relationships for hyperlinks (needed for MHT output)
      unless @hyperlinks.empty?
        doc_rels = Uniword::Ooxml::Relationships::PackageRelationships.new
        doc_rels.relationships = @hyperlinks.map do |hl|
          Uniword::Ooxml::Relationships::Relationship.new(
            id: hl[:id],
            target: hl[:url],
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            target_mode: "External"
          )
        end
        package.document_rels = doc_rels
      end

      package
    end

    # ============================================================
    # Phase 7: Headers, Footers, and Sections
    # ============================================================

    # Parse header.html and extract per-section header/footer content.
    # Returns an array of section hashes:
    #   [{ even_header: Header, default_header: Header,
    #      even_footer: Footer, default_footer: Footer, first_header: Header, first_footer: Footer }, ...]
    def parse_header_file
      return nil unless @header_file && File.exist?(@header_file)

      html = File.read(@header_file, encoding: "utf-8")
      doc = Nokogiri::HTML(html)

      # Extract sections from div IDs: eh1/h1/ef1/f1 → section 1, etc.
      sections = {}
      doc.css("div").each do |div|
        style = div["style"].to_s
        element_type = extract_mso_element_type(style)
        next unless element_type

        section_num = extract_section_num(div["id"].to_s)
        next unless section_num

        sections[section_num] ||= {}
        key = map_id_to_header_footer_key(div["id"].to_s, element_type)
        next unless key

        # Convert the div's content to a Header or Footer object
        if element_type.include?("header")
          obj = convert_header_footer_div(div, Uniword::Wordprocessingml::Header)
        elsif element_type.include?("footer")
          obj = convert_header_footer_div(div, Uniword::Wordprocessingml::Footer)
        else
          next
        end

        sections[section_num][key] = obj if obj
      end

      # Sort by section number and return as array
      sections.sort_by { |num, _| num }.map { |_, content| content }
    end

    def extract_mso_element_type(style)
      if m = style.match(/mso-element:\s*(\S+)/)
        m[1]
      end
    end

    def extract_section_num(id)
      # IDs: eh1, h1, ef1, f1, eh2, h2, etc.
      if m = id.match(/(\d+)$/)
        m[1].to_i
      end
    end

    def map_id_to_header_footer_key(id, element_type)
      # eh1 → even header, h1 → default header, ef1 → even footer, f1 → default footer
      prefix = id.sub(/\d+$/, "")
      case prefix
      when "eh" then "even_header"
      when "h"  then "default_header"
      when "oh" then "first_header"
      when "ef" then "even_footer"
      when "f"  then "default_footer"
      when "of" then "first_footer"
      else nil
      end
    end

    def convert_header_footer_div(div, klass)
      obj = klass.new

      div.children.each do |child|
        next if child.text? && child.text.strip.empty?

        case child.name
        when "p"
          para = convert_header_footer_paragraph(child)
          obj.paragraphs << para
        when "table"
          table = convert_table(child)
          obj.tables << table
        else
          if child.element?
            para = convert_header_footer_paragraph(child)
            obj.paragraphs << para
          end
        end
      end

      obj
    end

    def convert_header_footer_paragraph(element)
      para = Uniword::Wordprocessingml::Paragraph.new

      # Set style from class attribute
      css_class = element["class"]
      if css_class
        style_id = resolve_style(css_class)
        if style_id
          para.properties = Uniword::Wordprocessingml::ParagraphProperties.new(
            style: style_id
          )
        end
      end

      # Parse inline style for paragraph properties
      parse_paragraph_style(para, element["style"]) if element["style"]

      # Convert inline content with field code support
      convert_header_footer_inline(para, element)

      para
    end

    # Convert inline content in header/footer with IE conditional comment field codes
    def convert_header_footer_inline(para, element)
      children = element.children.to_a
      i = 0

      while i < children.size
        child = children[i]

        if child.comment?
          # Check for IE conditional comment with field codes
          comment_text = child.text

          if comment_text.include?("mso-element:field-begin")
            # Extract field instruction text from the comment
            instr_text = extract_field_instruction(comment_text)

            # Create field-begin run
            run_begin = Uniword::Wordprocessingml::Run.new
            run_begin.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "begin")
            para.runs << run_begin

            # Create instrText run
            run_instr = Uniword::Wordprocessingml::Run.new
            run_instr.instr_text = Uniword::Wordprocessingml::InstrText.new(text: " #{instr_text} ")
            para.runs << run_instr

            # Create field-separate run
            run_sep = Uniword::Wordprocessingml::Run.new
            run_sep.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "separate")
            para.runs << run_sep

            # Next child should be the cached result (element)
            i += 1
            if i < children.size && children[i].element?
              convert_inline_content(para, children[i])
            end

            # Look for field-end comment
            i += 1
            while i < children.size
              if children[i].comment? && children[i].text.include?("mso-element:field-end")
                run_end = Uniword::Wordprocessingml::Run.new
                run_end.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "end")
                para.runs << run_end
                break
              elsif children[i].element?
                # Skip elements between field-separator and field-end
                convert_inline_content(para, children[i])
              end
              i += 1
            end
          elsif comment_text.include?("mso-element:field-end")
            # Orphan field-end (shouldn't happen normally)
            run_end = Uniword::Wordprocessingml::Run.new
            run_end.field_char = Uniword::Wordprocessingml::FieldChar.new(fldCharType: "end")
            para.runs << run_end
          else
            # Other IE conditional comment — check for nested mso-special-character
            if comment_text.include?("footnote-separator")
              # Footnote separator — skip
            end
          end
        elsif child.text?
          text = child.text
          unless text.strip.empty?
            para.runs << create_run(text)
          end
        elsif child.element?
          convert_element(para, child)
        end

        i += 1
      end
    end

    # Extract field instruction from an IE conditional comment containing field-begin.
    # E.g., " PAGE \* MERGEFORMAT " from the comment HTML
    def extract_field_instruction(comment_text)
      doc = Nokogiri::HTML.fragment(comment_text)

      field_begin = doc.css("span").find { |s| s["style"]&.include?("mso-element:field-begin") }
      return "" unless field_begin

      parts = []
      node = field_begin.next
      while node
        break if node.element? && node["style"]&.include?("mso-element:field-separator")
        parts << node.text
        node = node.next
      end

      parts.join.strip
    end

    # Build the flat list of header/footer parts for document.header_footer_parts.
    # Each entry: {r_id, target, rel_type, content_type, content, ref_type}
    def build_header_footer_parts(sections)
      parts = []
      header_idx = 0
      footer_idx = 0
      multi_section = @section_boundaries && @section_boundaries.size > 1

      sections.each_with_index do |section, sect_num|
        %w[even_header default_header first_header].each do |key|
          obj = section[key]
          next unless obj

          # Multi-section cover page: only even header
          if multi_section && sect_num == 0 && key != "even_header"
            next
          end

          header_idx += 1
          r_id = "rIdHdr#{header_idx}"
          ref_type = key.sub(/_header$/, "")

          parts << {
            r_id: r_id,
            target: "header#{header_idx}.xml",
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
            content_type: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            content: obj,
            ref_type: ref_type,
            section_index: sect_num
          }
        end

        %w[even_footer default_footer first_footer].each do |key|
          obj = section[key]
          next unless obj

          # Multi-section cover page: no footers
          if multi_section && sect_num == 0
            next
          end

          footer_idx += 1
          r_id = "rIdFtr#{footer_idx}"
          ref_type = key.sub(/_footer$/, "")

          parts << {
            r_id: r_id,
            target: "footer#{footer_idx}.xml",
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
            content_type: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            content: obj,
            ref_type: ref_type,
            section_index: sect_num
          }
        end
      end

      parts
    end

    # Apply section properties to the body.
    # Detects WordSection boundaries from @section_boundaries (set during convert_body).
    # Falls back to header_file sections if no HTML section markers found.
    def apply_sections(body, header_sections)
      @section_pr_map = {}

      has_html_sections = @section_boundaries && !@section_boundaries.empty?
      has_header_sections = header_sections && !header_sections.empty?
      return unless has_html_sections || has_header_sections

      # Parse page dimensions from stylesheet for each section
      section_page_props = parse_page_sections

      if !has_html_sections
        # No WordSection divs in HTML: single section from header_file
        sect_pr = build_section_properties(header_sections&.dig(0) || {}, 0, section_page_props[1])
        body.section_properties = sect_pr
        @section_pr_map[0] = sect_pr
        return
      end

      num_sections = @section_boundaries.size
      if num_sections == 1
        sect_pr = build_section_properties(header_sections&.dig(0) || {}, 0, section_page_props[1])
        body.section_properties = sect_pr
        @section_pr_map[0] = sect_pr
        return
      end

      # Use WordSection boundaries from convert_body to place sectPr
      all_paras = body.paragraphs
      total = all_paras.size
      return if total == 0

      @section_boundaries.each_with_index do |para_idx, sect_idx|
        header_section = header_sections&.dig(sect_idx) || {}
        page_props = section_page_props[sect_idx + 1]
        sect_pr = build_section_properties(header_section, sect_idx, page_props)

        if sect_idx < @section_boundaries.size - 1
          # Intermediate section: sectPr on the last paragraph of this section
          # para_idx is the index of the first paragraph of the NEXT section
          target_idx = [para_idx - 1, total - 1].min
          # Find the actual last paragraph (skip tables in @body_content)
          target_idx = find_last_paragraph_idx(target_idx)
          para = all_paras[target_idx]
          if para
            para.properties ||= Uniword::Wordprocessingml::ParagraphProperties.new
            para.properties.section_properties = sect_pr
            @section_pr_map[sect_idx] = sect_pr
          end
        else
          # Last section: sectPr on body
          body.section_properties = sect_pr
          @section_pr_map[sect_idx] = sect_pr
        end
      end
    end

    # Find the paragraph count (index into body.paragraphs) for the last paragraph
    # at or before the given @body_content index.
    def find_last_paragraph_idx(max_content_idx)
      para_count = 0
      @body_content.each_with_index do |item, idx|
        break if idx > max_content_idx
        para_count += 1 if item.first == :p
      end
      [para_count - 1, 0].max
    end

    def build_section_properties(header_section, section_idx, page_props)
      sect_pr = Uniword::Wordprocessingml::SectionProperties.new

      if page_props
        sect_pr.page_size = Uniword::Wordprocessingml::PageSize.new(
          width: page_props[:width], height: page_props[:height]
        )
        sect_pr.page_margins = Uniword::Wordprocessingml::PageMargins.new(**page_props[:margins])
      else
        # A4 defaults
        sect_pr.page_size = Uniword::Wordprocessingml::PageSize.new(
          width: 11906, height: 16838
        )
        sect_pr.page_margins = Uniword::Wordprocessingml::PageMargins.new(
          top: 1440, right: 1440, bottom: 1440, left: 1440,
          header: 720, footer: 720, gutter: 0
        )
      end

      # Page numbering: start at 1 for body section (section 3+)
      if section_idx >= 2
        sect_pr.page_numbering = Uniword::Wordprocessingml::PageNumbering.new(start: "1")
      end

      # TitlePg if any first-page header/footer exists
      if header_section["first_header"] || header_section["first_footer"]
        sect_pr.title_pg = Uniword::Wordprocessingml::TitlePg.new
      end

      sect_pr
    end

    # Parse @page CSS rules from the stylesheet for each WordSection.
    # Returns a hash: { 1 => { width:, height:, margins: {} }, ... }
    def parse_page_sections
      return {} unless @stylesheet

      result = {}
      section_num = nil
      found_size = nil
      found_margin = nil
      in_page_rule = false

      @stylesheet.lines.each do |line|
        if (m = line.match(/@page\s+WordSection(\d+)/))
          section_num = m[1].to_i
          found_size = nil
          found_margin = nil
          in_page_rule = true
        elsif in_page_rule && line.include?("size:")
          if m = line.match(/size:\s*(\S+)\s+(\S+)/)
            found_size = { width: units_to_twips(m[1]), height: units_to_twips(m[2]) }
          end
        elsif in_page_rule && line.include?("margin:")
          if m = line.match(/margin:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+)/)
            found_margin = {
              top: units_to_twips(m[1]),
              right: units_to_twips(m[2]),
              bottom: units_to_twips(m[3]),
              left: units_to_twips(m[4]),
              header: 720, footer: 720, gutter: 0
            }
          end
        elsif in_page_rule && line.include?("}")
          if section_num && (found_size || found_margin)
            result[section_num] = {
              width: found_size&.dig(:width) || 11906,
              height: found_size&.dig(:height) || 16838,
              margins: found_margin || {
                top: 1440, right: 1440, bottom: 1440, left: 1440,
                header: 720, footer: 720, gutter: 0
              }
            }
          end
          in_page_rule = false
          section_num = nil
        end
      end

      result
    end

    def units_to_twips(value)
      case value
      when /(\d+\.?\d*)pt/ then ($1.to_f * 20).to_i
      when /(\d+\.?\d*)cm/ then ($1.to_f * 567).to_i
      when /(\d+\.?\d*)in/ then ($1.to_f * 1440).to_i
      when /(\d+\.?\d*)mm/ then ($1.to_f * 56.7).to_i
      when /(\d+\.?\d*)px/ then ($1.to_f * 15).to_i
      else value.to_i
      end
    end

    # Wire header/footer references from header_footer_parts into section properties.
    # Uses @section_pr_map built by apply_sections.
    def wire_header_footer_refs(body, parts)
      parts.each do |part|
        sect_pr = @section_pr_map[part[:section_index]]
        next unless sect_pr

        case part[:rel_type]
        when /header/
          sect_pr.header_references << Uniword::Wordprocessingml::HeaderReference.new(
            type: part[:ref_type], r_id: part[:r_id]
          )
        when /footer/
          sect_pr.footer_references << Uniword::Wordprocessingml::FooterReference.new(
            type: part[:ref_type], r_id: part[:r_id]
          )
        end
      end
    end
  end
end
