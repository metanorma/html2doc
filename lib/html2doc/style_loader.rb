require "uniword"

class Html2Doc
  module StyleLoader
    # Default template for ISO-style DOCX output
    DEFAULT_TEMPLATE = File.expand_path("../../spec/fixtures/iso-damd-fdis-sample.docx", __dir__)

    TEMPLATE_CACHE = {}

    class << self
      # Load all DOCX template parts from a real .docx file.
      # Returns a Uniword::Docx::Package with styles, numbering, etc.
      def template_package(template_path = DEFAULT_TEMPLATE)
        TEMPLATE_CACHE[template_path] ||= load_template(template_path)
      end

      def iso_styles
        pkg = template_package
        pkg&.styles
      end

      def iso_numbering
        pkg = template_package
        pkg&.numbering
      end

      def iso_font_table
        pkg = template_package
        pkg&.font_table
      end

      def iso_settings
        pkg = template_package
        pkg&.settings
      end

      def iso_theme
        pkg = template_package
        pkg&.theme
      end

      # Map CSS class name to OOXML style ID, auto-derived from the template.
      def style_for_class(css_class)
        class_to_style[css_class]
      end

      def class_to_style
        @class_to_style ||= build_class_to_style_map
      end

      # Clear cached template data (useful when switching templates)
      def reset!
        TEMPLATE_CACHE.clear
        @class_to_style = nil
      end

      private

      def load_template(path)
        unless path && File.exist?(path)
          warn "html2doc: template not found: #{path}"
          return nil
        end
        Uniword::Docx::Package.from_file(path)
      end

      # Build a CSS class → styleId map from the template's styles.
      #
      # Strategy:
      # 1. Mso* classes map to built-in styles by English name
      #    (e.g., MsoNormal → style with name "Normal" → its styleId)
      # 2. Non-Mso classes match via direct styleId or name-based lookup
      #    (e.g., "Note" matches style name "note" → styleId "Note0")
      def build_class_to_style_map
        styles = iso_styles
        return {} unless styles

        # Build name → styleId lookup from template (case-insensitive)
        name_to_id = {}
        styles.styles.each do |style|
          next unless style.respond_to?(:name) && style.name
          name_str = style.name.respond_to?(:val) ? style.name.val : style.name.to_s
          name_to_id[name_str.downcase] = style.id
        end

        map = {}

        # Build set of known styleIds for quick lookup
        known_ids = Set.new(styles.styles.map(&:id))

        # Standard Mso class → English name → styleId from template
        mso_mappings = {
          "MsoNormal" => "normal",
          "MsoHeading1" => "heading 1",
          "MsoHeading2" => "heading 2",
          "MsoHeading3" => "heading 3",
          "MsoHeading4" => "heading 4",
          "MsoHeading5" => "heading 5",
          "MsoHeading6" => "heading 6",
          "MsoTOC1" => "toc 1",
          "MsoTOC2" => "toc 2",
          "MsoTOC3" => "toc 3",
          "MsoTOC4" => "toc 4",
          "MsoTOC5" => "toc 5",
          "MsoTOC6" => "toc 6",
          "MsoTOC7" => "toc 7",
          "MsoTOC8" => "toc 8",
          "MsoTOC9" => "toc 9",
          # HTML uses mixed case (MsoToc1) while MSO standard uses uppercase (MsoTOC1)
          "MsoToc1" => "toc 1",
          "MsoToc2" => "toc 2",
          "MsoToc3" => "toc 3",
          "MsoToc4" => "toc 4",
          "MsoToc5" => "toc 5",
          "MsoToc6" => "toc 6",
          "MsoToc7" => "toc 7",
          "MsoToc8" => "toc 8",
          "MsoToc9" => "toc 9",
          "MsoFootnoteText" => "footnote text",
          "MsoFootnoteReference" => "footnote reference",
          "MsoEndnoteText" => "endnote text",
          "MsoEndnoteReference" => "endnote reference",
          "MsoHeader" => "header",
          "MsoFooter" => "footer",
          "Hyperlink" => "hyperlink",
        }.freeze

        mso_mappings.each do |css_class, style_name|
          style_id = name_to_id[style_name]
          map[css_class] = style_id if style_id
        end

        # Hardcoded aliases: MsoTitle/MsoSubtitle → ISO document title
        if known_ids.include?("zzSTDTitle")
          map["MsoTitle"] = "zzSTDTitle"
          map["MsoSubtitle"] = "zzSTDTitle"
        end

        # MsoListParagraph variants → Normal (default paragraph)
        if name_to_id["normal"]
          %w[MsoListParagraphCxSpFirst MsoListParagraphCxSpMiddle MsoListParagraphCxSpLast].each do |cls|
            map[cls] = name_to_id["normal"]
          end
        end

        # Dynamic mapping: any styleId in the template is a valid CSS class
        # This covers all ISO-specific classes automatically
        known_ids.each do |style_id|
          next if style_id.nil?
          map[style_id] = style_id unless map.key?(style_id)
        end

        # Name-based mapping: CSS class (lowercased) → styleId via name
        # Handles case variations like "Note" → name "note" → "Note0"
        name_to_id.each do |style_name, style_id|
          next if style_id.nil?
          # Register both the lowercase name as-is and common camelCase variants
          map[style_name] = style_id unless map.key?(style_name)
          # Also register the class-style (capitalized first letter)
          if style_name.include?(" ")
            # Multi-word names like "heading 1" are handled by Mso mappings
            next
          end
          # "note" → "Note", "h2annex" → "h2Annex" style
          camel = style_name.sub(/^(.)/) { $1.upcase }
          camel = camel.sub(/_(.)/) { $1.upcase } # snake_case → camelCase
          map[camel] = style_id unless map.key?(camel)
        end

        # Additional common aliases from HTML → template styleIds
        # Use explicit styleIds to avoid name collision issues
        # (e.g., "note" name exists as both styleId "note" and "Note0")
        aliases = {
          "Annex" => "ANNEX",
          "Biblio" => "biblio",
          "Note" => "note",
          "FigureTitle" => "figuretitle",
          "TableTitle" => "tabletitle",
          "Formula" => "Formula",
          "h2Annex" => "h2annex",
          "h3Annex" => "h3annex",
          "h4Annex" => "h4annex",
          "h5Annex" => "h5annex",
          "Sourcecode" => "sourcecode",
          "Admonition" => "admonition",
          "Example" => "example",
          "TableFootnoteRef" => "tablefootnoteref",
          "Section3" => "section3",
          "msotoctextspan1" => "msotoctextspan1",
          "MsoTocTextSpan" => "msotoctextspan1",
          "zzmovetofollowing" => "zzmovetofollowing",
          "zzMoveToFollowing" => "zzmovetofollowing",
          "zzSTDTitle1" => "zzSTDTitle",
        }.compact
        map.merge!(aliases)

        map.freeze
      end
    end
  end
end
