# frozen_string_literal: true

# Extracts styles, numbering, font table, settings, and theme from
# the ISO 690:2021 DOCX fixture using Uniword, saving each part as a
# reusable definition file under data/.
#
# Primary format: XML (full fidelity, loaded via StylesConfiguration.from_xml)
# Secondary format: StyleSet YAML (human-readable, may lose some properties)
#
# Usage:
#   bundle exec rake iso:extract_styles

require "uniword"

class Html2Doc
  module IsoStyleExtractor
    ISO_DOCX_PATH = File.expand_path(
      "../../../uniword/spec/fixtures/uniword-private/fixtures/iso/ISO_690_2021-Word_document(en).docx",
      __dir__
    )
    OUTPUT_DIR = File.expand_path("../../data", __dir__)

    class << self
      def extract_all
        unless File.exist?(ISO_DOCX_PATH)
          abort "ISO fixture not found: #{ISO_DOCX_PATH}"
        end

        FileUtils.mkdir_p(OUTPUT_DIR)
        pkg = Uniword::Ooxml::DocxPackage.from_file(ISO_DOCX_PATH)

        extract_styles_xml(pkg)
        extract_numbering(pkg)
        extract_font_table(pkg)
        extract_settings(pkg)
        extract_theme(pkg)
        extract_styles_yaml(pkg)

        puts "Extracted all ISO definitions to #{OUTPUT_DIR}/"
      end

      # Primary: full-fidelity XML serialization
      def extract_styles_xml(pkg)
        return unless pkg.styles

        path = File.join(OUTPUT_DIR, "iso_styles.xml")
        xml = pkg.styles.to_xml
        File.write(path, xml)
        count = pkg.styles.styles.size
        puts "  Styles (XML): #{count} styles -> #{path}"
      end

      # Secondary: human-readable StyleSet YAML (may lose some detail)
      def extract_styles_yaml(pkg)
        return unless pkg.styles

        styles = pkg.styles.styles
        styleset = Uniword::StyleSet.new(
          name: "ISO 690:2021",
          source_file: "ISO_690_2021-Word_document(en).docx",
          styles: styles,
        )

        path = File.join(OUTPUT_DIR, "iso_styles.yml")
        File.write(path, styleset.to_yaml)
        puts "  Styles (YAML): #{styles.size} styles -> #{path} (human-readable reference)"
      end

      def extract_numbering(pkg)
        return unless pkg.numbering

        path = File.join(OUTPUT_DIR, "iso_numbering.xml")
        File.write(path, pkg.numbering.to_xml)
        puts "  Numbering -> #{path}"
      end

      def extract_font_table(pkg)
        return unless pkg.font_table

        path = File.join(OUTPUT_DIR, "iso_font_table.xml")
        File.write(path, pkg.font_table.to_xml)
        puts "  Font table -> #{path}"
      end

      def extract_settings(pkg)
        return unless pkg.settings

        path = File.join(OUTPUT_DIR, "iso_settings.xml")
        File.write(path, pkg.settings.to_xml)
        puts "  Settings -> #{path}"
      end

      def extract_theme(pkg)
        return unless pkg.theme

        path = File.join(OUTPUT_DIR, "iso_theme.xml")
        File.write(path, pkg.theme.to_xml)
        puts "  Theme -> #{path}"
      end
    end
  end
end
