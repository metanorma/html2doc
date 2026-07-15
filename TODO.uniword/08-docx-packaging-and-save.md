# Phase 8: DOCX Packaging, Save, and CLI Integration

## Goal
Wire everything together: assemble the complete `DocxPackage` from all converted parts, save to disk, and update the CLI and `Html2Doc.process` API to support DOCX output.

## Background
A DOCX file is a ZIP archive containing XML parts. Uniword's `DocxPackage` manages all parts:
```
[Content_Types].xml      -- content type registry
_rels/.rels              -- package relationships
word/document.xml        -- main document body
word/styles.xml          -- style definitions
word/numbering.xml       -- list/numbering definitions
word/settings.xml        -- document settings
word/fontTable.xml       -- font declarations
word/webSettings.xml     -- web settings
word/theme/theme1.xml    -- theme (colors, fonts)
word/footnotes.xml       -- footnotes
word/endnotes.xml        -- endnotes
word/header*.xml         -- headers
word/footer*.xml         -- footers
word/media/image*        -- embedded images
word/_rels/document.xml.rels  -- document-level relationships
word/_rels/footnotes.xml.rels -- footnote relationships
docProps/core.xml        -- Dublin Core metadata
docProps/app.xml         -- application metadata
```

## Tasks

- [ ] Implement `DocxConverter#assemble_package(document, styles, numbering, options)`:
  ```ruby
  def assemble_package(body, styles, numbering, options)
    document = Uniword::Wordprocessingml::DocumentRoot.new
    document.body = body

    package = Uniword::Ooxml::DocxPackage.new
    package.document = document
    package.styles = styles                    # from StyleLoader
    package.numbering = numbering              # from NumberingConverter
    package.content_types = build_content_types
    package.package_rels = build_package_rels
    package.settings = build_settings(options)
    package.font_table = build_font_table
    package.web_settings = build_web_settings
    package.theme = load_iso_theme
    package.footnotes = @footnotes             # from Phase 5
    package.endnotes = @endnotes
    # headers/footers injected by DocxPackage during save
    # images registered via document.image_parts
    package
  end
  ```

- [ ] Implement `build_content_types`:
  - Default types for `.rels` and `.xml`
  - Override types for all document parts (document, styles, settings, etc.)
  - Override types for headers, footers, footnotes, endnotes
  - Default types for image formats (png, jpeg, gif) based on embedded images

- [ ] Implement `build_package_rels`:
  - rId1 → `word/document.xml` (officeDocument)
  - rId2 → `docProps/core.xml` (core-properties)
  - rId3 → `docProps/app.xml` (extended-properties)

- [ ] Implement `build_settings`:
  - Use ISO fixture settings as template: mirror margins, even/odd headers, default tab stop
  - Include footnote/endnote separator configuration
  - Include compatibility settings (`w:compat`)
  - Include language settings from document

- [ ] Implement `build_font_table`:
  - Use ISO fixture font table: Calibri, Cambria, Times New Roman, etc.
  - Add any additional fonts referenced in the document content

- [ ] Implement `build_web_settings`:
  - Minimal: `<w:webSettings><w:divs/><w:optimizeForBrowser/></w:webSettings>`

- [ ] Update `Html2Doc#process` for DOCX output:
  ```ruby
  def process(result)
    result = process_html(result, output_format: @output_format)

    if @output_format == :docx
      converter = DocxConverter.new(@options)
      package = converter.convert(result)
      package.save("#{@filename}.docx")
    else
      # existing MHT path
      generate_filelist(@filename, @dir1)
      File.open("#{@filename}.htm", "w:UTF-8") { |f| f.write(result) }
      mime_package result, @filename, @dir1
      rm_temp_files(@filename, @dir, @dir1) unless @debug
    end
  end
  ```

- [ ] Update CLI (`bin/html2doc`):
  - Add `--format` option: `mht` (default) or `docx`
  - Pass format to `Html2Doc.new`

- [ ] Add metadata support:
  - `title` → `docProps/core.xml` dc:title
  - `author` → `docProps/core.xml` dc:creator
  - `created` → `docProps/core.xml` dcterms:created
  - Accept via constructor options: `Html2Doc.new(title: "...", author: "...")`

- [ ] Error handling:
  - Graceful fallback if Uniword not available (gem not installed)
  - Clear error messages for unsupported features in DOCX mode

## Key Files to Create/Modify
- `lib/html2doc/docx_converter.rb` — package assembly
- `lib/html2doc/base.rb` — process method branching
- `bin/html2doc` — CLI format option

## Success Criteria
- `Html2Doc.new(filename: "test", output_format: :docx).process(html)` produces a valid `.docx`
- DOCX opens correctly in Microsoft Word, LibreOffice, and Google Docs
- All parts present: document, styles, numbering, settings, font table, theme
- File size reasonable (not bloated with unnecessary data)
- CLI `--format docx` flag works
