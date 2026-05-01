# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

html2doc is a Ruby gem that converts HTML documents into Microsoft Word format. Part of the Metanorma ecosystem. Two output formats are supported:

- **MHT** (`:mht`, default) — MIME MHTML with `.doc` extension (legacy)
- **DOCX** (`:docx`) — True OOXML `.docx` via the Uniword library

Requires Ruby >= 2.7.0.

## Commands

```bash
bundle exec rake              # run tests (default task)
bundle exec rspec             # run tests directly
bundle exec rspec spec/html2doc_spec.rb -e "test name"  # run single test by description
bundle exec rspec spec/docx_output_spec.rb    # DOCX output tests
bundle exec rspec spec/docx_converter_spec.rb # converter unit tests
bundle exec rspec spec/style_loader_spec.rb   # style loader tests
bundle exec rubocop           # lint
```

Tests use `equivalent-xml` for XML comparison and `rspec-match_fuzzy` for fuzzy string matching.

## Architecture

The `Html2Doc` class is split across multiple files using Ruby's open-class pattern. Each file adds methods to the same class, focusing on one concern:

- **base.rb** — Constructor, `process()` entry point, `cleanup` pipeline, image/stylesheet/bookmark handling
- **docx_converter.rb** — DOCX output: converts Nokogiri HTML to Uniword model objects (Paragraph, Run, Table, etc.)
- **style_loader.rb** — ISO template styles/numbering/fonts/theme/settings loading from `data/*.xml` via Uniword
- **math.rb** — MathML-to-OOXML via Plurimath; post-processing (unitalic, accents, plane1 fonts, centering)
- **notes.rb** — Footnote detection and Word footnote element creation
- **lists.rb** — HTML `<ul>`/`<ol>` to Word `MsoListParagraphCxSp*` styles
- **mime.rb** — MIME `.doc` packaging with base64-encoded images, image cleanup/resizing, filelist.xml generation
- **xml.rb** — XHTML conversion, MS Word XML fixups, namespace handling, doctype management

### MHT Processing pipeline (`base.rb#cleanup`)

`process_html` → `cleanup` (locate_landscape → namespace → image_cleanup → mathml_to_ooml → lists → footnotes → bookmarks → msonormal) → `define_head` → `msword_fix` → `generate_filelist` → `mime_package`

### DOCX Processing pipeline (`base.rb#process_docx`)

`to_xhtml` → `cleanup` (skip footnotes, skip images) → `DocxConverter.convert(docxml)` → `save_to_file`

### DocxConverter pipeline (`docx_converter.rb`)

1. Extract footnotes → Convert body (paragraphs, headings, tables, divs, lists, math)
2. Load ISO styles/numbering → Parse header.html for headers/footers/sections
3. Apply section properties → Wire header/footer references → Assemble DocxPackage

### Key patterns

- **Dual output**: `output_format: :mht` or `:docx`. MHT path unchanged; DOCX path uses Uniword.
- **Uniword wrapper types**: StyleReference, Alignment, NumberingId, NumberingLevel must use `Klass.new(value: "string")` for setters (constructors auto-wrap, setters do not).
- **ISO styles**: Finnish locale IDs in template (Normaali=Normal, Otsikko1=Heading1, Alaviitteenteksti=FootnoteText). Mapped via `StyleLoader.class_to_style`.
- **Header/Footer**: Parsed from MS Word HTML with IE conditional comments. PAGE field codes converted to OOXML `fldChar`/`instrText`. Multi-section support via `header_footer_parts` array on DocumentRoot.
- **Image handling (MHT)**: Local images renamed to UUIDs, MIME-embedded with `cid:` references. SVG not supported.
- **Image handling (DOCX)**: Via `Uniword::Builder::ImageBuilder`, binary stored in `image_parts`, auto-packaged into `word/media/`.
- **Math**: html2doc converts MathML to OMML via Plurimath (6 pre-existing test failures). DocxConverter detects `<m:oMath>`/`<m:oMathPara>` and passes through to Uniword.
- **CSS selectors**: DOCX path uses CSS selectors (not XPath) because XHTML namespace isn't registered after `to_xhtml`.

### CLI

`bin/html2doc` accepts `--stylesheet`, `--header`, and `--format` (mht/docx) options.

## Key dependencies

- **uniword** — OOXML document model and DOCX packaging
- **nokogiri** — XML/HTML parsing and manipulation
- **plurimath** — MathML/AsciiMath conversion to OOXML
- **vectory** — Image resizing calculations for Word page dimensions
- **marcel** — MIME type detection for images
- **plane1converter** — Plane 1 Unicode font conversion for math
- **htmlentities** — HTML entity encoding/decoding
- **uuidtools** — UUID generation for image renaming and MIME boundaries
