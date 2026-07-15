# Phase 1: Architecture and Dependency Setup

## Goal
Establish the Uniword dependency, create the integration module skeleton, and define the architectural boundary between html2doc's HTML processing and Uniword's DOCX generation.

## Background
html2doc currently produces MHT (.doc) via `mime.rb`. We will add a parallel DOCX output path using Uniword's `DocxPackage`. The key architectural decision is **where** the handoff occurs:

- html2doc's `cleanup` pipeline (math, footnotes, lists, bookmarks) operates on Nokogiri HTML/XML documents
- Uniword's `DocxPackage` needs OOXML model objects (`Paragraph`, `Run`, `Table`, etc.)
- The bridge is a new `Html2Doc::DocxConverter` that takes the cleaned Nokogiri document and produces Uniword model objects

## Architecture Diagram

```
HTML input
  → process_html (existing)
    → cleanup pipeline (existing: math, footnotes, lists, bookmarks)
      → [BRANCH POINT]
        → existing path: msword_fix → MIME packaging → .doc
        → NEW path: DocxConverter → Uniword::DocxPackage → .docx
```

## Tasks

- [ ] Add `uniword` gem dependency to `html2doc.gemspec`
  ```
  spec.add_dependency "uniword", "~> 0.1.0"
  ```
  Or use a path dependency for development:
  ```
  spec.add_dependency "uniword", path: "../uniword"
  ```

- [ ] Create `lib/html2doc/docx_converter.rb` — the bridge module
  - Single class `Html2Doc::DocxConverter` with method `convert(docxml, options)`
  - Takes the Nokogiri document after `cleanup` pipeline
  - Returns a `Uniword::Ooxml::DocxPackage` ready for saving

- [ ] Add `output_format` parameter to `Html2Doc.new` and `Html2Doc#process`
  - `:mht` (default, backward-compatible) — existing MIME path
  - `:docx` — new Uniword path
  - When `:docx`, skip `msword_fix` and `mime_package`, use `DocxConverter` instead

- [ ] Verify existing tests still pass with `:mht` (default) output

- [ ] Update `Gemfile` to point to local Uniword for development

## Key Files to Create/Modify
- `html2doc.gemspec` — add uniword dependency
- `Gemfile` — add uniword path for development
- `lib/html2doc/docx_converter.rb` — new bridge module
- `lib/html2doc/base.rb` — add output_format branching in `process` and `process_html`
- `lib/html2doc.rb` — require new file

## Success Criteria
- `bundle install` resolves with Uniword
- Existing `:mht` path unchanged and all tests pass
- `DocxConverter` class exists with stub implementation
- `output_format: :docx` option accepted without errors (even if DOCX output is minimal)
