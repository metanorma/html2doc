# Phase 10: Backward Compatibility, Integration with Metanorma, and Documentation

## Goal
Ensure the DOCX output path integrates cleanly with the upstream Metanorma ecosystem (isodoc → html2doc → DOCX), maintain full backward compatibility with the existing MHT path, and document the new capability.

## Background
html2doc is called by isodoc/metanorma as a post-processor:
```ruby
# In isodoc
Html2Doc.new(filename: filename, imagedir: imagedir, stylesheet: stylesheet,
             header_file: header_filename, dir: dir,
             asciimathdelims: asciimathdelims, liststyles: liststyles)
  .process(result)
```

The Metanorma ecosystem currently expects `.doc` (MHT) output. Adding `.docx` output must not break any existing caller.

## Tasks

- [ ] Verify backward compatibility:
  - Default `output_format` is `nil` or `:mht` — no breaking change
  - All existing `Html2Doc.new` options work identically
  - All existing tests pass unchanged
  - `bin/html2doc` without `--format` flag produces `.doc` as before

- [ ] Integrate with Metanorma (isodoc):
  - isodoc will need to pass `output_format: :docx` when DOCX output is desired
  - Coordinate with isodog for the API change (it's additive, not breaking)
  - The `filename` parameter semantics: currently produces `filename.doc`, with `:docx` it produces `filename.docx`

- [ ] Handle stylesheet differences:
  - MHT path: CSS stylesheet injected into HTML `<head>` with `@page` rules, `@list` rules, `mso-*` properties
  - DOCX path: Styles are OOXML XML in `word/styles.xml`, numbering in `word/numbering.xml`
  - When `output_format: :docx`:
    - Still parse the stylesheet for `@page` rules (needed for page dimensions)
    - Still parse `@list` rules (needed for list style mapping)
    - But do NOT inject CSS into the document — use OOXML styles instead
  - Accept optional `styles_file` parameter: path to a `.docx` or `.xml` to load styles from

- [ ] Handle the `header_file` difference:
  - MHT: `header.html` is a separate MIME part, referenced via `@page` CSS URLs
  - DOCX: Headers/footers are separate XML parts (`word/header*.xml`)
  - Parse `header.html` into OOXML header/footer objects
  - If no header_file provided: use ISO default headers/footers

- [ ] Update CLAUDE.md:
  - Add `output_format: :docx` to API documentation
  - Add DOCX-specific notes to Architecture section
  - Update processing pipeline description

- [ ] Update README.adoc:
  - Document the new DOCX output option
  - Add usage examples for DOCX generation
  - Note the difference between `.doc` (MHT) and `.docx` (OOXML) output

- [ ] Add migration guide for Metanorma integrators:
  - What changes: `output_format: :docx` parameter
  - What's different: real OOXML styles vs CSS, separate footnote part, etc.
  - Benefits: native Word format, better compatibility, smaller file size

- [ ] Consider deprecation timeline:
  - MHT output will remain the default for now
  - DOCX output opt-in via `output_format: :docx`
  - Future: flip default to `:docx` once validated in production

## Key Files to Create/Modify
- `lib/html2doc/base.rb` — API compatibility
- `CLAUDE.md` — architecture update
- `README.adoc` — user-facing documentation

## Success Criteria
- Zero breaking changes for existing users
- DOCX output works when `output_format: :docx` is specified
- Metanorma integration can be done with a single parameter addition
- Documentation clear about MHT vs DOCX tradeoffs
