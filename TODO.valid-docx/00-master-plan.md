# Master Plan: Valid DOCX Generation and Validation

## Goal
Three deliverables:
1. **DOCX validity rules** — fully documented, model-driven, enforced by reconciler + validation
2. **Round-trip test suite** — verifies Uniword can load → save any DOCX without Word errors
3. **HTML-to-DOCX output verification** — html2doc's generated DOCX matches spec/fixtures baselines

## Status

| Plan | Title | Status |
|------|-------|--------|
| 01 | Validity rules documentation | Pending |
| 02 | OOP validation framework in Uniword | Pending |
| 03 | Reconciler hardening | Pending |
| 04 | Round-trip test suite | Pending |
| 05 | HTML-to-DOCX output verification | Pending |

## Dependencies
- Plans 01 and 02 can be done in parallel
- Plan 03 depends on 01 (rules inform reconciler fixes)
- Plan 04 depends on 02 and 03
- Plan 05 depends on 04

## Key files

### Uniword (validation + reconciler)
- `lib/uniword/validation/` — existing validation system (7-layer pipeline, semantic rules)
- `lib/uniword/validation/rules/` — 12 semantic rule classes (DOC-001..DOC-091)
- `lib/uniword/validation/validators/` — 7 layer validators (ZIP, OPC, XML, relationships, etc.)
- `lib/uniword/docx/reconciler.rb` — reconciler that ensures DOCX invariants before save
- `lib/uniword/docx/package.rb` — package assembly and save

### html2doc (HTML-to-DOCX pipeline)
- `lib/html2doc/docx_converter.rb` — converts Nokogiri HTML → Uniword model objects
- `lib/html2doc/style_loader.rb` — loads ISO template styles/numbering/fonts
- `spec/docx_output_spec.rb` — existing DOCX output tests
- `spec/docx_converter_spec.rb` — converter unit tests
- `data/iso_*.xml` — ISO template data files
- `spec/fixtures/iso-damd-fdis-sample.docx` — reference DOCX for comparison

### Documentation
- `docs/docx-valid/` — NEW: validity rules reference (to be created in Uniword repo)
- `docs/_verification/` — existing verification docs (semantic-rules, three-layer-pipeline, etc.)

## Terminology
- **Reconciler** — pre-serialization invariants (ensures model is valid before writing XML)
- **Validator** — post-serialization checks (verifies existing DOCX is valid)
- **mc:Ignorable** — Markup Compatibility attribute listing namespace prefixes that consumers may ignore; all listed prefixes MUST have xmlns declarations in scope
- **Round-trip** — load DOCX → Uniword model → save DOCX; the output must open in Word without errors
