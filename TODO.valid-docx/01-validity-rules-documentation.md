# Plan 01: DOCX Validity Rules Documentation

## Goal
Create `docs/docx-valid/` in the Uniword repo with a complete, authoritative catalog of every rule that makes a DOCX file valid in Microsoft Word. This is the reference that drives both the reconciler (pre-save) and the validation rules (post-save).

## What we learned from debugging

These are the rules we discovered through swap-testing and repair analysis:

### R1. mc:Ignorable namespace consistency
**Rule:** Every namespace prefix listed in `mc:Ignorable` MUST have a corresponding `xmlns:PREFIX` declaration on the same element or an ancestor.

**Where it matters:** Any part that uses `mc:Ignorable`:
- `word/document.xml`
- `word/styles.xml`
- `word/settings.xml`
- `word/numbering.xml`
- `word/fontTable.xml`
- `word/webSettings.xml`

**What Word does on violation:** Refuses to open ("unreadable content") or offers repair. During repair, Word strips the mc:Ignorable attribute or adds the missing namespace declarations.

**Reconciler enforcement:** `namespace_scope` with `declare: :always` for all prefixes that could appear in mc:Ignorable.

**Validation check:** Parse root element; extract mc:Ignorable prefixes; verify each has xmlns declaration in scope.

### R2. w15:docId GUID format
**Rule:** `<w15:docId w:val="..."/>` in settings.xml MUST be a GUID in curly-brace format: `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.

**What Word does on violation:** "Summary Info 1 is corrupted" error.

**Reconciler enforcement:** Reconciler generates `"{#{SecureRandom.uuid.upcase}}"`.

**Validation check:** Parse settings.xml; find w15:docId; verify val matches GUID regex.

### R3. Theme fmtScheme completeness
**Rule:** `<a:fmtScheme>` in theme.xml MUST have non-empty child lists:
- `<a:fillStyleLst>` — at least 2 fill elements
- `<a:lnStyleLst>` — at least 3 line elements
- `<a:effectStyleLst>` — at least 3 effectStyle elements
- `<a:bgFillStyleLst>` — at least 2 fill elements

**What Word does on violation:** "Summary Info 1 is corrupted". During repair, Word replaces the entire theme with its default.

**Reconciler enforcement:** `default_format_scheme` method on ThemeTransformation; office_theme.xml template.

**Validation check:** Parse theme1.xml; count children in each style list.

### R4. Numbering level element preservation
**Rule:** All child elements of `<w:lvl>` must round-trip: `<w:start>`, `<w:numFmt>`, `<w:pStyle>`, `<w:suff>`, `<w:lvlRestart>`, `<w:lvlText>`, `<w:lvlJc>`, `<w:pPr>`, `<w:tabs>`, `<w:ind>`, `<w:rPr>`, and `@w:tentative`.

**What Word does on violation:** Varies. Missing `<w:suff>` or `<w:lvlRestart>` may cause numbering to render differently. Missing `<w:pStyle>` causes Word to remove the reference on repair.

**Reconciler enforcement:** Level model class in level.rb with all attributes and mappings.

**Validation check:** Round-trip diff of numbering.xml; verify no content loss.

### R5. Numbering instance lvlOverride preservation
**Rule:** `<w:lvlOverride>` and `<w:startOverride>` inside `<w:num>` must be preserved.

**What Word does on violation:** Numbering restart behavior changes.

**Reconciler enforcement:** LevelOverride/StartOverride model classes; lvl_overrides attribute on NumberingInstance.

**Validation check:** Round-trip diff of numbering.xml.

### R6. Relationship cross-reference integrity
**Rule:** rId values in any part XML (e.g., `r:id="rId6"` in document.xml) must match a corresponding Relationship element in the associated .rels file (e.g., `word/_rels/document.xml.rels`). The rId must have the correct type and target.

**Where it matters:**
- `word/document.xml` ↔ `word/_rels/document.xml.rels` (styles, settings, fontTable, theme, numbering, headers, footers, images)
- `word/_rels/document.xml.rels` ↔ actual ZIP entries
- `_rels/.rels` ↔ top-level parts

**What Word does on violation:** Content loss, broken references, repair.

**Reconciler enforcement:** `reconcile_document_rels` and `reconcile_package_rels` rebuild standard relationship sets; non-standard rels preserved.

**Validation check:** Extract all rId references from each XML part; verify each exists in the correct .rels file; verify each target exists in the ZIP.

### R7. Content Types consistency
**Rule:** Every part in the ZIP MUST be covered by either a `<Default>` or `<Override>` in `[Content_Types].xml`. Content type values must match the expected OOXML content types.

**Reconciler enforcement:** `reconcile_content_types` rebuilds defaults (rels, xml) and overrides for all standard parts.

**Validation check:** Enumerate ZIP entries; verify each has a content type.

### R8. Required parts
**Rule:** A valid DOCX MUST contain:
- `[Content_Types].xml`
- `_rels/.rels`
- `word/document.xml`
- `word/_rels/document.xml.rels`

And SHOULD contain:
- `word/styles.xml`
- `word/settings.xml`
- `word/fontTable.xml`
- `word/webSettings.xml`

**Validation check:** Already covered by `OoxmlPartValidator`.

### R9. Footnote/endnote consistency
**Rule:** If `footnotePr` exists in settings.xml, then `word/footnotes.xml` must exist and contain separator (id="-1") and continuation (id="0") entries. Same pattern for endnotes.

**Reconciler enforcement:** `reconcile_footnotes` and `reconcile_endnotes`.

**Validation check:** Already covered by `FootnotesRule` (DOC-020, DOC-021, DOC-022).

### R10. Style definition completeness
**Rule:** Styles.xml must contain at minimum:
- A paragraph-style default (styleId="Normal")
- A character-style default (styleId="DefaultParagraphFont")
- A table-style default (styleId="TableNormal")
- A numbering-style default (styleId="NoList")

**Reconciler enforcement:** `ensure_default_styles` in reconciler.

**Validation check:** Already covered by `StyleReferencesRule` (DOC-002).

### R11. Section properties
**Rule:** Document body MUST have a `<w:sectPr>` element (at least on the last paragraph). Section properties must have `<w:pgSz>` and `<w:pgMar>`.

**Reconciler enforcement:** `reconcile_section_properties`.

**Validation check:** New rule needed.

### R12. Document rsids
**Rule:** Paragraphs should have `w:rsidR` and `w:rsidRDefault` attributes. Section properties should have `w:rsidR`. These are not strictly required for validity but Word adds them on repair.

**Reconciler enforcement:** `reconcile_document_body`.

**Validation check:** New rule (warning level).

### R13. Font table signature completeness
**Rule:** Each `<w:font>` entry's `<w:sig>` element must have valid usb0-usb3 and csb0-csb1 attributes. Empty `<w:sig/>` is tolerated but non-standard.

**What Word does on violation:** Unknown error or silent correction on repair.

**Reconciler enforcement:** Font metadata loaded from config/font_metadata.yml.

**Validation check:** New rule (warning level).

### R14. Core properties namespace declarations
**Rule:** docProps/core.xml must declare `xmlns:dcterms`, `xmlns:dc`, and `xmlns:xsi`. The `dcterms:created` and `dcterms:modified` elements must have `xsi:type="dcterms:W3CDTF"`.

**Reconciler enforcement:** `reconcile_core_properties` rebuilds from scratch to ensure namespace_scope.

**Validation check:** New rule.

### R15. Header/footer reference consistency
**Rule:** `<w:headerReference>` and `<w:footerReference>` in `<w:sectPr>` must reference existing header/footer XML parts via valid rIds. The parts must be in `word/` and have corresponding relationship entries.

**Reconciler enforcement:** `build_header_footer_parts` and `wire_header_footer_refs`.

**Validation check:** Already covered by `HeadersFootersRule` (DOC-030, DOC-031).

## Deliverable structure

```
docs/docx-valid/
  index.adoc              — overview, links to each rule category
  rules/
    namespace-consistency.adoc    — R1
    settings-values.adoc          — R2
    theme-completeness.adoc       — R3
    numbering-preservation.adoc   — R4, R5
    relationship-integrity.adoc   — R6
    content-types.adoc            — R7
    required-parts.adoc           — R8
    footnote-consistency.adoc     — R9
    style-completeness.adoc       — R10
    section-properties.adoc       — R11
    rsids.adoc                    — R12
    font-table.adoc               — R13
    core-properties.adoc          — R14
    header-footer-refs.adoc       — R15
  reconciler-mapping.adoc — maps each rule to its reconciler method
  validation-mapping.adoc — maps each rule to its validation rule class
```

Each rule doc follows this template:
```
== Rule ID: R{n}
**Summary:** one-line description
**Severity:** error / warning / info
**Applies to:** which XML parts

=== What Word expects
Detailed description of the valid state.

=== What happens on violation
Error message or behavior.

=== Reconciler enforcement
Which reconciler method ensures this, and how.

=== Validation rule
Which DOC-{code} checks this, or "New rule needed: DOC-{code}".

=== Example
Valid vs invalid XML snippet.
```
