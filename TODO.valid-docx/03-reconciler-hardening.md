# Plan 03: Reconciler Hardening

## Goal
Ensure the reconciler enforces every validity rule from Plan 01 before serialization. The reconciler should guarantee that any model state produces a valid DOCX.

## Current reconciler methods

| Method | Group | What it ensures |
|--------|-------|----------------|
| `reconcile_section_properties` | 1 (always) | sectPr with pgSz/pgMar |
| `reconcile_footnotes` | 1 (always) | footnotes.xml ↔ footnotePr consistency |
| `reconcile_endnotes` | 1 (always) | endnotes.xml ↔ endnotePr consistency |
| `reconcile_theme` | 2 (profile) | Theme with complete fmtScheme |
| `reconcile_settings` | 2 (profile) | Settings with all defaults, GUID docId, mc:Ignorable |
| `reconcile_font_table` | 2 (profile) | Font table with profile fonts, mc:Ignorable |
| `reconcile_styles` | 2 (profile) | Styles with docDefaults, latentStyles, defaults, mc:Ignorable |
| `reconcile_numbering` | 2 (profile) | Numbering mc:Ignorable |
| `reconcile_web_settings` | 2 (profile) | WebSettings mc:Ignorable |
| `reconcile_app_properties` | 2 (profile) | App properties with statistics |
| `reconcile_core_properties` | 2 (profile) | Core properties rebuilt for namespace_scope |
| `reconcile_document_body` | 2 (profile) | mc:Ignorable, rsids, paraId/textId |
| `reconcile_content_types` | 3 (always) | Content types for all standard parts |
| `reconcile_package_rels` | 3 (always) | Package-level relationships |
| `reconcile_document_rels` | 3 (always) | Document-level relationships (rId1-rId6) |
| `clear_stored_namespace_plans` | 3 (always) | Force namespace_scope declarations |

## Missing reconciler guarantees

### M1. mc:Ignorable prefix filtering
**Problem:** If a source document has `mc:Ignorable="w14 w15 w16se..."` but the model only maps some of these prefixes, the output may reference undeclared prefixes.

**Current fix:** `namespace_scope` with `declare: :always` for all known prefixes.
**Remaining risk:** New prefix introduced by Word that we don't know about.

**Solution:** Add a reconciler step that strips from mc:Ignorable any prefix not declared in the output XML.

```ruby
def reconcile_mc_ignorable
  # For each part that has mc_ignorable, verify all prefixes are declared
  [package.settings, package.font_table, package.styles,
   package.numbering, package.web_settings, package.document].compact.each do |part|
    next unless part.respond_to?(:mc_ignorable) && part.mc_ignorable
    # namespace_scope handles declaration; mc:Ignorable value is preserved from source
  end
end
```

### M2. Header/footer relationship reconciliation
**Problem:** Headers/footers added via `header_footer_parts` must have:
1. Relationship entries in `word/_rels/document.xml.rels`
2. Content type entries in `[Content_Types].xml`
3. Actual XML part content in the ZIP

**Current state:** `reconcile_document_rels` only handles standard parts (styles, settings, etc.). Header/footer rels are added by `Package#to_zip_content` directly.

**Solution:** Ensure reconciler accounts for header/footer parts in document_rels and content_types.

### M3. Image parts reconciliation
**Problem:** Images registered via `Builder::ImageBuilder` must have:
1. Relationship entries in document.xml.rels
2. Content type entries
3. Default content types for image extensions (png, jpeg, gif, etc.)
4. Binary data in word/media/

**Current state:** Images are handled by `inject_image_parts` in Package.

**Solution:** Verify reconciler's content_types includes image defaults; verify image rels are in document_rels.

### M4. Numbering reconciliation enhancement
**Problem:** The current `reconcile_numbering` only sets mc:Ignorable. It should also:
- Ensure numbering instance abstract_num_id references existing definitions
- Validate that lvl_overrides have correct ilvl values (0-8)

**Solution:** Extend `reconcile_numbering` with model-level consistency checks.

### M5. Audit trail
**Problem:** No visibility into what the reconciler changed.

**Solution:** Add `applied_fixes` array and `record_fix` helper (see Plan 02, Gap 3).

## Implementation steps

1. Add `reconcile_mc_ignorable` to Group 3 (always)
2. Extend `reconcile_document_rels` to account for headers/footers and images
3. Extend `reconcile_content_types` to add image content type defaults when images exist
4. Extend `reconcile_numbering` with consistency checks
5. Add audit trail (`applied_fixes`, `record_fix`)
6. Add `reconcile_numbering` to clear_stored_namespace_plans (already done)
7. Run validation after reconcile to verify (optional, from Plan 02)

## Testing

- After each change, round-trip rice.docx and verify it opens
- Round-trip other DOCX files (see Plan 04)
- Verify reconciler applied_fixes array is populated correctly
