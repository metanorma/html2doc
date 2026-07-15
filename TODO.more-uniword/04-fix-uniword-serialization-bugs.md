# 04: Fix Uniword Serialization Bugs

## Summary

Fix two known Uniword serialization bugs that cause round-trip failures, preventing html2doc from producing fully valid DOCX output.

## Bug 1: styles.docx drops comments.xml

**Symptom**: When loading `styles.docx` and saving, the output is missing `word/comments.xml`.

**Impact**: Any document with comments loses them on round-trip.

**Investigation needed**:
- Check `DocxPackage#save` / serialization path for comments handling
- Verify `DocumentRoot` model includes `comments` attribute
- Check if `comments.xml` is in the package's part manifest

**Location**: `lib/uniword/docx/package_serialization.rb` or `lib/uniword/docx/reconciler.rb`

**Steps**:
1. Load `styles.docx` and inspect what parts are present:
   ```ruby
   pkg = Uniword::DocxPackage.from_file("styles.docx")
   puts pkg.document.comments&.size
   ```
2. Save and diff the ZIP contents:
   ```bash
   unzip -l styles.docx > before.txt
   # round-trip
   unzip -l styles-rt.docx > after.txt
   diff before.txt after.txt
   ```
3. Find where comments are dropped in the serialization pipeline
4. Fix and add regression test

## Bug 2: rice_roundtrip.docx drops header/footer parts

**Symptom**: When loading `rice_roundtrip.docx` and saving, 7 header/footer parts are missing from the output.

**Impact**: Multi-section documents lose their headers/footers on round-trip.

**Investigation needed**:
- Check how headers/footers are stored in the model
- Verify `DocumentRoot#headers` and `DocumentRoot#footers` are populated
- Check section property references to header/footer relationship IDs
- Verify `HeaderFooterPart` serialization

**Location**: `lib/uniword/docx/package_serialization.rb` or header/footer model classes

**Steps**:
1. Load `rice_roundtrip.docx` and inspect:
   ```ruby
   pkg = Uniword::DocxPackage.from_file("rice_roundtrip.docx")
   puts "Headers: #{pkg.document.headers&.size}"
   puts "Footers: #{pkg.document.footers&.size}"
   ```
2. Check section properties for headerReference/footerReference:
   ```ruby
   pkg.document.body.section_properties.each do |sp|
     puts sp.header_references&.map(&:type)
   end
   ```
3. Save and diff — identify which parts are missing
4. Fix the serialization to include all header/footer parts
5. Add regression test

## Acceptance Criteria

- [ ] styles.docx round-trip preserves comments.xml
- [ ] rice_roundtrip.docx round-trip preserves all header/footer parts
- [ ] Uniword repair_spec.rb tests pass for both files
- [ ] Regression tests added to Uniword's spec suite
