# Bug 06: Dead Code Shadowed by Mixin Includes

**Status**: FIXED
**Severity**: Low (dead code, no runtime impact — but confused readers and inflated PR)
**Location**: `lib/html2doc/docx_converter.rb`

## Symptom

`DocxConverter` had ~870 lines of methods defined directly in the class body that were silently overridden by mixin `include` statements. Ruby's `include` makes the mixin's methods override class-defined methods, so the inline versions were dead code that could never be called.

Affected methods included: `toc_paragraph?`, `footnote_link?`, section management methods, TOC building methods, footnote extraction methods, and image handling methods.

## Root Cause

When `SectionManager`, `TocBuilder`, `FootnoteManager`, and `ImageHandler` were extracted as mixins, the original inline implementations were left in `docx_converter.rb` instead of being removed. Ruby's method resolution order made the mixin versions win silently.

## Fix

Stripped all shadowed method bodies from `docx_converter.rb`. Retained only:
- Constructor (`initialize`)
- Public API (`convert`, `save_to_file`)
- Private helper methods used by the class itself AND mixins (`new_paragraph`, `set_paragraph_style`, etc.)
- Inline content conversion (`convert_paragraph`, `convert_heading`, `convert_element`, etc.)
- Table conversion
- Math passthrough
- Package assembly

**Line reduction**: 2159 → 1161 lines (46% reduction from original).

## Lesson

After extracting a mixin, always remove the original methods from the class body. Leaving them creates confusion about which version is actually called.
