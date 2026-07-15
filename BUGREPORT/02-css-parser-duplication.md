# Bug 02: CSS Parsing Duplication Between DocxConverter and CssParser

**Status**: FIXED
**Severity**: Medium (DRY violation — could cause divergent parsing behavior)
**Location**: `lib/html2doc/docx_converter.rb`, `lib/html2doc/css_parser.rb`

## Symptom

No runtime error, but `DocxConverter` had its own `parse_span_style`, `parse_color`, `parse_font_size`, `parse_paragraph_style`, and `css_length_to_twips` methods that duplicated the same logic in `CssParser`. If one was updated without the other, parsing behavior would diverge silently.

## Root Cause

`CssParser` was extracted as a focused parsing module but the original inline methods in `DocxConverter` were never removed. Both implementations parsed the same CSS properties with identical logic.

Additionally, field-style helper methods (`field_begin_style?`, `field_sep_style?`, `field_end_style?`, `tab_count_style?`, `field_wrapper_span?`) were duplicated between `docx_converter.rb` and `toc_builder.rb`.

## Fix

1. Removed `parse_span_style` from `DocxConverter` — now delegates to `CssParser.parse_run_style`
2. Rewrote `parse_paragraph_style` to use `CssParser.parse_paragraph_style` + OOXML construction
3. Removed `parse_color`, `parse_font_size`, `css_length_to_twips` — all delegate to `CssParser`
4. Added `field_wrapper_span?` to `CssParser` as the canonical implementation
5. Both `DocxConverter` and `TocBuilder` delegate field-style checks through `CssParser`

**Line reduction**: `docx_converter.rb` reduced from 1299 → 1161 lines (~140 lines of duplication removed).

## Lesson

`CssParser` must remain the single source of truth for CSS parsing. Any new CSS parsing should go there, not inline in the converter or mixins.
