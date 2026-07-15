# 2. Math (AsciiMath stem) not converted to OMML in DOCX output

## Problem

The HTML `rice.html` contains 28 `<span class="stem">` elements with AsciiMath expressions (backtick-delimited). The generated DOCX contains **0 `<m:oMath>` elements**. All math appears as plain text (e.g. `` `r = 1 %` `` renders as literal text).

## Root cause

The DOCX pipeline in `process_docx` (base.rb:30-43) runs:

```ruby
docxml = to_xhtml(result)
cleanup(docxml, skip_footnotes: true, skip_images: true)  # includes mathml_to_ooml
converter = DocxConverter.new(...)
package = converter.convert(docxml)
```

The `cleanup` method calls `mathml_to_ooml(docxml)` which converts `<math>` elements to `<m:oMath>` OMML. DocxConverter then passes `<m:oMath>` through to Uniword.

**But the stem spans are AsciiMath, not MathML.** The conversion chain needs to be:

1. `<span class="stem">` + `asciimathdelims` → `<math>` (AsciiMath-to-MathML)
2. `<math>` → `<m:oMath>` (MathML-to-OMML via Plurimath) — this already works

Step 1 never happens in the DOCX pipeline. The `@asciimathdelims` option is stored in the Html2Doc constructor but the stem-to-MathML conversion is not invoked before `mathml_to_ooml`.

## Required changes

1. **Add AsciiMath-to-MathML conversion** in the DOCX pipeline, before `mathml_to_ooml` runs. This should happen in `cleanup` or in a new step in `process_docx`.

2. **Pass `asciimathdelims` to the pipeline** — the test invocation above didn't pass `asciimathdelims: ["`", "`"]`, but even when it is passed, the conversion needs to be triggered.

3. **Verify the full chain**: AsciiMath spans → MathML → OMML → Uniword OMathPara objects.

### Where the stem conversion happens in the legacy path

In the legacy MHT path, the HTML goes through `process_html` → `cleanup` → `define_head` → `msword_fix` — the stem-to-MathML conversion may happen as part of a Metanorma preprocessing step that's external to html2doc. Need to trace how `asciimathdelims` is actually used.

## Scope in rice.html

28 stem expressions including:
- Simple: `` `r` ``, `` `t_90` ``, `` `w` ``
- Subscripts/superscripts: `` `m_D` ``, `` `s_r` ``, `` `w_(wax)` ``
- Formulas: `` `r = 1 %` ``, `` `R = 3 %` ``, `` `w = (m_D) / (m_s)` ``
- Complex: `` `w_(wax) = (m_1) / (m_1 + m_2) xx 100` ``
- Matrix: `` `{:(+0.02),(0):}` ``

Some stem expressions are inside table cells, which makes this depend on item 1 (tables).

## Acceptance criteria

- All 28 stem expressions convert to `<m:oMath>` in the DOCX
- Inline stem (inside `<p>`) renders as `m:oMath` within a run
- Block stem (inside `<div class="formula">`) renders as `m:oMathPara` in its own paragraph
- Stem inside table cells works (depends on tables being fixed)
