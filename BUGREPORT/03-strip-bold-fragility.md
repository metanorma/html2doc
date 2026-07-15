# Bug 03: Fragile RunProperties Reconstruction in strip_bold_from_runs

**Status**: FIXED
**Severity**: Medium (latent bug — adding new RunProperties attributes would be silently dropped)
**Location**: `lib/html2doc/docx_converter.rb:312` (original)

## Symptom

`strip_bold_from_runs` reconstructed the entire `RunProperties` object by explicitly listing every property except `bold`:

```ruby
run.properties = Uniword::Wordprocessingml::RunProperties.new(
  style: run.properties.style,
  italic: run.properties.italic,
  underline: run.properties.underline,
  vertical_align: run.properties.vertical_align,
  fonts: run.properties.fonts,
  color: run.properties.color,
  size: run.properties.size,
)
```

If a new property were added to `RunProperties` (e.g., `shading_fill`, `highlight`, `small_caps`), it would be silently dropped during heading conversion. This is an open/closed principle violation.

The same pattern existed in `apply_cell_formatting`.

## Root Cause

The code was written to work around a (false) assumption that `run.properties.bold = nil` wouldn't work. In lutaml-model, attribute setters accept `nil` and the attribute defaults to `nil`, so setting it to `nil` is equivalent to removing it.

## Fix

`strip_bold_from_runs` reduced from 30 lines to 6:

```ruby
def strip_bold_from_runs(para)
  strip_bold = ->(run) {
    run.properties.bold = nil if run.properties&.bold
  }
  para.runs.each(&strip_bold)
  para.hyperlinks.each { |hl| hl.runs.each(&strip_bold) }
end
```

`apply_cell_formatting` reduced from 25 lines to 11:

```ruby
def apply_cell_formatting(para, bold, align)
  if bold
    para.runs.each do |r|
      r.properties ||= Uniword::Wordprocessingml::RunProperties.new
      r.properties.bold = Uniword::Properties::Bold.new
    end
  end
  if align && %w[left center right justify].include?(align)
    para.properties ||= Uniword::Wordprocessingml::ParagraphProperties.new
    para.properties.alignment = Uniword::Properties::Alignment.new(value: align == "justify" ? "both" : align)
  end
end
```

## Lesson

Never reconstruct value objects to remove a single property. Use direct assignment (`prop = nil`) instead.
