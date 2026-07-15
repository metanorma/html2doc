# 03: Refactor DocxConverter to Use Uniword::Builder

## Summary

Rewrite `DocxConverter` to use Uniword's `DocumentBuilder` API instead of directly constructing model objects. This reduces code by ~60% and improves readability.

## Motivation

The current `docx_converter.rb` (2159 lines) constructs low-level model objects directly:
```ruby
para = Wordprocessingml::Paragraph.new
run = Wordprocessingml::Run.new(text: "Hello")
run.properties = Wordprocessingml::RunProperties.new(
  bold: Properties::Bold.new(value: true)
)
para.runs << run
```

With Uniword's builder API, this becomes:
```ruby
doc.paragraph { |p| p << Builder.text("Hello", bold: true) }
```

The builder handles properties, validation, and assembly automatically.

## Prerequisites

- 02: Parameterize Template (template parameter needed)

## Tasks

### 1. Replace DocumentPackage assembly with DocumentBuilder

Current pattern (docx_converter.rb):
```ruby
pkg = Uniword::DocxPackage.from_file(template_path)
doc_root = pkg.document  # Get DocumentRoot
body = doc_root.body
body.paragraphs << para   # Add paragraphs directly
```

New pattern:
```ruby
doc = Uniword::Builder::DocumentBuilder.from_file(template_path)
doc.paragraph { |p| p << "text" }
doc.save(output_path)
```

### 2. Replace Paragraph/Run construction with builders

Map existing methods to builder calls:

| Current | Builder equivalent |
|---|---|
| `Wordprocessingml::Paragraph.new` | `ParagraphBuilder.new` |
| `Wordprocessingml::Run.new(text: s)` | `RunBuilder.new.text(s).build` |
| `run.properties.bold = Properties::Bold.new(value: true)` | `RunBuilder.new.bold.text(s).build` |
| `para.properties.style = Properties::StyleReference.new(value: id)` | `para.style = id` |
| `para.properties.alignment = Properties::Alignment.new(value: "center")` | `para.align = :center` |

### 3. Replace Table construction with TableBuilder

Current:
```ruby
table = Wordprocessingml::Table.new
row = Wordprocessingml::TableRow.new
cell = Wordprocessingml::TableCell.new
para = Wordprocessingml::Paragraph.new
run = Wordprocessingml::Run.new(text: "cell text")
para.runs << run
cell.paragraphs << para
row.cells << cell
table.rows << row
```

New:
```ruby
doc.table do |t|
  t.row do |r|
    r.cell(text: "cell text")
  end
end
```

### 4. Replace list construction with ListBuilder

Current: Complex manual numbering with `MsoListParagraphCxSp*`.

New:
```ruby
doc.list(type: :bullet) do |l|
  l.item("First item")
  l.item("Second item")
end
```

### 5. Replace footnote/endnote with FootnoteBuilder

Current: Manual footnote ID management, `FootnoteReference`, `Footnote` entry construction.

New:
```ruby
doc.paragraph do |p|
  p << "Some text"
  p << doc.footnote("Footnote text")
end
```

### 6. Replace header/footer with HeaderFooterBuilder

Current: Manual `Header`/`Footer` construction, section property wiring.

New:
```ruby
doc.header(type: "default") do |h|
  h << "Header content"
end

doc.footer(type: "default") do |f|
  f << Builder.page_number_field
end
```

### 7. Estimate line reduction

Current docx_converter.rb: ~2159 lines.
Expected after refactoring: ~800 lines (60% reduction).

Most savings come from:
- Eliminating manual property construction (bold, italic, color, size → RunBuilder chain)
- Eliminating manual table cell/row/table assembly (TableBuilder DSL)
- Eliminating manual footnote/endnote ID management (FootnoteBuilder)
- Eliminating manual numbering setup (ListBuilder)

### 8. Preserve all existing behavior

Every existing test must pass without modification. The refactoring is purely internal — the external API (`Html2Doc.new.process()`) stays the same.

## Acceptance Criteria

- [ ] DocxConverter uses DocumentBuilder internally
- [ ] All 166 existing tests pass without modification
- [ ] Line count reduced by ~60%
- [ ] No direct `Wordprocessingml::*` construction in DocxConverter
- [ ] All HTML elements handled: paragraphs, headings, tables, lists, footnotes, images, math, bookmarks, headers/footers
