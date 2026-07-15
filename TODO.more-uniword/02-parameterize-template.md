# 02: Parameterize DOCX Template

## Summary

Make the DOCX template path a parameter of `Html2Doc.new` instead of hardcoding the ISO fixture in `StyleLoader`.

## Motivation

`StyleLoader::DEFAULT_TEMPLATE` currently points to `spec/fixtures/iso-damd-fdis-sample.docx`. This forces all html2doc users to use the ISO template. A general-purpose HTML→DOCX converter should accept any template.

## Tasks

### 1. Add `template` option to Html2Doc constructor

In `lib/html2doc/base.rb`, accept the template option:

```ruby
class Html2Doc
  def initialize(options = {})
    @template = options[:template]  # Path to DOCX template
    # ... existing options
  end
end
```

### 2. Pass template to StyleLoader

In `lib/html2doc/style_loader.rb`, replace the hardcoded default:

```ruby
module StyleLoader
  class << self
    def template_package(template_path)
      return nil unless template_path
      TEMPLATE_CACHE[template_path] ||= load_template(template_path)
    end

    def build_class_to_style_map(template_path)
      pkg = template_package(template_path)
      return {} unless pkg
      # ... existing mapping logic
    end
  end
end
```

### 3. Provide a sensible default template

If no template is provided, create a minimal DOCX using Uniword's `DocumentBuilder`:

```ruby
def self.default_template
  @default_template ||= begin
    doc = Uniword::Builder::DocumentBuilder.new
    # Add basic styles: Normal, Heading1-6, Title, etc.
    path = File.join(Dir.tmpdir, "html2doc_default_template.docx")
    doc.save(path)
    path
  end
end
```

Or bundle a minimal template with html2doc at `data/default_template.docx`.

### 4. Pass template through DocxConverter

`DocxConverter.convert(docxml)` needs the template path:

```ruby
def process_docx(docxml, filename)
  docxml = to_xhtml(docxml)
  docxml = cleanup(docxml)
  DocxConverter.convert(docxml, template: @template)
end
```

### 5. CLI support

In `bin/html2doc`, add `--template` option:

```ruby
opts.on("--template PATH", "DOCX template file") do |v|
  options[:template] = v
end
```

### 6. Remove hardcoded ISO fixture reference

Delete `StyleLoader::DEFAULT_TEMPLATE` constant. The template is always provided by the caller or defaults to a minimal built-in template.

## Acceptance Criteria

- [ ] `Html2Doc.new(template: "path/to.docx")` loads that template's styles
- [ ] `Html2Doc.new` without template uses a minimal default
- [ ] CLI `--template` option works
- [ ] Existing tests updated to pass template explicitly
- [ ] No hardcoded ISO fixture paths in production code
