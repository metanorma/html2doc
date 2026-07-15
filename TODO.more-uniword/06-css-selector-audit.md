# 06: CSS Selector Audit & Cleanup

## Summary

Audit all CSS selectors in `docx_converter.rb` to ensure none use XPath or regex on HTML content. Replace any remaining XPath/regex with pure CSS selectors.

## Motivation

The user explicitly requires CSS selectors only for HTML parsing in html2doc. XPath and regex on HTML are fragile and don't work with Nokogiri's HTML parser (which doesn't register namespaces the same way as XML). Two bugs were already fixed (CSS.escape NameError, m\:r XPath SyntaxError), but a full audit ensures no similar issues remain.

## Tasks

### 1. Find all selectors in docx_converter.rb

```bash
grep -n 'css\|xpath\|search\|at_css\|at_xpath\|\/\//' lib/html2doc/docx_converter.rb
```

### 2. For each selector, verify it's CSS (not XPath)

XPath patterns to flag:
- Paths with `//` or `.//`
- Paths with `@attribute`
- Paths with `text()`, `node()`, `*`
- Namespaced paths like `m:r`, `m:oMath`

Regex patterns to flag on parsed HTML:
- `element.inner_html.match?(/regex/)`
- `element.text.gsub(/regex/)`
- String manipulation that should use DOM methods

### 3. Fix any non-CSS selectors

Replace with CSS equivalents:
- `doc.xpath("//m:r")` → `doc.traverse { |n| ... if n.name == "r" && n.namespace&.prefix == "m" }`
- `element.at_xpath(".//text()")` → `element.css("text")` or `element.text`
- Regex on inner_html → DOM traversal + CSS selectors

### 4. Add a linter rule (optional)

Create a RuboCop custom rule that warns on `xpath` calls in `docx_converter.rb`:
```ruby
# In .rubocop.yml or a custom cop
Html2doc/NoXPathOnHtml:
  Enabled: true
  Include: ['lib/html2doc/docx_converter.rb']
```

## Acceptance Criteria

- [ ] Zero XPath calls in docx_converter.rb
- [ ] Zero regex on parsed HTML content in docx_converter.rb
- [ ] All CSS selectors tested and working
- [ ] No selector-related test failures
