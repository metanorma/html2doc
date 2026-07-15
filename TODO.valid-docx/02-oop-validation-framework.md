# Plan 02: OOP Validation Framework in Uniword

## Goal
Ensure the Uniword validation system is fully OOP, model-driven, open/closed, and encapsulates all validity rules. The system should be extensible for new rules without modifying existing code.

## Current state

The validation system already has a solid foundation:
- **7-layer pipeline** (`DocumentValidator`) — file structure, ZIP, OOXML parts, XML schema, relationships, content types, semantics
- **12 semantic rules** (`Validation::Rules::*Rule`) — each extends `Base`, has code/category/severity, registered in `Rules::Registry`
- **Layer validators** (`Validators::*Validator`) — each extends `LayerValidator`
- **`DocumentContext`** — provides lazy-loaded parsed XML access to rules

## Gaps to fill

### Gap 1: Missing semantic rules for discovered validity issues

New rules needed (the existing DOC-001..DOC-091 don't cover these):

| New Rule | Code | Category | What it checks |
|----------|------|----------|----------------|
| McIgnorableNamespaceRule | DOC-100 | namespaces | mc:Ignorable prefixes have xmlns declarations in scope |
| SettingsValuesRule (extend) | DOC-101 | settings | w15:docId is GUID format, w14:docId is hex string |
| ThemeCompletenessRule | DOC-102 | theme | fmtScheme has minimum required children |
| NumberingPreservationRule | DOC-103 | numbering | lvlOverride/startOverride preserved |
| SectionPropertiesRule | DOC-104 | structure | sectPr exists with pgSz and pgMar |
| CorePropertiesNamespaceRule | DOC-105 | core-properties | dcterms/dc/xsi namespaces declared, xsi:type on dates |
| ContentTypesCoverageRule (extend) | DOC-106 | content-types | every ZIP entry has content type coverage |
| FontTableSignatureRule | DOC-107 | fonts | sig elements have valid attributes |

### Gap 2: Model-driven validation (not string/XML-based)

Current rules use Nokogiri XML parsing and XPath queries. For model-driven validation:

- Add a `validate` or `check_consistency` method to key model classes that returns issues
- OR enhance `DocumentContext` to provide model-level access (via `Package.from_file`)
- The reconciler already works at the model level; validation should too

**Approach:** Add an optional `model` accessor to `DocumentContext` that provides a lazily-loaded `Package` instance. Rules can then validate model objects directly instead of parsing raw XML.

```ruby
class DocumentContext
  def model
    @model ||= Uniword::Docx::Package.from_file(@path)
  end
end
```

### Gap 3: Reconciler rule tracking

The reconciler applies rules but doesn't track what it did. Add an audit trail:

```ruby
class Reconciler
  attr_reader :applied_fixes

  def initialize(package, profile: nil)
    @package = package
    @profile = profile
    @applied_fixes = []
  end

  private

  def record_fix(code, message)
    @applied_fixes << { code: code, message: message, timestamp: Time.now }
  end
end
```

Each reconciler method calls `record_fix("R1", "Added namespace_scope for w14")` when it applies a fix.

### Gap 4: Rule documentation in code

Each rule class should have a `description` method and reference the validity rule ID:

```ruby
class McIgnorableNamespaceRule < Base
  def code = "DOC-100"
  def validity_rule = "R1"  # references docs/docx-valid/rules/namespace-consistency.adoc
  def description = "mc:Ignorable prefixes must have xmlns declarations in scope"
  def category = :namespaces
  def severity = "error"
end
```

### Gap 5: Reconciler ↔ Validation round-trip

After the reconciler runs, run validation to verify no issues remain. Add to `Package#save`:

```ruby
def save(path)
  reconciler.reconcile
  # Optional: verify reconciliation was complete
  if @validate_after_reconcile
    report = Validation::DocumentValidator.new.validate(temp_path)
    # log or raise on any error-level issues
  end
end
```

## Implementation steps

1. **Add `validity_rule` and `description` to `Base`** — default nil, subclasses override
2. **Create 8 new rule classes** in `lib/uniword/validation/rules/`:
   - `mc_ignorable_namespace_rule.rb` (DOC-100, R1)
   - `settings_values_rule.rb` (DOC-101, R2) — extends existing settings_rule.rb or new
   - `theme_completeness_rule.rb` (DOC-102, R3)
   - `numbering_preservation_rule.rb` (DOC-103, R4+R5)
   - `section_properties_rule.rb` (DOC-104, R11)
   - `core_properties_namespace_rule.rb` (DOC-105, R14)
   - `content_types_coverage_rule.rb` (DOC-106, R7)
   - `font_table_signature_rule.rb` (DOC-107, R13)
3. **Register all new rules** in `Rules::Registry`
4. **Add `model` accessor** to `DocumentContext` (lazy Package.from_file)
5. **Add audit trail** to Reconciler (`applied_fixes`, `record_fix`)
6. **Update docs/_verification/semantic-rules.adoc** with new rules
7. **Write tests** for each new rule class in `spec/validation/rules/`

## File structure after changes

```
lib/uniword/validation/rules/
  base.rb                          — add validity_rule, description methods
  registry.rb                      — no changes needed
  document_context.rb              — add model accessor
  mc_ignorable_namespace_rule.rb   — NEW: DOC-100
  settings_values_rule.rb          — NEW: DOC-101
  theme_completeness_rule.rb       — NEW: DOC-102
  numbering_preservation_rule.rb   — NEW: DOC-103
  section_properties_rule.rb       — NEW: DOC-104
  core_properties_namespace_rule.rb — NEW: DOC-105
  content_types_coverage_rule.rb   — NEW: DOC-106
  font_table_signature_rule.rb     — NEW: DOC-107
  bookmarks_rule.rb                — existing, no changes
  fonts_rule.rb                    — existing, no changes
  footnotes_rule.rb                — existing, no changes
  headers_footers_rule.rb          — existing, no changes
  images_rule.rb                   — existing, no changes
  numbering_rule.rb                — existing, no changes
  settings_rule.rb                 — existing, no changes
  style_references_rule.rb         — existing, no changes
  tables_rule.rb                   — existing, no changes
  theme_rule.rb                    — existing, no changes

lib/uniword/docx/reconciler.rb     — add applied_fixes audit trail
```
