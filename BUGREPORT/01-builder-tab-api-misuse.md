# Bug 01: Uniword Builder.tab API Misuse in TocBuilder

**Status**: FIXED
**Severity**: High (caused 12 rice fixture test failures, blocked DOCX output)
**Location**: `lib/html2doc/toc_builder.rb:160`
**Root cause**: `uniword/lib/uniword/builder.rb:154`

## Symptom

All 12 rice fixture DOCX tests failed with:

```
NoMethodError: undefined method 'val' for an instance of Uniword::Wordprocessingml::Run
```

Stack trace pointed to `lutaml-model` serialization:

```
uniword/wordprocessingml/document_root.rb:89:in 'to_xml'
uniword/docx/package_serialization.rb:81:in 'serialize_package_parts'
```

## Root Cause

`Uniword::Builder.tab` returns a **`Run`** object (a Run containing a Tab), not a bare `Tab`:

```ruby
# uniword/lib/uniword/builder.rb:154
def self.tab
  run = Wordprocessingml::Run.new
  run.tab = Wordprocessingml::Tab.new
  run  # returns Run, not Tab
end
```

`toc_builder.rb:160` assigned this to `run.tab`:

```ruby
run.tab = Uniword::Builder.tab  # assigns a Run into a Tab slot
```

lutaml-model's setters don't type-check, so this silently succeeded. The `@tab` instance variable held a `Run` instead of a `Tab`.

## Why It Manifested as `NoMethodError: undefined method 'val'`

During XML serialization:

1. lutaml-model compiles a `Transformation` for `Run` with a rule: `attribute_name=:tab, attribute_type=Tab`
2. `extract_rule_value` calls `model_instance.public_send(:tab)` which returns the incorrectly stored `Run`
3. Since the value is not nil and `Tab` is a `Serializable`, lutaml-model treats it as a nested model and applies Tab's transformation to it
4. Tab's compiled rules include `attribute_name=:val` (Tab has `attribute :val, :string`)
5. Tab's `extract_rule_value` calls `model_instance.public_send(:val)` on the `Run`, which has no `val` method
6. `method_missing` only handles setters, so it raises `NoMethodError`

## Fix

Changed `toc_builder.rb:160` from:

```ruby
run.tab = Uniword::Builder.tab
```

to:

```ruby
run.tab = Uniword::Wordprocessingml::Tab.new
```

## Upstream Recommendation

The `Builder.tab` API is documented as returning a `Run`, but the name `tab` strongly implies it returns a `Tab`. Either:

1. Rename to `Builder.tab_run` (consistent with `Builder.page_break` which also returns a Run)
2. Add a separate `Builder.create_tab` that returns a bare `Tab`

See Uniword PR (pending) for the upstream fix.

## Reproduction

```ruby
require "uniword"

run = Uniword::Wordprocessingml::Run.new
run.text = "test"
run.tab = Uniword::Builder.tab  # Bug: assigns Run, not Tab

doc = Uniword::Wordprocessingml::DocumentRoot.new
para = Uniword::Wordprocessingml::Paragraph.new
para.runs << run
doc.body.paragraphs << para

doc.to_xml  # => NoMethodError: undefined method 'val' for Run
```
