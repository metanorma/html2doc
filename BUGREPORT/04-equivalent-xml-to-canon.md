# Bug 04: equivalent-xml Dependency Replaced With Canon

**Status**: FIXED
**Severity**: Low (migration, not a bug — but blocked test execution)
**Location**: `spec/spec_helper.rb`, `spec/html2doc_spec.rb:486`

## Context

The `equivalent-xml` gem was a dev dependency for XML comparison in specs. It was replaced with `canon` (from the lutaml ecosystem) which provides `be_xml_equivalent_to` as the replacement for `be_equivalent_to`.

## Changes

1. `Gemfile`: Replaced `gem "equivalent-xml", "~> 0.6"` with `gem "canon"`
2. `spec/spec_helper.rb`: Replaced `require "equivalent-xml"` with `require "canon"` + `Canon::Config.configure`
3. `spec/html2doc_spec.rb:486`: Changed `be_equivalent_to` → `be_xml_equivalent_to`

Only one test used `be_equivalent_to` — the MathML `oMathPara` comparison test.
