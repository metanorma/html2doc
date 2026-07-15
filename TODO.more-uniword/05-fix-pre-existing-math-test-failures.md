# 05: Fix Pre-Existing Math Test Failures

## Summary

Investigate and fix 6 pre-existing MathML test failures in html2doc's test suite.

## Motivation

These failures have existed since before the Uniword migration. They indicate bugs in the MathML→OMML conversion pipeline (via Plurimath) or in html2doc's post-processing. Fixing them ensures the DOCX path produces correct math output.

## Known Failures

The 6 failures are in `spec/html2doc_spec.rb` tests that use `<stem>` elements with MathML content. The failures manifest as incorrect OMML output (wrong structure, missing elements, incorrect formatting).

## Investigation Steps

### 1. Identify the specific failing tests

```bash
bundle exec rspec spec/html2doc_spec.rb 2>&1 | grep "FAILED\|Failure"
```

### 2. Categorize failures

Possible categories:
- **Plurimath conversion errors**: MathML → OMML produces incorrect XML
- **Post-processing errors**: html2doc's `math.rb` post-processing corrupts valid OMML
- **Comparison errors**: Expected output doesn't match due to whitespace/namespace differences

### 3. Fix each failure

For Plurimath issues: Report upstream or work around in html2doc.
For post-processing issues: Fix in `lib/html2doc/math.rb`.
For comparison issues: Normalize comparison in tests.

### 4. Verify fixes don't break other tests

```bash
bundle exec rspec
```

## Acceptance Criteria

- [ ] All 6 math test failures resolved
- [ ] No new test failures introduced
- [ ] Root cause documented for each fix
