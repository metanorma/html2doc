# 01: Clean Up PR #99 (feat/docx-output-via-uniword)

## Summary

Remove unused fixtures, dead code, and ISO-specific content from the html2doc PR branch before merge. The PR should contain only the HTML→DOCX converter and its tests.

## Motivation

The PR currently has 138 docx files in `spec/examples/`, of which 101 are unused debugging artifacts. There are 6 unused `data/iso_*` files and an unused `iso_style_extractor.rb`. These inflate the PR and confuse reviewers.

## Tasks

### 1. Remove unused spec/examples/*.docx files

Keep only the 37 files used by tests:
- 18 broken + 18 repaired pairs used by `uniword/spec/integration/repair_spec.rb`
- `rice.docx` used by `spec/rice_fixture_spec.rb`

Remove the 101 unused files:
- `f4_swap_*` (non-fixture debug artifacts)
- `r7_swap_*`, `r8_swap_*`, `r9_swap_*` (non-numbering swaps)
- `swap_test_*`, `test_fix_*` (debug artifacts)
- `rice_*` debugging variants (keep `rice.docx` only)
- `minimal_fresh*.docx`, `minimal_helloworld*.docx` variants (keep originals if used)

```bash
# First, identify what's actually used:
grep -roh 'spec/examples/[^"]*\.docx' spec/ | sort -u > /tmp/used.txt
# Then remove everything else
cd spec/examples && ls *.docx | grep -v -f /tmp/used.txt | xargs rm
```

### 2. Remove `data/` directory

These 6 files are never read at runtime — `StyleLoader` loads directly from the DOCX fixture:
- `data/iso_styles.xml`
- `data/iso_styles.yml`
- `data/iso_numbering.xml`
- `data/iso_font_table.xml`
- `data/iso_settings.xml`
- `data/iso_theme.xml`

Also remove the `Dir.glob("data/**/*.{xml,yml}")` line from `html2doc.gemspec:26`.

### 3. Remove `lib/html2doc/iso_style_extractor.rb`

This one-time extraction tool is not needed at runtime. The ISO template will be shipped in metanorma-iso instead.

### 4. Verify tests still pass after cleanup

```bash
bundle exec rspec
```

All 166 tests should still pass. No test should reference removed files.

### 5. Squash or rebase the cleanup commits

Keep the PR history clean — one cleanup commit at the end is fine.

## Acceptance Criteria

- [ ] Zero unused docx files in `spec/examples/`
- [ ] `data/` directory removed
- [ ] `iso_style_extractor.rb` removed
- [ ] `Dir.glob("data/**/*.{xml,yml}")` removed from gemspec
- [ ] All 166 tests pass
- [ ] PR diff is minimal and reviewable
