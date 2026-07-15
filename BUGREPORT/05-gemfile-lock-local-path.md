# Bug 05: Gemfile.lock Lost Local Uniword Path Reference

**Status**: FIXED
**Severity**: High (caused 26 additional test failures when resolved from rubygems)
**Location**: `Gemfile`, `Gemfile.lock`

## Symptom

Running `bundle update uniword` resolved uniword from rubygems (published 1.0.7) instead of the local development copy at `../uniword/`. This caused 26 additional test failures because the published gem lacks the `run_position` attribute on `Hyperlink` and `OMath` that html2doc depends on.

## Root Cause

The `Gemfile` had uniword declared through `gemspec` only (which specifies `~> 1.0.6`). Without an explicit `path:` directive, bundler resolves from rubygems. The `Gemfile.lock` lost its `PATH remote: ../uniword` entry.

## Fix

Added `gem "uniword", path: "../uniword"` to the Gemfile before `gemspec`. This forces bundler to use the local development copy.

## Lesson

For local development with unreleased features, the Gemfile needs an explicit `path:` directive. Remove it before release to let bundler resolve from rubygems.

**Note**: The published uniword 1.0.7 on rubygems is behind the local version. The local copy has:
- `run_position` attribute on `Hyperlink` and `OMath`
- Possibly other unreleased improvements

A new uniword release (1.0.8+) must be published before html2doc can work with the rubygems version.
