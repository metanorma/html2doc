require "simplecov"
SimpleCov.start do
  add_filter "/spec/"
end

require "bundler/setup"
require "rspec/match_fuzzy"
require "html2doc"
require "rspec/matchers"
require "equivalent-xml"
require "canon"

Canon::Config.instance.tap do |cfg|
  # Configure Canon to use spec-friendly match profiles
  cfg.xml.match.profile = :spec_friendly
  cfg.html.match.profile = :spec_friendly

  # Configure Canon to show all diffs (including inactive diffs)
  cfg.html.diff.show_diffs = :all
  cfg.xml.diff.show_diffs = :all

  # Enable verbose diff output for debugging
  cfg.html.diff.verbose_diff = true
  cfg.xml.diff.verbose_diff = true
end

RSpec.configure do |config|
  # Enable flags like --only-failures and --next-failure
  config.example_status_persistence_file_path = ".rspec_status"

  # Disable RSpec exposing methods globally on `Module` and `main`
  config.disable_monkey_patching!

  config.expect_with :rspec do |c|
    c.syntax = :expect
  end
end
