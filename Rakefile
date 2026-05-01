require "bundler/gem_tasks"
require "rspec/core/rake_task"

RSpec::Core::RakeTask.new(:spec)

namespace :iso do
  desc "Extract styles, numbering, fonts, settings, and theme from the ISO DOCX fixture"
  task :extract_styles do
    require_relative "lib/html2doc/iso_style_extractor"
    Html2Doc::IsoStyleExtractor.extract_all
  end
end

task default: :spec
