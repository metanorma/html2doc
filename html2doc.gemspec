
lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "html2doc/version"

Gem::Specification.new do |spec|
  spec.name          = "html2doc"
  spec.version       = Html2Doc::VERSION
  spec.authors       = ["Ribose Inc."]
  spec.email         = ["open.source@ribose.com"]

  spec.summary       = "Convert HTML document to Microsoft Word document"
  spec.description   = <<~DESCRIPTION
    Convert HTML document to Microsoft Word document.

    This gem is in active development.
  DESCRIPTION

  spec.homepage      = "https://github.com/riboseinc/html2doc"
  spec.licenses       = ["CC-BY-SA-3.0", "BSD-2-Clause"]

  spec.bindir        = "bin"
  spec.require_paths = ["lib"]
  spec.files         = `git ls-files`.split("\n")
  spec.test_files    = `git ls-files -- {spec}/*`.split("\n")
  spec.required_ruby_version = Gem::Requirement.new(">= 2.3.0")

  spec.add_dependency "htmlentities", "~> 4.3.4"
  spec.add_dependency "image_size"
  spec.add_dependency "mime-types"
  spec.add_dependency "nokogiri"
  spec.add_dependency "thread_safe"
  spec.add_dependency "uuidtools"
  spec.add_dependency "parallel"
  spec.add_dependency "ruby-progressbar"
  #spec.add_dependency "ruby-xslt"
  spec.add_dependency "asciimath", "~> 1.0.7"

  spec.add_development_dependency "bundler", "~> 2.0.1"
  spec.add_development_dependency "byebug", "~> 9.1"
  spec.add_development_dependency "equivalent-xml", "~> 0.6"
  spec.add_development_dependency "guard", "~> 2.14"
  spec.add_development_dependency "guard-rspec", "~> 4.7"
  spec.add_development_dependency "rake", "~> 12.0"
  spec.add_development_dependency "rspec", "~> 3.6"
  spec.add_development_dependency "rubocop", "= 0.54.0"
  spec.add_development_dependency "simplecov", "~> 0.15"
  spec.add_development_dependency "timecop", "~> 0.9"
  spec.add_development_dependency "rspec-match_fuzzy", "~> 0.1.3"
end
