lib = File.expand_path("lib", __dir__)
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

  spec.homepage = "https://github.com/metanorma/html2doc"
  spec.licenses = ["CC-BY-SA-3.0", "BSD-2-Clause"]

  spec.bindir        = "bin"
  spec.require_paths = ["lib"]
  spec.files         = `git ls-files -z`.split("\x0").reject do |f|
    f.match(%r{^(test|spec|features|bin|.github)/}) \
    || f.match(%r{Rakefile|bin/rspec})
  end
  spec.required_ruby_version = Gem::Requirement.new(">= 2.7.0")

  spec.add_dependency "base64"
  spec.add_dependency "htmlentities", "~> 4.3.4"
  spec.add_dependency "lutaml-model", "~> 0.7.0"
  spec.add_dependency "metanorma-utils", ">= 1.9.0"
  spec.add_dependency "mime-types"
  spec.add_dependency "nokogiri", "~> 1.18.3"
  spec.add_dependency "plane1converter", "~> 0.0.1"
  spec.add_dependency "plurimath", "~> 0.9.0"
  spec.add_dependency "thread_safe"
  spec.add_dependency "uuidtools"
  spec.add_dependency "unitsml"
  spec.add_dependency "vectory", "~> 0.8"

  spec.add_development_dependency "debug"
  spec.add_development_dependency "equivalent-xml", "~> 0.6"
  spec.add_development_dependency "guard", "~> 2.14"
  spec.add_development_dependency "guard-rspec", "~> 4.7"
  spec.add_development_dependency "rake", "~> 12.0"
  spec.add_development_dependency "rspec", "~> 3.6"
  spec.add_development_dependency "rspec-match_fuzzy", "~> 0.2.0"
  spec.add_development_dependency "rubocop", "~> 1"
  spec.add_development_dependency "rubocop-performance"
  spec.add_development_dependency "simplecov", "~> 0.15"
  spec.add_development_dependency "timecop", "~> 0.9"
end
