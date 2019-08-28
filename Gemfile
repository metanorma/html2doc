Encoding.default_external = Encoding::UTF_8
Encoding.default_internal = Encoding::UTF_8

source "https://rubygems.org"

group :development, :test do
  gem "rspec"
end

gemspec

if File.exists?('Gemfile.devel') then
  eval File.read('Gemfile.devel'), nil, 'Gemfile.devel'
end
