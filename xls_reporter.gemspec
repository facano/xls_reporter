# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xls_reporter/version'

Gem::Specification.new do |spec|
  spec.name          = "xls_reporter"
  spec.version       = XlsReporter::VERSION
  spec.authors       = ["Felipe Cano"]
  spec.email         = ["felicanoo@gmail.com"]
  spec.summary       = "Export data to Excel file."
  spec.description   = "Export data to Excel file. Iterate in a collection and map values in each row to fill Excel spreadsheet"
  spec.homepage      = "https://github.com/facano"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "pry"

  spec.add_dependency "write_xlsx", "~> 0.76.0"
end
