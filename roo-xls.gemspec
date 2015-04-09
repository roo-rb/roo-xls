# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'roo/xls/version'

Gem::Specification.new do |spec|
  spec.name          = "roo-xls"
  spec.version       = Roo::Xls::VERSION
  spec.authors       = ["Thomas Preymesser", "Hugh McGowan", "Ben Woosley"]
  spec.email         = ["ruby.ruby.ruby.roo@gmail.com"]
  spec.summary       = "Roo::Excel can access the contents of classic xls files."
  spec.description   = "Roo can access the contents of various spreadsheet files. It can handle\n* OpenOffice\n* Excel\n* Google spreadsheets\n* Excelx\n* LibreOffice\n* CSV"
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency "roo", ">= 2.0.0beta1", "< 3"
  spec.add_dependency "nokogiri"
  spec.add_dependency "spreadsheet", "> 0.9.0"

  spec.add_development_dependency "bundler", ">= 1.7"
  spec.add_development_dependency "rake", ">= 10.0"
end
