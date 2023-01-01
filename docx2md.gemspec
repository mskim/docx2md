$:.unshift File.expand_path("../lib", __FILE__)
require 'docx2md/version'

Gem::Specification.new do |s|
  s.name        = 'docx2md'
  s.version     = Docx2md::VERSION
  s.licenses    = ['MIT']
  s.summary     = 'a ruby library/gem for converting docx files to markdown'
  s.description = 'thin wrapper around rubyzip and nokogiri as a way to get started with docx files'
  s.authors     = ['Min Soo Kim']
  s.email       = ['mskim@gmail.com']
  s.homepage    = 'https://github.com/mskim/docx2md'
  s.files       = Dir['README.md', 'LICENSE.md', 'lib/**/*.rb']
  s.required_ruby_version = '>= 2.6.0'

  s.add_dependency 'nokogiri', '~> 1.13', '>= 1.13.0'
  s.add_dependency 'rubyzip',  '= 3.0.0'

  s.add_development_dependency 'coveralls_reborn', '~> 0.21'
  s.add_development_dependency 'rake', '~> 13.0'
  s.add_development_dependency 'rspec', '~> 3.7'
end
