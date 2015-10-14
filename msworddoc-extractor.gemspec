# -*- encoding: utf-8 -*-

Gem::Specification.new do |gem|
  gem.authors       = ["ITO Nobuaki"]
  gem.email         = ["daydream.trippers@gmail.com"]
  gem.description   = %q{Extract text contents from Microsoft Word Document.}
  gem.summary       = %q{Extract text contents from Microsoft Word Document}
  gem.homepage      = ""

  gem.files         = [
    'bin/worddoc-extract',
    'lib/msworddoc-extractor.rb',
    'lib/msworddoc/extractor.rb',
    'test/test_msworddoc.rb',
    'test/test_fareast.rb',
    'test/test_io.rb',
    'test/lorem.doc',
    'test/fareast.doc',
  ]
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.name          = "msworddoc-extractor"
  gem.require_paths = ["lib"]
  gem.version       = '0.2.0'

  has_rake = RUBY_VERSION >= '1.9.'

  if gem.respond_to? :specification_version then
    gem.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      gem.add_runtime_dependency 'ruby-ole'
      gem.add_development_dependency 'rake'  unless has_rake
    else
      gem.add_dependency 'ruby-ole'
      gem.add_dependency 'rake'  unless has_rake
    end
  else
    gem.add_dependency 'ruby-ole'
    gem.add_dependency 'rake'  unless has_rake
  end
end
