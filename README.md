# MSWordDoc::Extractor

Extract text contents from Microsoft Word Document

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'msworddoc-extractor', :git =>
  'git://github.com/dayflower/msworddoc-extractor.git'
```

And then execute:

    $ bundle install

## Usage

```ruby
require 'msworddoc-extractor'

doc = MSWordDoc::Extractor.load('sample.doc')
puts doc.contents   # doc is MSWordDoc::Essence
# You have to close document explicitly
doc.close()

# Or call load() with block argument (recommended way)
MSWordDoc::Extractor.load('sample.doc') do |doc|
  puts doc.header
end
```

### Properties of `MSWordDoc::Essence`

* `document`
* `header`
* `footnote`
* `macro`
* `annotation`
* `endnote`
* `textbox`
* `header_textbox`
* `whole_contents`

## Limitations

Only supports Microsoft Word binary document.  
Does not support Microsoft Word XML document (.docx).

This module does not handle `PAP` (PAragraph Properties) and `CHP` (CHaracter
Properties), that define paragraphs and characters style.
Those styling information are required to determine functionalities of
some of special characters (such as row mark, footnote reference, and etc),
but are just ignored in the module, so extracted text will be inaccurate.

Also this module does not handle summary information stream in Word file.

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Added some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
