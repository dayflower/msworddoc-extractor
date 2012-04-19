require 'rubygems'
require 'test/unit'
begin
  require 'redgreen'
rescue LoadError
end

require 'msworddoc-extractor'
require 'stringio'

class TestIO < Test::Unit::TestCase
  def test_fileio
    open('test/lorem.doc', 'r') do |file|
      MSWordDoc::Extractor.load(file) do |doc|
        assert_match %r{ \A Lorem \s+ ipsum \s+ }xm, doc.document, "document"
      end
    end
  end

  def test_stringio
    data = ''
    open('test/lorem.doc', 'r') do |file|
      data = file.read()
    end

    io = StringIO.new(data, 'r')

    MSWordDoc::Extractor.load(io) do |doc|
      assert_match %r{ \A Lorem \s+ ipsum \s+ }xm, doc.document, "document"
    end

    io.close()
  end
end
