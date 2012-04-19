# coding: utf-8
require 'rubygems'
require 'test/unit'
begin
  require 'redgreen'
rescue LoadError
end

require 'msworddoc-extractor'

class TestFareast < Test::Unit::TestCase
  def setup
    @doc = MSWordDoc::Extractor.load('test/fareast.doc')
  end

  def teardown
    @doc.close
  end

  def test_document
    assert_match %r{ 色は匂へど \s+ 散りぬるを }xm, @doc.document, "document"
  end

  def test_header
    assert_match %r{ いろは歌 }xm, @doc.header, "header"
  end

  def test_footnote
    assert_match %r{ いろはにほへとちりぬるを }xm, @doc.footnote, "footnote"
  end

end
