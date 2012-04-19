require 'rubygems'
require 'test/unit'
begin
  require 'redgreen'
rescue LoadError
end

require 'msworddoc-extractor'

class TestMsworddoc < Test::Unit::TestCase
  def setup
    @doc = MSWordDoc::Extractor.load('test/lorem.doc')
  end

  def teardown
    @doc.close
  end

  def test_document
    assert_match %r{ \A Lorem \s+ ipsum \s+ }xm, @doc.document, "document"
    assert_match %r{ \s+ Duis \s+ aute \s+ }xm, @doc.document, "document"
    assert_match %r{ \s+ id \s+ est \s+ laborum[.] }xm, @doc.document, "document"
  end

  def test_header
    assert_match %r{ \A Lorem \s+ ipsum \s+ ... }xm, @doc.header, "header"
  end

  def test_footnote
    assert_match %r{ The \s+ quick \s+ brown \s+ fox \s+ }xm, @doc.footnote, "footnote"
    assert_match %r{ \s+ jumps \s+ over \s+ the \s+ lazy \s+ dog[.] }xm, @doc.footnote, "footnote"
  end

end
