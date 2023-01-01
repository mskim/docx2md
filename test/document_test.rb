require File.dirname(File.expand_path(__FILE__)) + "/test_helper"


describe 'parse book.docx' do
  before do
    @docx_path = File.dirname(__FILE__) + '/book.docx'
    @markdown_path = File.dirname(__FILE__) + '/book.md'
    doc = Docx2md::Document.open(@docx_path)
    md = doc.to_markdown
    File.open(@markdown_path,'w'){|f| f.write md }
  end

  it 'should sace md file' do
    assert File.exist?(@markdown_path) 
  end

end
