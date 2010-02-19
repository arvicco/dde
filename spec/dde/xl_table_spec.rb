require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest
  describe DDE::XLTable do
    before(:each ){ @server = DDE::Server.new }

    it 'starts out empty and without item/topic' do
      table = DDE::XLTable.new
      table.should be_empty
      table.buf.should == nil
    end
  end
end