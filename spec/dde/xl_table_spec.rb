require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest
  describe DDE::XlTable do

    it 'starts out empty and without item/topic' do
      table = DDE::XlTable.new
      table.should be_empty
      table.buf.should == nil
    end
  end
end