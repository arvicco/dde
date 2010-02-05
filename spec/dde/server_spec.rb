require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest
  describe DDE::Server do
    it "starts" do
      server = DDE::Server.new
      server.name.should == 'excel'
      server.id.should == 0
    end

    it "starts with given name" do
      server = DDE::Server.new 'my_server'
      server.name.should == 'my_server'
      server.id.should == 0
    end

    it 'connects to DDE' do
      server = DDE::Server.new
      res = server.connect &dde_callback
      res.should == true
    end
  end
end