require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')
require File.expand_path(File.dirname(__FILE__) + '/server_shared')

module DDETest

  describe DDE::Server, ' in general:' do
    it_should_behave_like "DDE Server"
  end

  describe DDE::Server do
    before(:each ){ @server = DDE::Server.new }
    after(:each) do
      @server.stop_service if @server.service_active?
      @server.stop_dde if @server.dde_active?
    end

    it 'new without parameters creates Server but does not activate DDEML or start service' do
      @server.id.should == nil
      @server.service.should == nil
      @server.dde_active?.should == false
      @server.service_active?.should == false
    end

    describe '#start_service' do

      it 'service name should be given explicitly' do
        expect{@server.start_dde{|*args|}.start_service}.to raise_error ArgumentError, /0 for 1/
        expect{@server.start_service {|*args|}}.to raise_error ArgumentError, /0 for 1/
      end

      it 'callback block should be given explicitly' do
        lambda{@server.start_service('myservice')}.should raise_error DDE::Errors::ServiceError
        @server.service_active?.should == false
      end
    end #describe '#start_service'
  end # describe DDE::Server do
end