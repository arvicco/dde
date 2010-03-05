require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')
require File.expand_path(File.dirname(__FILE__) + '/server_shared')

module DDETest

  describe DDE::XlServer do
    before(:each ){ @server = DDE::XlServer.new }
    after(:each) do
      @server.stop_service if @server.service_active?
      @server.stop_dde if @server.dde_active?
    end

    it_should_behave_like "DDE Server"

    it 'new without parameters creates empty table attribute' do
      @server.table.should be_an DDE::XlTable
      @server.table.should be_empty
    end

    it 'new with attached callback block creates empty table attribute' do
      server = DDE::XlServer.new {|*args|}
      @server.table.should be_an DDE::XlTable
      @server.table.should be_empty
    end

    describe '#start_service' do
      context 'with inactive (uninitialized) DDE:' do
        it 'service name defaults to "excel" if not given explicitly' do
          @server.start_service {|*args|}.should be_true

          @server.service.should be_a DDE::DdeString
          @server.service.should == 'excel'
          @server.service.name.should == 'excel'
          @server.service.handle.should be_an Integer
          @server.service.handle.should_not == 0
          @server.service_active?.should == true
        end

        it 'starts new service with default callback block' do
          @server.start_service('myservice1')
          @server.service_active?.should == true
          @server.service.name.should == 'myservice1'
        end
      end

      context 'with active (initialized) DDE:' do
        before(:each ){ @server = DDE::XlServer.new {|*args|}}

        it 'service name defaults to "excel" if not given explicitly' do
          @server.start_service.should be_true

          @server.service.should be_a DDE::DdeString
          @server.service.should == 'excel'
          @server.service.name.should == 'excel'
          @server.service.handle.should be_an Integer
          @server.service.handle.should_not == 0
        end
      end
    end # describe '#start_service'

  end
end
