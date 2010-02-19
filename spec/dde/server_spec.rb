require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest

  describe DDE::Server do
    before(:each ){ @server = DDE::Server.new }

    it 'new without parameters creates Server but does not activate DDEML or start service' do
      @server.id.should == nil
      @server.service.should == nil
      @server.dde_active?.should == false
      @server.service_active?.should == false
      @server.table.should be_empty
    end

    it 'new with attached callback block creates Server and activates DDEML, but does not start service' do
      server = DDE::Server.new {|*args|}
      server.id.should be_an Integer
      server.id.should_not == 0
      server.dde_active?.should == true
      server.service.should == nil
      server.service_active?.should == false
      server.table.should be_empty
    end

    describe '#start_service' do

      context 'with inactive (uninitialized) DDE:' do
        it 'initializes DDE with attached block and starts new service' do
          res = @server.start_service('myservice') {|*args|}
          res.should == true

          @server.service.should be_a DDE::DdeString
          @server.service.should == 'myservice'
          @server.service.name.should == 'myservice'
          @server.service.handle.should be_an Integer
          @server.service.handle.should_not == 0
          @server.service_active?.should == true
        end

        it 'service name defaults to "excel" if not given explicitly' do
          res = @server.start_service {|*args|}
          res.should == true

          @server.service.should be_a DDE::DdeString
          @server.service.should == 'excel'
          @server.service.name.should == 'excel'
          @server.service.handle.should be_an Integer
          @server.service.handle.should_not == 0
          @server.service_active?.should == true
        end

        it 'fails to starts new service without block' do
          lambda{@server.start_service('myservice')}.should raise_error DDE::Errors::ServiceError
          @server.service_active?.should == false
          lambda{@server.start_service}.should raise_error DDE::Errors::ServiceError
          @server.service_active?.should == false
        end

      end

      context 'with active (initialized) DDE:' do
        before(:each ){ @server = DDE::Server.new {|*args|}}

        it 'starts new service if DDE is already activated' do
          res = @server.start_service 'myservice'
          res.should == true

          @server.service.should be_a DDE::DdeString
          @server.service.should == 'myservice'
          @server.service.name.should == 'myservice'
          @server.service.handle.should be_an Integer
          @server.service.handle.should_not == 0
        end


        it 'fails to starts new service (if DDE is not active yet)' do
          lambda{@server.start_service(11)}.should raise_error DDE::Errors::ServiceError
          @server.service_active?.should == false
        end

      end
    end
  end
end