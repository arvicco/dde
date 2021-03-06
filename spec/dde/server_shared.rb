require File.expand_path(File.dirname(__FILE__) + '/app_shared')

module DdeTest
  shared_examples_for "DDE Server" do

    it_should_behave_like "DDE App"

    context 'specific to Servers:' do
      before(:each ){ @server = described_class.new {|*args|}}
      after(:each ){ @server.stop_dde if @server.dde_active?}

      it 'new with attached callback block creates Server and activates DDEML, but does not start service' do
        @server = described_class.new {|*args|}
        @server.id.should be_an Integer
        @server.id.should_not == 0
        @server.dde_active?.should == true
        @server.service.should == nil
        @server.service_active?.should == false
      end

      describe '#start_service' do

        context 'with inactive (uninitialized) DDE:' do
          it 'with attached block, initializes DDE and starts new service' do
            @server.start_service('myservice') {|*args|}.should be_true

            @server.service.should be_a Dde::DdeString
            @server.service.should == 'myservice'
            @server.service.name.should == 'myservice'
            @server.service.handle.should be_an Integer
            @server.service.handle.should_not == 0
            @server.service_active?.should == true
          end

          it 'returns self if success (allows method chain)' do
            res = @server.start_service('myservice') {|*args|}
            res.should == @server
          end
        end # context 'with inactive (uninitialized) DDE:'

        context 'with active (initialized) DDE:' do

          it 'starts new service with given name' do
            res = @server.start_service 'myservice'
            res.should be_true

            @server.service.should be_a Dde::DdeString
            @server.service.should == 'myservice'
            @server.service.name.should == 'myservice'
            @server.service.handle.should be_an Integer
            @server.service.handle.should_not == 0
          end

          it 'fails to starts new service if name is not a String' do
            lambda{@server.start_service(11)}.should raise_error Dde::Errors::ServiceError
            @server.service_active?.should == false
          end
        end # context 'with active (initialized) DDE:'
      end # describe '#start_service'

      describe '#stop_service' do
        context 'with inactive (uninitialized) DDE:' do
          it 'fails to stop service' do
            lambda{@server.stop_service}.should raise_error Dde::Errors::ServiceError
          end
        end

        context 'with active (initialized) DDE: and already registered DDE service: "myservice"' do
          before(:each){ @server = described_class.new {|*args|}; @server.start_service('myservice')}
          after(:each){ @server.stop_service if @server.service_active?}

          it 'stops previously registered service' do
            @server.stop_service.should be_true

            @server.service.should == nil
            @server.service_active?.should == false
          end

          it 'does not stop DDE instance' do
            @server.stop_service
            @server.id.should_not == nil
            @server.dde_active?.should == true
          end

          it 'returns self if success (allows method chain)' do
            @server.stop_service.should == @server
          end
        end # context 'with active (initialized) DDE: and already registered DDE service: "myservice"'
      end # describe '#stop_service'
    end # context 'specific to Servers:'
  end # shared_examples_for "DDE Server"
end