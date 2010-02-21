require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest

  describe DDE::Monitor do
    after(:each){@monitor.stop_dde}
    
    it 'starts without constructor parameters' do
      @monitor = DDE::Monitor.new

      @monitor.id.should be_an Integer
      @monitor.id.should_not == 0
      @monitor.dde_active?.should == true
    end

#    context 'with existing DDE clients and server supporting "service" topic' do
#      before(:each )do
#        @client_calls = []
#        @server_calls = []
#        @client = DDE::Client.new {|*args| @client_calls << args; 1}
#        @server = DDE::Server.new {|*args| @server_calls << args; 1}
#        @server.start_service('service')
#      end
#
#      it 'starts new conversation if DDE is already activated' do
#        res = @client.start_conversation 'service', 'topic'
#        res.should == true
#        @client.conversation_active?.should == true
#      end
#
#      it 'sets @conversation, @service and @topic attributes' do
#        @client.start_conversation 'service', 'topic'
#
#        @client.conversation.should be_an Integer
#        @client.conversation.should_not == 0
#        @client.service.should be_a DDE::DdeString
#        @client.service.should == 'service'
#        @client.service.name.should == 'service'
#        @client.conversation.should be_an Integer
#        @client.conversation.should_not == 0
#      end
#
#      it 'initiates XTYP_CONNECT transaction to service`s callback' do
#        @client.start_conversation 'service', 'topic'
#
#        @server_calls.first[0].should == XTYP_CONNECT
#        @server_calls.first[3].should == @client.topic.handle
#        @server_calls.first[4].should == @client.service.handle
#      end
#
#      it 'if server confirms connect, XTYP_CONNECT_CONFIRM transaction to service`s callback follows' do
#        @client.start_conversation 'service', 'topic'
#
#        @server_calls[1][0].should == XTYP_CONNECT_CONFIRM
#        @server_calls[1][3].should == @client.topic.handle
#        @server_calls[1][4].should == @client.service.handle
#      end
#
#      it 'client`s callback receives no transactions' do
#        @client.start_conversation 'service', 'topic'
#
#        p @server_calls, @client.service.handle, @client.topic.handle, @client.conversation
#        @client_calls.should == []
#      end
#
#      it 'fails if another conversation is already in progress' do
#        @client.start_conversation 'service', 'topic'
#
#        lambda{@client.start_conversation 'service1', 'topic1'}.
#                should raise_error /Another conversation already established/
#      end
#
#      it 'fails to start conversation on unsupported service' do
#        lambda{@client.start_conversation('not_a_service', 'topic')}.
#                should raise_error /A client`s attempt to establish a conversation has failed/
#        @client.conversation_active?.should == false
#      end
#
#    end
  end
end