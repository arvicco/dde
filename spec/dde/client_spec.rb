require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')
require File.expand_path(File.dirname(__FILE__) + '/app_shared')

module DDETest

  def start_callback_recorder
    @client_calls = []
    @server_calls = []
    @client = DDE::Client.new {|*args| @client_calls << extract_values(*args); 1}
    @server = DDE::Server.new do |*args|
      @server_calls << extract_values(*args)  #[Win::DDE::TYPES[args.shift]]+args; 1}
      #puts "#{Time.now.strftime('%T.%6N')} #{extract_values(*args)}"
      args.first == XTYP_CONNECT ? 1 : DDE_FACK
    end
    @server.start_service('service')
  end

  def stop_callback_recorder
    @client.stop_conversation if @client.conversation_active?
    @server.stop_service if @server.service_active?
    @server.stop_dde if @server.dde_active?
  end

  def extract_values(type, format, conv, hsz1, hsz2, data, data1, data2)
    [Win::DDE::TYPES[type], format, conv,
     dde_query_string(@client.id, hsz1),
     dde_query_string(@client.id, hsz2),
     data, data1, data2]
  end

  describe DDE::Client do
    before(:each){ @client = DDE::Client.new }
    after(:each){ @client.stop_dde if @client.dde_active?}

    it_should_behave_like "DDE App"

    it 'new without parameters creates Client but does not activate DDEML' do
      @client.id.should == nil
      @client.conversation.should == nil
      @client.service.should == nil
      @client.topic.should == nil
      @client.dde_active?.should == false
      @client.conversation_active?.should == false
    end

    it 'new with attached callback block creates Client and activates DDEML' do
      @client = DDE::Client.new {|*args|}
      @client.id.should be_an Integer
      @client.id.should_not == 0
      @client.dde_active?.should == true
      @client.conversation.should == nil
      @client.conversation_active?.should == false
      @client.service.should == nil
      @client.topic.should == nil
    end

    describe '#start_conversation' do

      context 'with inactive (uninitialized) DDE:' do
        it 'fails to starts new conversation' do
          lambda{@client.start_conversation('service', 'topic')}.
                  should raise_error /DDE is not initialized/
          @client.conversation_active?.should == false
          lambda{@client.start_conversation(nil, nil)}.
                  should raise_error /DDE is not initialized/
          @client.conversation_active?.should == false
          lambda{@client.start_conversation}.
                  should raise_error /DDE is not initialized/
          @client.conversation_active?.should == false
        end
      end

      context 'with active (initialized) DDE AND existing DDE server supporting "service" topic' do
        before(:each ){start_callback_recorder}
        after(:each )do
          stop_callback_recorder
          #p @server_calls, @client_calls    # ?????????? No XTYP_DISCONNECT ? Why ?
        end

        it 'starts new conversation if DDE is already activated' do
          res = @client.start_conversation 'service', 'topic'
          res.should be_true
          @client.conversation_active?.should == true
        end

        it 'returns self if success (allows method chain)' do
          @client.start_conversation('service', 'topic').should == @client
        end

        it 'sets @conversation, @service and @topic attributes' do
          @client.start_conversation 'service', 'topic'

          @client.conversation.should be_an Integer
          @client.conversation.should_not == 0
          @client.service.should be_a DDE::DdeString
          @client.service.should == 'service'
          @client.service.name.should == 'service'
          @client.conversation.should be_an Integer
          @client.conversation.should_not == 0
        end

        it 'initiates XTYP_CONNECT transaction to service`s callback' do
          @client.start_conversation 'service', 'topic'

          @server_calls.first[0].should == 'XTYP_CONNECT'
          @server_calls.first[3].should == @client.topic
          @server_calls.first[4].should == @client.service
        end

        it 'if server confirms connect, XTYP_CONNECT_CONFIRM transaction to service`s callback follows' do
          @client.start_conversation 'service', 'topic'

          p @server_calls, @client_calls    # ?????????? No XTYP_DISCONNECT ? Why ?
          @server_calls[1][0].should == 'XTYP_CONNECT_CONFIRM'
          @server_calls[1][3].should == @client.topic
          @server_calls[1][4].should == @client.service
        end

        it 'client`s callback receives no transactions' do
          @client.start_conversation 'service', 'topic'

          p @server_calls, @client.service.handle, @client.topic.handle, @client.conversation
          @client_calls.should == []
        end

        it 'fails if another conversation is already in progress' do
          @client.start_conversation 'service', 'topic'

          lambda{@client.start_conversation 'service1', 'topic1'}.
                  should raise_error /Another conversation already established/
        end

        it 'fails to start conversation on unsupported service' do
          lambda{@client.start_conversation('not_a_service', 'topic')}.
                  should raise_error /A client`s attempt to establish a conversation has failed/
          @client.conversation_active?.should == false
        end

      end
    end

    describe '#stop_conversation' do

      context 'with inactive (uninitialized) DDE:' do
        it 'fails to stop conversation' do
          lambda{@client.stop_conversation}.
                  should raise_error /DDE not started/
          @client.conversation_active?.should == false
        end

      end # context 'with inactive (uninitialized) DDE:'

      context 'with active (initialized) DDE AND existing DDE server supporting "service" topic' do
        before(:each ){start_callback_recorder}
        after(:each )do
          stop_callback_recorder
          #p @server_calls, @client_calls    # ?????????? No XTYP_DISCONNECT ? Why ?
        end

        it 'fails to stop conversation' do
          lambda{@client.stop_conversation}.
                  should raise_error /Conversation not started/
          @client.conversation_active?.should == false
        end

        context 'conversation already started' do
          before(:each ){@client.start_conversation 'service', 'topic'}

          it 'stops conversation' do
            res = @client.stop_conversation
            res.should be_true
            @client.conversation_active?.should == false
          end

          it 'unsets @conversation, @service and @topic attributes' do
            @client.stop_conversation
            @client.conversation.should == nil
            @client.service.should == nil
            @client.topic.should == nil
          end

          it 'does not stop DDE' do
            @client.stop_conversation
            @client.dde_active?.should == true
          end

          it 'initiates XTYP_DISCONNECT transaction to service`s callback' do
            pending
            @client.stop_conversation
            p @server_calls, @client_calls    # ?????????? No XTYP_DISCONNECT ? Why ?
            @server_calls.last[0].should == 'XTYP_DISCONNECT'
          end

        end # context 'conversation already started'

      end # context 'with active (initialized) DDE AND existing DDE server supporting "service" topic'
    end
  end
end