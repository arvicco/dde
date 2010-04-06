require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')
require File.expand_path(File.dirname(__FILE__) + '/app_shared')

module DdeTest

  describe Dde::Client do
    before(:each){ @client = Dde::Client.new }
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
      @client = Dde::Client.new {|*args|}
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
      end  # context 'with inactive (uninitialized) DDE:'

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
          @client.service.should be_a Dde::DdeString
          @client.service.should == 'service'
          @client.service.name.should == 'service'
          @client.conversation.should be_an Integer
          @client.conversation.should_not == 0
        end

        it 'initiates XTYP_CONNECT transaction to service`s callback' do
          @client.start_conversation 'service', 'topic'

          @server_calls.first[0].should == :XTYP_CONNECT
          @server_calls.first[3].should == @client.topic
          @server_calls.first[4].should == @client.service
        end

        it 'if server confirms connect, XTYP_CONNECT_CONFIRM transaction to service`s callback follows' do
          @client.start_conversation 'service', 'topic'

          # p @server_calls, @client_calls    # ?????????? No XTYP_DISCONNECT ? Why ?
          @server_calls[1][0].should == :XTYP_CONNECT_CONFIRM
          @server_calls[1][3].should == @client.topic
          @server_calls[1][4].should == @client.service
        end

        it 'client`s callback receives no transactions' do
          @client.start_conversation 'service', 'topic'

          #p @server_calls, @client.service.handle, @client.topic.handle, @client.conversation
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

          #ensure
          dde_free_string_handle(@client.id, @client.topic.handle)
          dde_free_string_handle(@client.id, @client.service.handle)
        end

      end # context 'with active (initialized) DDE AND existing DDE server supporting "service" topic' do
    end # describe '#start_conversation'

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
        after(:each ){stop_callback_recorder}

        it 'fails to stop conversation' do
          lambda{@client.stop_conversation}.
                  should raise_error /Conversation not started/
          @client.conversation_active?.should == false
        end

        context 'conversation already started' do
          before(:each){ @client.start_conversation 'service', 'topic' }
          after(:each){ @client.stop_conversation if @client.conversation_active?}

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
            @server_calls.last[0].should == :XTYP_DISCONNECT
          end

        end # context 'conversation already started'

      end # context 'with active (initialized) DDE AND existing DDE server supporting "service" topic'
    end # describe '#stop_conversation'

    describe '#send_data' do
      context 'with active (initialized) DDE AND existing DDE server supporting "service" topic' do
        before(:each )do
          start_callback_recorder do |*args|
            @server_calls << extract_values(*args)
            if args[0] == XTYP_POKE
              @data, @size = dde_get_data(args[5])
            end
            DDE_FACK
          end
          @client.start_conversation 'service', 'topic'
        end
        after(:each ){stop_callback_recorder}

        it 'sends data to server' do
          @client.send_data TEST_STRING, CF_TEXT, "item"
          @server_calls.last[0].should == :XTYP_POKE
          @data.get_bytes(0, @size).rstrip.should == TEST_STRING
        end

      end # context 'with active (initialized) DDE'
    end # describe #send_data

  end # describe DDE::Client
end