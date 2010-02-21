module DDE

  # Class encapsulates DDE Client that requests connection with DDE server and exchanges data with it via DDE
  class Client < App

    attr_reader :conversation, # active DDE conversation that client is engaged in
                :service, #service that the client is connected to
                :topic, # active DDE conversation topic
                :item   # active DDE conversation item

    # Creates new DDE client instance
    def initialize(init_flags = nil, &dde_callback )
      super init_flags, &dde_callback
    end

    # establish a conversation with a server application that supports the specified service
    # name and topic name pair.
    def start_conversation( service=nil, topic=nil )
      begin
        error "DDE is not initialized" unless dde_active?
        error "Another conversation already established" if conversation_active?

        # Create DDE strings for service and topic unless they are omitted
        @service = DDE::DdeString.new(@id, service) if service
        @topic = DDE::DdeString.new(@id, topic) if topic

        # Initiate new DDE conversation, returns conversation handle or nil 
        error unless @conversation = dde_connect(@id, @service.handle, @topic.handle)
      rescue => e
        raise DDE::Errors::ClientError, "Unable to start conversation #{service} #{topic}: #{e}"
      end
      self
    end

    def stop_conversation
      begin
        error "DDE not started" unless dde_active?
        error "Conversation not started" unless conversation_active?

        # Stop DDE conversation
        error unless dde_disconnect(@conversation)

        # Free string handles for service and topic names
        error unless dde_free_string_handle(@id, @service.handle)
        error unless dde_free_string_handle(@id, @service.handle)

      rescue => e
        raise DDE::Errors::ClientError, "Unable to stop conversation: #{e}"
      end
      # Unset attributes for conversation, service and topic
      @conversation = nil
      @service = nil
      @topic = nil
      self
    end

    def conversation_active?
      !!@conversation
    end

  end
end