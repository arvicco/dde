module DDE

  # Class encapsulates DDE Client that requests connection with DDE server and exchanges data with it via DDE
  class Client < App

    attr_reader :conversation, # active DDE conversation that client is engaged in
                :service, #service that the client is connected to
                :topic, # active DDE conversation topic
                :item   # active DDE conversation item

#    # Creates new DDE client instance
#    def initialize(init_flags = nil, &dde_callback )
#      super init_flags, &dde_callback
#    end

    # establish a conversation with a server application that supports the specified service
    # name and topic name pair.
    def start_conversation( service=nil, topic=nil )
      try "Starting conversation #{service} #{topic}", DDE::Errors::ClientError do
        error "DDE is not initialized" unless dde_active?
        error "Another conversation already established" if conversation_active?

        # Create DDE strings for service and topic unless they are omitted
        @service = DDE::DdeString.new(@id, service) if service
        @topic = DDE::DdeString.new(@id, topic) if topic

        # Initiate new DDE conversation, returns conversation handle or nil 
        error unless @conversation = dde_connect(@id, @service.handle, @topic.handle)
      end
    end

    def stop_conversation
      try "Stopping conversation", DDE::Errors::ClientError do
        error "DDE not started" unless dde_active?
        error "Conversation not started" unless conversation_active?

        error unless dde_disconnect(@conversation) &&    # Stop DDE conversation
        dde_free_string_handle(@id, @service.handle) &&  # Free string handles for service name
        dde_free_string_handle(@id, @topic.handle)       # Free string handles for topic name

        # Unset attributes for conversation, service and topic
        @conversation = nil
        @service = nil
        @topic = nil
      end
    end

    def conversation_active?
      !!@conversation
    end

  end
end