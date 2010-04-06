module Dde

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

    # Establish a conversation with a server application that supports the specified service
    # name and topic name pair.
    def start_conversation( service=nil, topic=nil )
      try "Starting conversation #{service} #{topic}", Dde::Errors::ClientError do
        error "DDE is not initialized" unless dde_active?
        error "Another conversation already established" if conversation_active?

        # Create DDE strings for service and topic unless they are omitted
        @service = Dde::DdeString.new(@id, service) if service
        @topic = Dde::DdeString.new(@id, topic) if topic

        # Initiate new DDE conversation, returns conversation handle or nil 
        error unless @conversation = dde_connect(@id, @service.handle, @topic.handle)
      end
    end

    # Stops active conversation, raises error if no conversations active
    def stop_conversation
      try "Stopping conversation", Dde::Errors::ClientError do
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

    # Sends XTYP_POKE transaction to server if conversation was already established.
    # data:: data being sent (will be coerced to String unless is already a (packed) String)
    # format:: standard clipboard format of submitted data item (default CF_TEXT)
    def send_data( data, format = CF_TEXT, item = "" )
      data_pointer = FFI::MemoryPointer.from_string(data.to_s)
      result, trans_id = start_transaction(XTYP_POKE, data_pointer, data_pointer.size, format, item)
      result
    end

    # Initiates transaction to server if conversation was already established.
    # transaction_type:: XTYP_ADVSTART, XTYP_ADVSTOP, XTYP_EXECUTE, XTYP_POKE, XTYP_REQUEST
    # data_pointer:: pointer to data being sent (either FFI::MemoryPointer or DDE data_handle)
    # cb:: data set size (or -1 to indicate that data_pointer is in fact DDE data_handle)
    # format:: standard clipboard format of submitted data item (default CF_TEXT)
    # item:: item to which transaction is related (String, DdeString or DDE string handle)
    # timeout:: timeout in milliseconds or TIMEOUT_ASYNC to indicate async transaction
    #
    # *Returns*:: A pair of [result, trans_id]. Result is nil for failed transactions,
    # DDE data handle for synchronous transactions in which the client expects data from the server,
    # nonzero for successful transactions where clients does not expect data from server.
    # Trans_id: for asynchronous transactions, a unique transaction identifier for use with the
    # DdeAbandonTransaction function and the XTYP_XACT_COMPLETE transaction. For synchronous transactions,
    # the low-order word of this variable contains any applicable DDE_ flags resulting from the transaction.
    #
    def start_transaction( transaction_type, data_pointer=nil, cb = data_pointer ? data_pointer.size : 0,
            format=CF_TEXT, item=0, timeout=1000)

      result = nil
      trans_id = FFI::MemoryPointer.new(:uint32).put_uint32(0,0)

      try "Sending data to server", Dde::Errors::ClientError do
        error "DDE not started" unless dde_active?
        error "Conversation not started" unless conversation_active?

        item_handle = case item
          when String
            Dde::DdeString.new(@id, service).handle
          when DdeString
            item.handle
          else
            item
        end

        error unless result = dde_client_transaction(data_pointer, cb, @conversation, item_handle,
                                                     format, transaction_type, timeout, trans_id)
      end
      [result, trans_id.get_uint32(0)]
    end

    def conversation_active?
      !!@conversation
    end

  end
end