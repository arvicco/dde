require 'dde/xl_table'

module Dde

  # Class encapsulates DDE Server mimicking Excel. It is used to create DDE server with specific service name
  # (default name 'excel') and store data received by the server via DDE
  class XlServer < Server

    attr_reader :format, # data format(s) (registered clipboard formats) that server supports
                :data    # data storage/processor

    attr_accessor :actions # Actions to be run on table after each successful DDE input (:draw, :debug, :timer)

    # Creates new Xl Server instance
    def initialize(init_flags = nil, &dde_callback )

      @data = Dde::XlTable.new

      # Trying to register or retrieve existing format XlTable
      try 'Registering format XlTable', Dde::Errors::FormatError do
        @format = register_clipboard_format("XlTable")
      end

      super init_flags, &dde_callback
    end

    #  HDDEDATA CALLBACK DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, HSZ hsz2,
    #                                      HDDEDATA hData, DWORD dwData1, DWORD dwData2)
    def default_callback
      lambda do |type, format, conv, hsz1, hsz2, data_handle, data1, data2|
        case type
          when XTYP_CONNECT  # Request to connect from client, creating data exchange channel
            #     format:: Not used.
            #     conv:: Not used.
            #     hsz1:: Handle to the topic name.
            #     hsz2:: Handle to the service name.
            #     data_handle:: Handle to DDE data. Meaning depends on the type of the current transaction.
            #     data1:: Pointer to a CONVCONTEXT structure that contains context information for the conversation.
            #             If the client is not a DDEML application, this parameter is 0.
            #     data2:: Specifies whether the client is the same application instance as the server. If the parameter
            #             is 1, the client is the same instance. If it is 0, the client is a different instance.
            #     *Returns*:: A server callback function should return TRUE(1, but DDE_FACK works just fine too)
            #                 to allow the client to establish a conversation on the specified service name and topic
            #                 name pair, or the function should return FALSE to deny the conversation. If the callback
            #                 function returns TRUE and a conversation is successfully established, the system passes
            #                 the conversation handle to the server by issuing an XTYP_CONNECT_CONFIRM transaction to
            #                 the server's callback function (unless the server specified the CBF_SKIP_CONNECT_CONFIRMS
            #                 flag in the DdeInitialize function).

            if hsz2 == @service.handle
              cout "Service #{@service}: connect requested by client\n"
              DDE_FACK # instead of true     # Yes, this server supports requested (name) handle
            else
              cout "Service #{@service} unable to process connection request for #{hsz2}\n"
              DDE_FNOTPROCESSED # 0 instead of false    # No, server does not support requested (name) handle
            end

          when XTYP_POKE  # Client initiated XTYP_POKE transaction to push unsolicited data to the server
            #     format:: Specifies the format of the data sent from the server.
            #     conv:: Handle to the conversation.
            #     hsz1:: Handle to the topic name. (Excel: [topic]item ?!)
            #     hsz2:: Handle to the item name.
            #     data_handle:: Handle to the data that the client is sending to the server.
            #     *Returns*:: A server callback function should return the DDE_FACK flag if it processes this
            #                 transaction, the DDE_FBUSY flag if it is too busy to process this transaction,
            #                 or the DDE_FNOTPROCESSED flag if it rejects this transaction.

            @data.topic = dde_query_string(@id, hsz1)  # Convert hsz1 into "[topic]item" string and
            if @data.receive(data_handle)              # Receive incoming DDE data and process it

              # Perform actions like :draw, :debug, :timer, :formats on received data (default :timer)
              @actions.each{|action| @data.send(action.to_sym)}
              DDE_FACK  # Transaction successful
            else
              @data.debug
              cout "Service #{@service} unable to process data request (XTYP_POKE) for #{hsz2}"
              DDE_FNOTPROCESSED   # 0 Transaction NOT successful - return (HDDEDATA)TRUE; ?!(why TRUE, not FALSE)
            end
          else
            DDE_FNOTPROCESSED   # 0 - return((HDDEDATA)NULL);// is it the same as 0 ?!
        end
      end
    end

    def start_service( name='excel', init_flags=nil, &dde_callback)
      super name, init_flags, &dde_callback || default_callback
    end

  end
end