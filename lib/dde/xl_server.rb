require 'dde/xl_table'

module DDE

  # Class encapsulates DDE Server mimicking Excel. It is used to create DDE server with specific service name
  # (default name 'excel') and store data received by the server via DDE
  class XlServer < Server

    attr_reader :format, # data format(s) (registered clipboard formats) that server supports
                :table # data table for data storage

    # Creates new Xl Server instance
    def initialize(init_flags = nil, &dde_callback )

      @table = DDE::XlTable.new

      # Trying to register or retrieve existing format XlTable
      try 'Registering format XlTable', DDE::Errors::FormatError do
        @format = register_clipboard_format("XlTable")
      end

      super init_flags, &dde_callback
    end

    #  HDDEDATA CALLBACK DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, HSZ hsz2,
    #                                      HDDEDATA hData, DWORD dwData1, DWORD dwData2)
    def default_callback
      lambda do |type, format, conv, hsz1, hsz2, data, data1, data2|
        case type
          when XTYP_CONNECT  # Request to connect from client, creating data exchange channel
            #             format:: Not used.
            #             conv:: Not used.
            #             hsz1:: Handle to the topic name.
            #             hsz2:: Handle to the service name.
            #             data:: Handle to DDE data. Meaning of this parameter depends on the type of the current transaction.
            #             data1:: Pointer to a CONVCONTEXT structure that contains context information for the conversation.
            #                       If the client is not a Dynamic Data Exchange Management Library (DDEML) application,
            #                       this parameter is 0.
            #             data2:: Specifies whether the client is the same application instance as the server. If the
            #                       parameter is 1, the client is the same instance. If the parameter is 0, the client
            #                       is a different instance.
            #             *Returns*:: A server callback function should return TRUE(1) to allow the client to establish a
            #                         conversation on the specified service name and topic name pair, or the function
            #                         should return FALSE to deny the conversation. If the callback function returns TRUE(1)
            #                         and a conversation is successfully established, the system passes the conversation
            #              todo:      handle to the server by issuing an XTYP_CONNECT_CONFIRM transaction to the server's
            #                         callback function (unless the server specified the CBF_SKIP_CONNECT_CONFIRMS flag
            #                         in the DdeInitialize function).

            if hsz2 == @service.handle
              cout "Service #{@service}: connect requested by client\n"
              1 # instead of true     # Yes, this server supports requested (name) handle
            else
              cout "Service #{@service} unable to process connection request for #{hsz2}\n"
              DDE_FNOTPROCESSED # 0 instead of false    # No, server does not support requested (name) handle
            end

          when XTYP_POKE  # Client initiated XTYP_POKE transaction to push unsolicited data to the server
            #             format:: Specifies the format of the data sent from the server.
            #             conv:: Handle to the conversation.
            #             hsz1:: Handle to the topic name. (Excel: [topic]item ?!)
            #             hsz2:: Handle to the item name.
            #             data_handle:: Handle to the data that the client is sending to the server.
            #             *Returns*:: A server callback function should return the DDE_FACK flag if it processes this
            #                         transaction, the DDE_FBUSY flag if it is too busy to process this transaction,
            #                         or the DDE_FNOTPROCESSED flag if it rejects this transaction.

            if @table.get_data(data) # Extract incoming DDE data from client into server's table
              # Converting hsz1 into "[topic]item" string and
              @table.topic = dde_query_string(@id, hsz1)
              # @table.draw # Simply printing it for now, no queues
              @table.timer
              #                  // Placing table into print queue
              #                  WaitForSingleObject(hMutex1,INFINITE);
              #                      q.push(server.xltable);
              #                  ReleaseMutex(hMutex1);
              #                  // Allowing the table output thread to start...
              #                  ReleaseSemaphore(hSemaphore,1,NULL);
              #
              DDE_FACK  # Transaction successful
            else
              cout "Service #{@service} unable to process data request (XTYP_POKE) for #{hsz2}"
              DDE_FNOTPROCESSED   # 0 Transaction NOT successful - return (HDDEDATA)TRUE; ?!(why TRUE, not FALSE)
            end

          when XTYP_DISCONNECT # DDE client disconnects
            #              server.xltable.Delete();
            #              break;
            DDE_FNOTPROCESSED   # 0 - return((HDDEDATA)NULL);// is it the same as 0 ?!

          when XTYP_ERROR # DDE Error
            #              WaitForSingleObject(hMutex, INFINITE);
            #                  std::cerr<<"DDE error.\n";
            #              ReleaseMutex(hMutex);
            #              break;
            DDE_FNOTPROCESSED   # 0 - return((HDDEDATA)NULL);// is it the same as 0 ?!

          else
            DDE_FNOTPROCESSED   # 0 - return((HDDEDATA)NULL);// is it the same as 0 ?!
        end
      end
    end

    # Make 'excel' the default name for named service
    alias_method :__start_service, :start_service

    def start_service( name='excel', init_flags=nil, &dde_callback)
      dde_callback ||= default_callback
      __start_service( name, init_flags, &dde_callback )
    end

  end
end