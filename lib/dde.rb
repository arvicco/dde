# console output redirection (may need to wrap it in synchronization code, etc)
def cout *args
  print *args
end

require 'win/dde'
require 'dde/dde_string'
require 'dde/app'
require 'dde/xl_table'
require 'dde/server'

server = DDE::Server.new  # create server with default name 'excel'

# std::queue<XlTable> q;				// очередь, содержит таблицы, которые необходимо вывести на консоль
#  XlTable xlt;						// содержит выводимую таблицу

#  HDDEDATA CALLBACK DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, HSZ hsz2,
#                                      HDDEDATA hData, DWORD dwData1, DWORD dwData2)
dde_callback = lambda do |type, format, conv, hsz1, hsz2, data_handle, data1, data2|
  case type
    when XTYP_CONNECT  # Request to connect from client, creating data exchange channel
      #             hsz1:: Handle to the topic name.
      #             hsz2:: Handle to the service name.
      #             dwData1:: Pointer to a CONVCONTEXT structure that contains context information for the conversation.
      #                       If the client is not a Dynamic Data Exchange Management Library (DDEML) application,
      #                       this parameter is 0.
      #             dwData2:: Specifies whether the client is the same application instance as the server. If the
      #                       parameter is 1, the client is the same instance. If the parameter is 0, the client
      #                       is a different instance.
      #             *Returns*:: A server callback function should return TRUE to allow the client to establish a
      #                         conversation on the specified service name and topic name pair, or the function
      #                         should return FALSE to deny the conversation. If the callback function returns TRUE
      #                         and a conversation is successfully established, the system passes the conversation
      #              todo:      handle to the server by issuing an XTYP_CONNECT_CONFIRM transaction to the server's
      #                         callback function (unless the server specified the CBF_SKIP_CONNECT_CONFIRMS flag
      #                         in the DdeInitialize function).
      if hsz2 == server.service
        true     # Yes, this server supports requested (name) handle
      else
        cout "Unable to process connection request for #{hsz2}, service is #{server.service}\n"
        false    # No, server does not support requested (name) handle
      end
    when XTYP_POKE  # Client initiated XTYP_POKE transaction to push unsolicited data to the server
      #             format:: Specifies the format of the data sent from the server.
      #             conv:: Handle to the conversation.
      #             hsz1:: Handle to the topic name.
      #             hsz2:: Handle to the item name.
      #             data_handle:: Handle to the data that the client is sending to the server.
      #             *Returns*:: A server callback function should return the DDE_FACK flag if it processes this
      #                         transaction, the DDE_FBUSY flag if it is too busy to process this transaction,
      #                         or the DDE_FNOTPROCESSED flag if it rejects this transaction.
      #
      #              # CHAR buf[200];
      flag = server.table.get_data(data_handle) # extract client's DDE data into server's xltable
      if flag
                   # // Преобразуем HSZ в string (название топика и итема в формате [topic]item)
    #                  DdeQueryStringA(server.GetId(),hsz1,buf,200,CP_WINANSI);
    #                  server.xltable.GetBuf(buf);
    #                  // Помещаем таблицу в очередь
    #                  WaitForSingleObject(hMutex1,INFINITE);
    #                      q.push(server.xltable);
    #                  ReleaseMutex(hMutex1);
    #                  // Позволяем запустится потоку вывода таблицы на экран
    #                  ReleaseSemaphore(hSemaphore,1,NULL);
    #
        DDE_FACK  # Transaction successful
      else
        cout "Unable to receive dataprocess connection request for #{hsz2}, server handle is #{server.handle}\n"
        DDE_FNOTPROCESSED   # Transaction NOT successful
      end
        #              }
    #              return (HDDEDATA)TRUE;			 // Признак неудачной транзакции
    #          }
    #      case XTYP_DISCONNECT:		// Отключение DDE-клиента
    #          {
    #              server.xltable.Delete();
    #              break;
    #          }
    #      case XTYP_ERROR:			// Ошибка
    #          {
    #              WaitForSingleObject(hMutex, INFINITE);
    #                  std::cerr<<"DDE error.\n";
    #              ReleaseMutex(hMutex);
    #              break;
    #          }
    #
    #      //----------------------------------------------------------------------
    #      default:
    #          return (HDDEDATA)NULL;
    #      //----------------------------------------------------------------------
    #      }
    #      return((HDDEDATA)NULL);
    #  }
  end
end