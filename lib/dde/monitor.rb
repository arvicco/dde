module DDE

  # Class encapsulates DDE Monitor that prints all DDE transactions to console
  class Monitor < App

    attr_accessor :print, :calls

    # Creates new DDE monitor instance
    def initialize(init_flags=nil, print = nil, &callback)
      init_flags ||=
              APPCLASS_MONITOR |  # this is monitor
              MF_CALLBACKS     |  # monitor callback functions
              MF_CONV          |  # monitor conversation data
              MF_ERRORS        |  # monitor DDEML errors
              MF_HSZ_INFO      |  # monitor data handle activity
              MF_LINKS         |  # monitor advise loops
              MF_POSTMSGS      |  # monitor posted DDE messages
              MF_SENDMSGS         # monitor sent DDE messages

      @print = print
      @calls = []

      callback ||= lambda do |*args|
        time = Time.now.strftime('%T.%6N')
        values = extract_values(*args)
        @calls << [time, values]
        puts "#{time} #{values}" if @print
        DDE_FACK
      end

      super init_flags, &callback
    end

    def extract_values(*args)
      values = args.map {|arg| interprete_value(arg)}

      # if this is a MONITOR transaction, extract hdata using the DdeAccessData
      if values.first == :XTYP_MONITOR
        data_type = case values.last
          when :MF_CALLBACKS
            MonCbStruct #.new(dde_get_data(args[5]).first)
          # cb:: Specifies the structure's size, in bytes.
          # dwTime:: Specifies the Windows time at which the transaction occurred. Windows time is the number of
          #          milliseconds that have elapsed since the system was booted.
          # hTask:: Handle to the task (app instance) containing the DDE callback function that received the transaction.
          # dwRet:: Specifies the value returned by the DDE callback function that processed the transaction.
          # wType:: Specifies the transaction type.
          # wFmt:: Specifies the format of the data exchanged (if any) during the transaction.
          # hConv:: Handle to the conversation in which the transaction took place.
          # hsz1:: Handle to a string.
          # hsz2:: Handle to a string.
          # hData:: Handle to the data exchanged (if any) during the transaction.
          # dwData1:: Specifies additional data.
          # dwData2:: Specifies additional data.
          # cc:: Specifies a CONVCONTEXT structure containing language information used to share data in different languages.
          # cbData:: Specifies the amount, in bytes, of data being passed with the transaction. This value can be
          #          more than 32 bytes.
          # Data:: Contains the first 32 bytes of data being passed with the transaction (8 * sizeof(DWORD)).

          when :MF_CONV
            MonConvStruct
          when :MF_ERRORS
            MonErrStruct
          when :MF_HSZ_INFO
            MonHszStruct
          when :MF_LINKS
            MonLinksStruct
          else
            MonMsgStruct
        end

        #casting DDE data pointer into appropriate struct type
        struct_pointer, size = dde_get_data(args[5])
        data = data_type.new(struct_pointer)

        values = [values.first, values.last] + data.members.map do |member|
          value = data[member] rescue 'plonk'
          "#{member}: #{interprete_value(value)}"
        end
      end

      values
    end

    def interprete_value(arg)
      return arg unless arg.kind_of? Fixnum rescue return 'plAnk'
      return 0 if arg == 0
      #Trying to interpete arg as a DDE string
      dde_query_string(@id, arg)\
          || Win::DDE.constants(false).inject(nil) {|res, const| arg == Win::DDE.const_get(const) ? res || const : res }\
          || arg
    end
  end
end