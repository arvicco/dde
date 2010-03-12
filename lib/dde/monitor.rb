module DDE

  # Class encapsulates DDE Monitor that prints all DDE transactions to console
  class Monitor < App

    # Creates new DDE monitor instance
    def initialize(init_flags=nil, &callback)
      init_flags ||=
              APPCLASS_MONITOR |  # this is monitor
              MF_CALLBACKS     |  # monitor callback functions
              MF_CONV          |  # monitor conversation data
              MF_ERRORS        |  # monitor DDEML errors
              MF_HSZ_INFO      |  # monitor data handle activity
              MF_LINKS         |  # monitor advise loops
              MF_POSTMSGS      |  # monitor posted DDE messages
              MF_SENDMSGS         # monitor sent DDE messages

      callback ||= lambda do |*args|
        puts "#{Time.now.strftime('%T.%6N')} #{extract_values(*args)}"
        1
      end

      super init_flags, &callback
    end

    def extract_values(*args)
      values = []
      args.each do |arg|
        #Zero if zero arg
        value = 0 if arg == 0

        #Trying to interpete arg as a DDE string
        value ||= dde_query_string(@id, arg)

        #Trying to interpete arg as Win::DDE constant
        value ||= Win::DDE.constants(false).inject(nil) do |res, const|
          arg == Win::DDE.const_get(const) ? const : res
        end

        values << (value || arg)
      end
      # if this is a MONITOR transaction, extract hdata using the DdeAccessData
      if values.first == :XTYP_MONITOR
        data_type = case values.last
          when :MF_CALLBACKS
            MonCbStruct
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
        data = data_type.new(dde_get_data(args[5]))

        values = [values.first, values.last] + data.members
      end

      values
    end
  end
end