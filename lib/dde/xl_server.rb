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

    # Make 'excel' the default name for named service
    alias_method :__start_service, :start_service
    def start_service( name='excel', init_flags=nil, &dde_callback )
      __start_service( name, init_flags, &dde_callback )
    end

  end
end