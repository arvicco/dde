module DDE

  # Class encapsulates DDE Server mimicking Excel. It is used to create DDE server with specific service name
  # (default name 'excel'),
  # connect/disconnect it to/from DDEML library and store data received by the server via DDE
  class Server < App

    attr_reader :service, # service(s) that this Server supports
                :format, # data format(s) (registered clipboard formats) that server supports
                :table # data table for data storage

    # Creates new DDE application instance
    def initialize(init_flags = nil, &dde_callback )
      super init_flags, &dde_callback

      @table = DDE::XLTable.new

      # Trying to register or retrieve existing format XlTable
      unless @format = register_clipboard_format("XlTable")
        raise DDE::Errors::FormatError, "Unable to register XlTable format"
      end

      # todo: Destructor to ensure Dde instance is uninitialized and string handles freed (is it even working?)
      #ObjectSpace.define_finalizer self, ->(id) { disconnect }
    end

    def start_service( name='excel', init_flags=nil, &dde_callback )
      begin
        start_dde( init_flags, &dde_callback ) unless dde_active?

        # Create DDE string for name with handle that can be passed to DDEML functions.
        @service = DDE::DdeString.new(@id, name)

        # Register new DDE service, returns true/false success code
        raise dde_get_last_error unless dde_name_service(@id, @service.handle, DNS_REGISTER)
      rescue => e
        raise DDE::Errors::ServiceError, "Unable to start service #{name}: #{e}"
      end
      true
    end

    def stop_service
      return false unless dde_active? && service_active?

      # Unregister service by name
      return false unless dde_name_service(@id, @service.handle, DNS_UNREGISTER);

      # Free string handle for service name
      # clear handle if uninitialization successful
      @service = nil if dde_free_string_handle(@id, @service.handle)
    end

    def service_active?
      !!@service
    end

  end
end