module DDE

  # Class encapsulates DDE Server mimicking Excel. It is used to create DDE server with specific service name
  # (default name 'excel') and store data received by the server via DDE
  class Server < App

    attr_reader :service, # service(s) that this Server supports
                :format, # data format(s) (registered clipboard formats) that server supports
                :table # data table for data storage

    # Creates new DDE server instance
    def initialize(init_flags = nil, &dde_callback )
      super init_flags, &dde_callback

      @table = DDE::XLTable.new

      # Trying to register or retrieve existing format XlTable
      error unless @format = register_clipboard_format("XlTable")

      # todo: Destructor to ensure Dde instance is uninitialized and string handles freed (is it even working?)
      #ObjectSpace.define_finalizer self, ->(id) { disconnect }
    end

    def start_service( name='excel', init_flags=nil, &dde_callback )
      begin
        error unless dde_active? || start_dde( init_flags, &dde_callback )

        # Create DDE string for name with handle that can be passed to DDEML functions.
        @service = DDE::DdeString.new(@id, name)

        # Register new DDE service, returns true/false success code
        error unless dde_name_service(@id, @service.handle, DNS_REGISTER)
      rescue => e
        raise DDE::Errors::ServiceError, "Unable to start service #{name}: #{e}"
      end
      self
    end

    def stop_service
      begin
        error "Either DDE or service not initialized" unless dde_active? && service_active?

        # Unregister DDE service, returns true/false success code
        error unless dde_name_service(@id, @service.handle, DNS_UNREGISTER);

        # Free string handle for service name
        error unless dde_free_string_handle(@id, @service.handle)
      rescue => e
        raise DDE::Errors::ServiceError, "Unable to stop service #{@service}: #{e}"
      end
      # Clear handle if service successfuly stopped
      @service = nil
      self
    end

    def service_active?
      !!@service
    end

  end
end