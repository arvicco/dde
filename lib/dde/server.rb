module Dde

  # Class encapsulates DDE Server with basic functionality (starting/stopping named service)
  class Server < App

    attr_reader :service # service(s) that this Server supports

    def start_service( name, init_flags=nil, &dde_callback )
      try "Starting service #{name}", Dde::Errors::ServiceError do
        # Trying to start DDE if it was inactive
        error unless dde_active? || start_dde( init_flags, &dde_callback )

        # Create DDE string for name (this creates handle that can be passed to DDEML functions)
        @service = Dde::DdeString.new(@id, name)

        # Register new DDE service, returns true/false success code
        error unless dde_name_service(@id, @service.handle, DNS_REGISTER)
      end
    end

    def stop_service
      try "Stopping active service", Dde::Errors::ServiceError do
        error "Either DDE or service not initialized" unless dde_active? && service_active?

        # Unregister DDE service, returns true/false success code
        error unless dde_name_service(@id, @service.handle, DNS_UNREGISTER);

        # Free string handle for service name
        error unless dde_free_string_handle(@id, @service.handle)

        # Clear handle if service successfuly stopped
        @service = nil
      end
    end

    def service_active?
      !!@service
    end

  end
end