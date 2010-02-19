module DDE

  module Errors                             # :nodoc:
    class InitError < RuntimeError          # :nodoc:
    end
    class FormatError < RuntimeError        # :nodoc:
    end
    class StringError < RuntimeError        # :nodoc:
    end
    class ServiceError < RuntimeError        # :nodoc:
    end

  end

  # Class encapsulates DDE application. DDE::App serves as a base for more specific types,
  # such as DDE::Server or DDE:: Client.
  class App
    include Win::DDE

    attr_reader :id, :init_flags

    # Creates new DDE application and starts DDE instance
    # if dde_callback block is attached
    def initialize( init_flags=nil, &dde_callback )
      @init_flags = init_flags

      start_dde init_flags, &dde_callback if dde_callback
    end

    # (Re)Initialize application with DDEML library, providing attached dde callback
    # either preserved @init_flags or init_flags argument are used
    def start_dde( init_flags=nil, &dde_callback )
      @init_flags = init_flags || @init_flags || APPCLASS_STANDARD

      begin
        @id, status = dde_initialize @id, @init_flags, &dde_callback
      rescue => e
        status = e
      end
      raise DDE::Errors::InitError, "DdeInitialize failed with: #{status}" unless @id && status == DMLERR_NO_ERROR
      true
    end

    # (Re)Initialize application with DDEML library, providing attached dde callback
    def stop_dde
      # Uninitialize app with DDEML library and clear instance id if uninitialization successful
      raise DDE::Errors::InitError, "DdeUninitialize failed" unless @id && dde_uninitialize(@id)
      @id = nil
      true
    end

    def dde_active?
      !!@id
    end

  end
end