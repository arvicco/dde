module DDE

  module Errors                             # :nodoc:
    def self.[](error_code)
      Win::DDE::ERRORS[error_code]
    end

    class InitError < RuntimeError          # :nodoc:
    end
    class FormatError < RuntimeError        # :nodoc:
    end
    class StringError < RuntimeError        # :nodoc:
    end
    class ServiceError < RuntimeError        # :nodoc:
    end
    class ClientError < RuntimeError        # :nodoc:
    end
  end

  # Class encapsulates DDE application. DDE::App serves as a base for more specific types,
  # such as DDE::Server or DDE:: Client.
  class App
    include Win::DDE

    attr_reader :id, :init_flags

    # Creates new DDE application (and starts DDE instance if dde_callback block is attached)
    def initialize( init_flags=nil, &dde_callback )
      @init_flags = init_flags

      start_dde init_flags, &dde_callback if dde_callback

    end
#    # todo: Destructor to ensure Dde instance is uninitialized and string handles freed...
#      ObjectSpace.define_finalizer( self, self.class.finalize))
#    end
#
#    # need to have class method, otherwise proc traps reference to instance (self) and the object
#    # is never garbage-collected (http://www.mikeperham.com/2010/02/24/the-trouble-with-ruby-finalizers/)
#    def self.finalize()
#      proc { stop_dde } #does NOT work since stop_dde is instance method (depends on self)
#    end

    # (Re)Initialize application with DDEML library, providing attached dde callback
    # either preserved @init_flags or init_flags argument are used
    def start_dde( init_flags=nil, &dde_callback )
      @init_flags = init_flags || @init_flags || APPCLASS_STANDARD

      try "Starting DDE" do
        @id, status = dde_initialize @id, @init_flags, &dde_callback
        error(status) unless @id && status == DMLERR_NO_ERROR
      end
    end

    # (Re)Initialize application with DDEML library, providing attached dde callback
    def stop_dde
      try "Stopping DDE" do
        error "DDE not started" unless dde_active?
        error unless dde_uninitialize(@id)   # Uninitialize app with DDEML library
        @id = nil                                 # Clear instance id if uninitialization successful
      end
    end

    # Expects a block, yields to it inside a rescue block, raises given error_type with extended fail message.
    # Returns self in case of success (to enable method chaining).
    def try( action, error_type=DDE::Errors::InitError )
      begin
        yield
      rescue => e
        raise error_type, action + " failed with: #{e}"
      end
      self
    end

    # Raises Runtime error with message based on given message (DdeGetLastError message if no message given)
    def error( message = nil )
      raise case message
        when Integer
          DDE::Errors[message]
        when nil
          DDE::Errors[dde_get_last_error(@id)]
        else
          message
      end
    end

    def dde_active?
      !!@id
    end

  end
end