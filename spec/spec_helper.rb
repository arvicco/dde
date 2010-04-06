#require 'rubygems'
#require 'spork'
#
#Spork.prefork do
#  # Loading more in this block will cause your tests to run faster. However,
#  # if you change any configuration or code from libraries loaded here, you'll
#  # need to restart spork for it take effect.
#
#end
#
#Spork.each_run do
#  # This code will be run each time you run your specs.
#
#end
#
# --- Instructions ---
# - Sort through your spec_helper file. Place as much environment loading 
#   code that you don't normally modify during development in the 
#   Spork.prefork block.
# - Place the rest under Spork.each_run block
# - Any code that is left outside of the blocks will be ran during preforking
#   and during each_run!
# - These instructions should self-destruct in 10 seconds.  If they don't,
#   feel free to delete them.

$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'spec'
require 'spec/autorun'
require 'dde'

# Customize RSpec with my own extensions
module SpecMacros

  # wrapper for it method that extracts description from example source code, such as:
  # spec { use{    function(arg1 = 4, arg2 = 'string')  }}
  def spec &block
    it description_from(caller[0]), &block # it description_from(*block.source_location), &block
    #do lambda(&block).should_not raise_error end
  end

  # reads description line from source file and drops external brackets like its{}, use{}
  # accepts as arguments either file name and line or call stack member (caller[0])
  def description_from(*args)
    case args.size
      when 1
        file, line = args.first.scan(/\A(.*?):(\d+)/).first
      when 2
        file, line = args
    end
    File.open(file) do |f|
      f.lines.to_a[line.to_i-1].gsub( /(spec.*?{)|(use.*?{)|}/, '' ).strip
    end
  end
end

Spec::Runner.configure { |config| config.extend(SpecMacros) }

module DdeTest

  include Win::Dde
#  @@monitor = Dde::Monitor.new

  TEST_IMPOSSIBLE = 'Impossible'
  TEST_STRING = "Data String"

  def use
    lambda {yield}.should_not raise_error
  end

  def any_block
    lambda {|*args| args}
  end

  def start_callback_recorder(&server_block)
    @client_calls = []
    @server_calls = []
    @client = Dde::Client.new {|*args| @client_calls << extract_values(*args); DDE_FACK}
    @server = Dde::Server.new &server_block || proc {|*args| @server_calls << extract_values(*args); DDE_FACK }
    @server.start_service('service')
  end

  def stop_callback_recorder
    @client.stop_conversation if @client.conversation_active?
    @server.stop_service if @server.service_active?
    @server.stop_dde if @server.dde_active?
    @client.stop_dde if @client.dde_active? #for some reason, need to stop @server FIRST, and @client LATER
  end

  def extract_values(*args)
    args.map do |arg|
      case arg
        when 0
          0
        when Integer
          id = @client.id if @client
          id ||= @server.id if @server
          dde_query_string(id, arg)\
              || Win::Dde.constants(false).inject(nil) {|res, const| arg == Win::Dde.const_get(const) ? res || const : res }\
              || arg
        else
          arg
      end
    end
  end
end
