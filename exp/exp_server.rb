# Quick and dirty DDE Server (for experimentation)

require 'win/gui/message'
include Win::GUI::Message

#require_relative 'exp_lib'
#include DDELib

require 'win/dde'
include Win::DDE

calls = []
buffer = FFI::MemoryPointer.new(:long).write_long(0)
buffer.address

callback = lambda do |*args|
  calls << [*args]
  puts "#{Time.now.strftime('%T.%6N')} #{args.map{|e|e.respond_to?(:address) ? e.address : (Win::DDE::TYPES[e] || e)}}"
  args.first == XTYP_CONNECT ? 1 : DDE_FACK
end

status = DdeInitialize(buffer, callback, APPCLASS_STANDARD, 0)
id = buffer.read_long

service = FFI::MemoryPointer.from_string('test_service')

p handle = DdeCreateStringHandle(id, service, CP_WINANSI)

p DdeNameService(id, handle, 0, DNS_REGISTER)

#p DdeDisconnect(conv_handle)

msg = Msg.new  # pointer to Msg FFI struct

# Starting message loop (necessary for DDE processing)
puts "Starting message loop\n"
while msg = get_message()
  translate_message(msg)
  dispatch_message(msg)
end

p calls.map{|c| c.map{|e|e.respond_to?(:address) ? e.address : (Win::DDE::TYPES[e] || e)}}

p Win::DDE::ERRORS[DdeGetLastError(id)]