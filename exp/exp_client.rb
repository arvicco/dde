# Quick and dirty DDE Client (for experimentation)

require 'win/gui/message'
include Win::GUI::Message

#require_relative 'exp_lib'
#include DDELib

require 'win/dde'
include Win::Dde

calls = []
buffer = FFI::MemoryPointer.new(:long).write_long(0)
buffer.address

callback = lambda do |*args|
  calls << [*args]
  DDE_FACK
end

p status = DdeInitialize(buffer, callback, APPCLASS_STANDARD, 0)
p id = buffer.read_long

service = FFI::MemoryPointer.from_string('test_service')

p handle = DdeCreateStringHandle(id, service, CP_WINANSI)

p conv_handle = DdeConnect(id, handle, handle, nil)

str = FFI::MemoryPointer.from_string("Poke_string\n\x00\x00")

p DdeClientTransaction(str, str.size, conv_handle, handle, CF_TEXT, XTYP_POKE, 1000, nil)
p Win::Dde::ERRORS[DdeGetLastError(id)]
p DdeClientTransaction(str, str.size, conv_handle, handle, CF_TEXT, XTYP_EXECUTE, 1000, nil)
p Win::Dde::ERRORS[DdeGetLastError(id)]
sleep 0.01
p DdeClientTransaction(str, str.size, conv_handle, handle, CF_TEXT, XTYP_EXECUTE, TIMEOUT_ASYNC, nil)
p Win::Dde::ERRORS[DdeGetLastError(id)]

p DdeDisconnect(conv_handle)

p calls.map{|c| c.map{|e|e.respond_to?(:address) ? e.address : (Win::Dde::TYPES[e] || e)}}

p Win::Dde::ERRORS[DdeGetLastError(id)]