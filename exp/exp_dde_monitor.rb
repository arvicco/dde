# Objective DDE Server (for experimentation)

require 'win/gui/message'
include Win::GUI::Message

require 'dde'
include Win::Dde

calls = []
$monitor = Dde::Monitor.new
msg = Msg.new  # pointer to Msg FFI struct

# Starting message loop (necessary for DDE processing)
puts "Starting message loop\n"
while msg = get_message()
  translate_message(msg)
  dispatch_message(msg)
end
