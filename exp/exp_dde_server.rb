# Objective DDE Server (for experimentation)

require 'win/gui/message'
include Win::GUI::Message

require 'dde'
include Win::DDE

calls = []
$server = DDE::Server.new do |*args|
  calls << extract_values(*args)  #[Win::DDE::TYPES[args.shift]]+args; 1}
  puts "#{Time.now.strftime('%T.%6N')} #{extract_values(*args)}"
  args.first == XTYP_CONNECT ? 1 : DDE_FACK
end
sleep 0.05
$server.start_service('test_service')

def extract_values(type, format, conv, hsz1, hsz2, data, data1, data2)
  [Win::DDE::TYPES[type], format, conv,
   dde_query_string($server.id, hsz1),
   dde_query_string($server.id, hsz2),
   data, data1, data2]
end

msg = Msg.new  # pointer to Msg FFI struct

# Starting message loop (necessary for DDE processing)
puts "Starting message loop\n"
while msg = get_message()
  translate_message(msg)
  dispatch_message(msg)
end

p calls.map{|c| c.map{|e|e.respond_to?(:address) ? e.address : (Win::DDE::TYPES[e] || e)}}

p Win::DDE::ERRORS[DdeGetLastError($server.id)]