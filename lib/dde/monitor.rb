module DDE

  # Class encapsulates DDE Monitor that prints all DDE transactions to console
  class Monitor < App

    # Creates new DDE monitor instance
    def initialize(init_flags=nil, &callback)
      init_flags ||=
              APPCLASS_MONITOR |  # this is monitor
              MF_CALLBACKS     |  # monitor callback functions
              MF_CONV          |  # monitor conversation data
              MF_ERRORS        |  # monitor DDEML errors
              MF_HSZ_INFO      |  # monitor data handle activity
              MF_LINKS         |  # monitor advise loops
              MF_POSTMSGS      |  # monitor posted DDE messages
              MF_SENDMSGS         # monitor sent DDE messages

      callback ||= lambda do |*args|
        p args.unshift(Win::DDE::TYPES[args.shift]).push(Win::DDE::FLAGS[args.pop])
        1
      end
      
      super init_flags, &callback
    end
  end
end