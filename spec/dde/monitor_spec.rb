require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')
require File.expand_path(File.dirname(__FILE__) + '/app_shared')

module DDETest

  describe DDE::Monitor, " in general" do
    it_should_behave_like "DDE App"
  end

  describe DDE::Monitor do
    before(:each){ }
    after(:each){ @monitor.stop_dde if @monitor.dde_active? }
    # SEEMS LIKE IT DOESN'T stop system from sending :XTYP_MONITOR transactions to already dead callback :(


    it 'starts without constructor parameters' do
      @monitor = DDE::Monitor.new

      @monitor.id.should be_an Integer
      @monitor.id.should_not == 0
      @monitor.dde_active?.should == true

      @monitor.init_flags.should == APPCLASS_MONITOR |  # this is monitor
      MF_CALLBACKS     |  # monitor callback functions
      MF_CONV          |  # monitor conversation data
      MF_ERRORS        |  # monitor DDEML errors
      MF_HSZ_INFO      |  # monitor data handle activity
      MF_LINKS         |  # monitor advise loops
      MF_POSTMSGS      |  # monitor posted DDE messages
      MF_SENDMSGS         # monitor sent DDE messages
    end

  end
end