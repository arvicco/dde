require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest
  describe DDE::DdeString do
    before(:each ){ @app = DDE::App.new {|*args|}}

    context ' with valid instance id of active DDE application' do
      it 'can be created from normal string' do
        dde_string = DDE::DdeString.new(@app.id, "My_String")
        dde_string == "My_String"
        dde_string.handle.should be_an Integer
        dde_string.handle.should_not == 0
      end

      it 'can be created from valid DDE string handle' do
        string_handle = dde_create_string_handle(@app.id, 'My String')
        dde_string = DDE::DdeString.new(@app.id, string_handle)
        dde_string == "My_String"
        dde_string.handle.should be_an Integer
        dde_string.handle.should_not == 0
      end
    end

    context ' without instance id of active DDE application' do
      it 'cannot be created from String' do
        lambda{DDE::DdeString.new(nil, "My_String")}.should raise_error DDE::Errors::StringError
        lambda{DDE::DdeString.new(12345, "My_String")}.should raise_error DDE::Errors::StringError
        lambda{DDE::DdeString.new(0, "My_String")}.should raise_error DDE::Errors::StringError
      end

      it 'cannot be created from valid string handle' do
        string_handle = dde_create_string_handle(@app.id, 'My String')
        lambda{DDE::DdeString.new(nil, string_handle)}.should raise_error DDE::Errors::StringError
        lambda{DDE::DdeString.new(12345, string_handle)}.should raise_error DDE::Errors::StringError
        lambda{DDE::DdeString.new(0, string_handle)}.should raise_error DDE::Errors::StringError
      end
    end


#    it 'starts with nil id and flags if no arguments given' do
#      @app.id.should == nil
#      @app.init_flags.should == nil
#    end
#
#    it 'starts DDE (initializes as STANDARD DDE app) with given callback block' do
#      app = DDE::App.new {|*args|}
#      app.id.should be_an Integer
#      app.id.should_not == 0
#      app.init_flags.should == APPCLASS_STANDARD
#    end
#
#    describe '#start_dde' do
#      it 'starts DDE with callback and default init_flags' do
#        res = @app.start_dde {|*args|}
#        res.should == true
#        @app.id.should be_an Integer
#        @app.id.should_not == 0
#        @app.init_flags.should == APPCLASS_STANDARD
#      end
#
#      it 'starts DDE with callback and given init_flags' do
#        res = @app.start_dde( APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS ){|*args|}
#        res.should == true
#        @app.id.should be_an Integer
#        @app.id.should_not == 0
#        @app.init_flags.should == APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS
#      end
#
#      it 'raises InitError if no callback was given' do
#        lambda{ @app.start_dde}.should raise_error DDE::Errors::InitError
#      end
#
#      it 'reinitializes with new flags and callback if it was already initialized' do
#        @app.start_dde {|*args| 1}
#        old_id = @app.id
#        res = @app.start_dde( APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS ){|*args| 2}
#        res.should == true
#        @app.id.should == old_id
#        @app.init_flags.should == APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS
#      end
#    end
#
#    describe '#stop_dde' do
#      it 'raises InitError if dde was not started first' do
#        lambda{ @app.stop_dde}.should raise_error DDE::Errors::InitError
#      end
#
#    end


  end
end