module DdeTest
  shared_examples_for "DDE App" do
    before(:each ){ @app = described_class.new }
    after(:each ){ @app.stop_dde if @app.dde_active?}

      it 'starts with nil id and flags if no arguments given' do
        @app.id.should == nil
        @app.init_flags.should == nil
        @app.dde_active?.should == false
      end unless described_class == Dde::Monitor

      it 'starts DDE (initializes as STANDARD DDE app) with given callback block' do
        @app = described_class.new {|*args|}
        @app.id.should be_an Integer
        @app.id.should_not == 0
        @app.init_flags.should == APPCLASS_STANDARD
        @app.dde_active?.should == true
      end unless described_class == Dde::Monitor

    describe '#start_dde' do
      it 'starts DDE with callback and default init_flags' do
        res = @app.start_dde {|*args|}
        res.should be_true
        @app.id.should be_an Integer
        @app.id.should_not == 0
        @app.init_flags.should == APPCLASS_STANDARD
        @app.dde_active?.should == true
      end unless described_class == Dde::Monitor

      it 'returns self if success (allows method chain)' do
        @app.start_dde {|*args|}.should == @app
      end

      it 'starts DDE with callback and given init_flags' do
        res = @app.start_dde( APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS ){|*args|}
        res.should be_true
        @app.id.should be_an Integer
        @app.id.should_not == 0
        @app.init_flags.should == APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS
        @app.dde_active?.should == true
      end

      it 'raises InitError if no callback was given' do
        lambda{ @app.start_dde}.should raise_error Dde::Errors::InitError
      end

      it 'reinitializes with new flags and callback if it was already initialized' do
        @app.start_dde {|*args| 1}
        old_id = @app.id
        res = @app.start_dde( APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS ){|*args| 2}
        res.should be_true
        @app.id.should == old_id
        @app.init_flags.should == APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS
        @app.dde_active?.should == true
      end
    end

    describe '#stop_dde' do
      it 'stops DDE that was active' do
        @app.start_dde {|*args| 1}

        @app.stop_dde
        @app.id.should == nil
        @app.dde_active?.should == false
      end

      it 'preserves init_flags after DDE is stopped (for reinitialization)' do
        @app.start_dde(APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS) {|*args| 1}

        @app.stop_dde
        @app.init_flags.should == APPCLASS_STANDARD | CBF_FAIL_CONNECTIONS
      end

      it 'returns self if success (allows method chain)' do
        @app.start_dde{|*args|}.stop_dde.should == @app
      end

      it 'raises InitError if dde was not active first' do
        lambda{ @app.stop_dde}.should raise_error Dde::Errors::InitError
      end
    end unless described_class == Dde::Monitor


  end
end  