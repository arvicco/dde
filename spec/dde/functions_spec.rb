require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DDETest
  include DDE::Functions

  describe DDE, ' contains a set of pre-defined Windows API functions' do
    describe 'register_clipboard_format' do
      spec{ use{ RegisterClipboardFormat(format_name = "XlTable") }}
      spec{ use{ register_clipboard_format(format_name = "XlTable") }}

      it 'returns format id (int) if successfully registered format' do
        id = register_clipboard_format("XlTable")
        id.should_not == 0
        id.should_not == nil
      end

      it 'returns same format id for already registered format' do
        id1 = register_clipboard_format("XlTable")
        id2 = register_clipboard_format("XlTable")
        id1.should == id2
      end

      it 'returns nil if not able to register format' do
        register_clipboard_format("").should == nil
      end
    end

    describe 'dde_initialize' do
      spec{ use{ status = DdeInitialize( id = [0].pack('L'), dde_callback, dde_cmd, unused = 0)}}
      spec{ use{ id, status = dde_initialize( id = 0, dde_cmd) do|*args| end }}

      it 'returns integer id and 0 if initialization successful' do
        id, status = dde_initialize(APPCLASS_STANDARD, dde_callback) {|*args| }
        id.should be_an Integer
        id.should_not == 0
        status.should == DMLERR_NO_ERROR
      end

      it 'returns error status if initialization unsuccessful' do
        id, status = dde_initialize(1, APPCLASS_STANDARD) {|*args| }
        p id, status
        status.should == DMLERR_INVALIDPARAMETER
        id.should == nil
      end

      it 'returns error status if initialization unsuccessful' do
        id, status = dde_initialize(APPCLASS_STANDARD) {|*args| }
        id, status = dde_initialize(id, APPCLASS_STANDARD) {|*args| }
        p id, status
        status.should == DMLERR_INVALIDPARAMETER
        id.should == nil
      end

      it 'returns nil if not able to initialize' do
        pending
        register_clipboard_format("").should == nil
      end
    end
  end
end
