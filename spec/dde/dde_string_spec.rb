require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

module DdeTest
  describe Dde::DdeString do
    before(:each ){ @app = Dde::App.new {|*args|}}

    context ' with valid instance id of active DDE application' do
      it 'can be created from normal string' do
        dde_string = Dde::DdeString.new(@app.id, "My_String")
        dde_string == "My_String"
        dde_string.handle.should be_an Integer
        dde_string.handle.should_not == 0
      end

      it 'can be created from valid DDE string handle' do
        string_handle = dde_create_string_handle(@app.id, 'My String')
        dde_string = Dde::DdeString.new(@app.id, string_handle)
        dde_string == "My_String"
        dde_string.handle.should be_an Integer
        dde_string.handle.should_not == 0
      end
    end

    context ' without instance id of active DDE application' do
      it 'cannot be created from String' do
        lambda{Dde::DdeString.new(nil, "My_String")}.should raise_error Dde::Errors::StringError
        lambda{Dde::DdeString.new(12345, "My_String")}.should raise_error Dde::Errors::StringError
        lambda{Dde::DdeString.new(0, "My_String")}.should raise_error Dde::Errors::StringError
      end

      it 'cannot be created from valid string handle' do
        string_handle = dde_create_string_handle(@app.id, 'My String')
        lambda{Dde::DdeString.new(nil, string_handle)}.should raise_error Dde::Errors::StringError
        lambda{Dde::DdeString.new(12345, string_handle)}.should raise_error Dde::Errors::StringError
        lambda{Dde::DdeString.new(0, string_handle)}.should raise_error Dde::Errors::StringError
      end
    end

  end
end