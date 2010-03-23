require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')
require File.expand_path(File.dirname(__FILE__) + '/app_shared')

module DDETest
  describe DDE::App do
    it_should_behave_like "DDE App"
  end
end