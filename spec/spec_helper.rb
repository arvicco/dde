$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'spec'
require 'spec/autorun'
require 'dde'

# Customize RSpec with my own extensions
module SpecMacros

  # wrapper for it method that extracts description from example source code, such as:
  # spec { use{    function(arg1 = 4, arg2 = 'string')  }}
  def spec &block
    it description_from(caller[0]), &block # it description_from(*block.source_location), &block
    #do lambda(&block).should_not raise_error end
  end

  # reads description line from source file and drops external brackets like its{}, use{}
  # accepts as arguments either file name and line or call stack member (caller[0])
  def description_from(*args)
    case args.size
      when 1
        file, line = args.first.scan(/\A(.*?):(\d+)/).first
      when 2
        file, line = args
    end
    File.open(file) do |f|
      f.lines.to_a[line.to_i-1].gsub( /(spec.*?{)|(use.*?{)|}/, '' ).strip
    end
  end
end

Spec::Runner.configure { |config| config.extend(SpecMacros) }

module DDETest

  include Win::DDE
#  @@monitor = DDE::Monitor.new

  TEST_IMPOSSIBLE = 'Impossible'

  def use
    lambda {yield}.should_not raise_error
  end

  def any_block
    lambda {|*args| args}
  end

end