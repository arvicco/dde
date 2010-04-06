module Dde
  extend FFI::Library  # todo: < Array ?
#  class Conv < FFI::Union
#    layout( :w, :ushort, # word should be 2 bytes, not 8
#            :d, :double, # it is 8 bytes
#            :b, [:char, 8]) # it is 8 bytes
#  end
#

  # XLTable class represents a single chunk of DDE data formatted as an Excel table
  class XlTable
    include Win::Dde

    # Received data types
    TDT_FLOAT = 1
    TDT_STRING = 2
    TDT_BOOL = 3
    TDT_ERROR = 4
    TDT_BLANK = 5
    TDT_INT = 6
    TDT_SKIP = 7
    TDT_TABLE = 16

    TDT_TYPES = {
            TDT_FLOAT => 'TDT_FLOAT',
            TDT_STRING =>'TDT_STRING',
            TDT_BOOL => 'TDT_BOOL',
            TDT_ERROR => 'TDT_ERROR',
            TDT_BLANK => 'TDT_BLANK',
            TDT_INT => 'TDT_INT',
            TDT_SKIP => 'TDT_SKIP',
            TDT_TABLE => 'TDT_TABLE'
    }

    attr_accessor :topic, # topic prefix
                  :time, # time spent parsing last transaction data
                  :total_time, # total time spent parsing data
                  :num_trans # number of data transactions

    def initialize
      @table = []    # Array contains Arrays of Strings
      @col = 0
      @row = 0
      @total_time = 0
      @total_records = 0
      # omitting separators for now
    end

    # tests if table data is empty or contains data in inconsistent state
    def empty?
      @table.empty? ||
              @row == 0 || @col == 0 ||
              @row != @table.size ||
              @col != @table.first.size  # assumes first element is also an Array
    end

    def data?;
      !empty?
    end

    def draw
      return false if empty?
      Encoding.default_external = 'cp866'
      # omitting separator gymnastics for now
      cout "-----\n"
      @table.each{|row| cout @topic; row.each {|col| cout " #{col}"}; cout "\n"}
    end

    def debug
      return false if empty?
      Encoding.default_external = 'cp866'
      # omitting separator gymnastics for now
      cout "-----\n"
      @table.each_with_index{|row, i| (cout @topic, i; p row) unless row == []}
      STDIN.gets
    end
    
    def receive(data_handle, mode = :collect)
      $mode = mode
      start = Time.now

      @offset = 0
      @pos = 0 #; @c=0; @r=0

      @data, total_size = dde_get_data(data_handle) #dde_access_data(dde_handle)
p @data.get_bytes(0, total_size) if $mode == :debug

      return nil unless @data &&  # DDE data is present at given dde_handle
      read_int == TDT_TABLE &&    # and, first data block is tdtTable
      read_int == 4               # and, its length is 4 bytes

      @row = read_int
      @col = read_int
      return nil unless @row != 0 && @col != 0  # Make sure nonzero row and col

p "data set size #{total_size}, row #{@row}, col #{@col}" if $mode == :debug
@strings = @floats = @flints = @ints = @blanks = @skips = @bools = @errors = 0

      @table = Array.new(@row){||Array.new}

      while @offset <= total_size-4   # Need at least 4 bytes ahead to read data type and size
        type = read_int         # Next data field(s) type
        size = read_int    # Next data field(s) length in bytes

p "type #{TDT_TYPES[type]}, cb #{size}, row #{@pos/@col}, col #{@pos%@col}" if $mode == :debug
        case type
          when TDT_STRING       # Strings, length byte followed by chars, no zero termination
            field_end = @offset + size
            while @offset < field_end do
              length = read_char
              self.table = @data.get_bytes(@offset, length) #read_bytes(length)#.force_encoding('CP1251').encode('CP866')
              @offset += length
              @strings += 1
            end
          when TDT_FLOAT        # Float, 8 bytes (used to represent Integers too in Quik!)
            (size/8).times do
              float_or_int = @data.get_float64(@offset)   # self.table = read_double
              @offset += 8
              int = float_or_int.round
              self.table = float_or_int == int ? (@flints += 1; int) : (@floats +=1; float_or_int)
            end
          when TDT_BLANK        # Number of blank cells, 2 bytes
            (size/2).times { read_int.times { self.table = ""; @blanks += 1 } }
          when TDT_SKIP         # Number of cells to skip, 2 bytes - in Quik, it means that these cells contain 0
            (size/2).times { read_int.times { self.table = 0; @skips += 1 } }
          when TDT_INT          # Int, 2 bytes
            (size/2).times { self.table = read_int; @ints += 1 }
          when TDT_BOOL         # Bool, 2 bytes 0/1
            (size/2).times { self.table = read_int == 0; @bools += 1 }
          when TDT_ERROR        # Error enum, 2 bytes
            (size/2).times { self.table = "Error:#{read_int}"; @errors += 1 }
          else
            cout "Type: #{type}, #{TDT_TYPES[type]}"
            return nil
        end
      end
#TODO:	free FFI::Pointer ?  delete []data;  // Free memory
      @time = Time.now - start
      @total_time += @time
      @total_records += @row
      #dde_unaccess_data(dde_handle)
      true      # Data acquisition successful
    end

    def timer
      cout "Last: #{@row} in #{@time} s(#{@time/@row} s/rec), total: #{@total_records} in  #{
              @total_time} s(#{@total_time/@total_records} s/rec)\n"
    end
    
    def formats
      cout "Strings #{@strings} Floats #{@floats} FlInts #{@flints} Ints #{@ints} Blanks #{
              @blanks} Skips #{@skips} Bools #{@bools} Errors #{@errors}\n"
    end

    def read_char
      @offset += 1
      @data.get_int8(@offset-1)
    end

    def read_int
      @offset += 2
      @data.get_int16(@offset-2)
    end

    def read_double
      @offset += 8
      @data.get_float64(@offset-8)
    end

    def read_bytes(length=1)
      @offset += length
      @data.get_bytes(@offset-length, length)
    end

    def table=(value)
#      @table[@r][@c] = value
#      todo: Add code for adding value to data row here (pack?)
#      @c+=1
#      if @c == @col
#        @c =0
#        @r+=1
#      end
#      todo: Add code for (sync!) publishing of assembled data row here (bunny? rosetta_queue?)     
      p value if $mode == :debug
      @table[@pos/@col][@pos%@col] = value
      @pos += 1
    end

  end
end
