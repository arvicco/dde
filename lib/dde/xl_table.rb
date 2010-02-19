module DDE
  extend FFI::Library  # todo: < Array ?
  class Conv < FFI::Union
    layout( :w, :ushort, # word should be 2 bytes, not 8
            :d, :double, # it is 8 bytes
            :b, [:char, 8]) # it is 8 bytes
  end


  # XLTable class represents a single chunk of DDE data formatted as an Excel table
  class XLTable
    include Win::DDE

    # Received data types
    TDT_FLOAT = 1
    TDT_STRING = 2
    TDT_BOOL = 3
    TDT_ERROR = 4
    TDT_BLANK = 5
    TDT_INT = 6
    TDT_SKIP = 7
    TDT_TABLE = 16


    attr_accessor :buf # topic_item

    def initialize
      @table_data = []    # Array contains Arrays of Strings
      @col = 0
      @row = 0
      # omitting separators for now
    end

    # tests if table data is empty or contains data in inconsistent state
    def empty?
      @table_data.empty? ||
              @row == 0 || @col == 0 ||
              @row != @table_data.size ||
              @col != @table_data.first.size  # assumes first element is also an Array
    end

    def data?;
      !empty?
    end

    def draw
      return false if empty?

      # omitting separator gymnastics for now
      cout "-----\n"
      @table_data.each{|row| row.each {|col| cout col, " "}; cout "\n"}
    end

    def get_data(dde_handle)
      conv = DDE::Conv.new  # Union for data conversion

      # Copy DDE data from dde_handle (FFI::MemoryPointer is returned)
      return nil unless data = dde_get_data(dde_handle) # raise 'DDE data not extracted'
      offset = 0

      # Make sure that the first block is tdtTable
      return nil unless data.get_int16(offset) == TDT_TABLE # raise 'DDE data not TDT_TABLE'
      offset += 2

      # Make sure cb == 4
      return nil unless data.get_int16(offset) == 4 # raise 'TDT_TABLE data length wrong'
      offset += 2

      row = data.get_int16(offset)
      col = data.get_int16(offset+2)

      @table_data = Array.new(row, [])
      # Make sure nonzero row and col
      return nil if row == 0 || col == 0   # raise 'col or row zero in TDT_TABLE'
      offset += 4

      r = 0
      c = 0
      while offset < data.size
        type = data.get_int16(offset)   # Next data field(s) type
        cb = data.get_int16(offset)     # Next data field(s) length in bytes
        offset += 4

        case type
          when TDT_FLOAT       # Float, 8 bytes per field
            (cb/8).times do
              @table_data[r][c] = data.get_float64(offset) # TODO: check if data.get_double(offset) even exists  ???
              offset += 8
              c += 1
              if c == col # end of row
                c = 0
                r += 1
              end
            end
          when TDT_STRING
            end_field = offset + cb
            while offset < end_field do
              length = data.get_int16(offset)
              offset += 2
              @table_data[r][c] = data.get_bytes(offset, length)
              offset += length
              c += 1
              if c == col # end of row
                c = 0
                r += 1
              end
            end
          when TDT_BOOL        # Bool, 2 bytes per field
            (cb/2).times do
              @table_data[r][c] = data.get_int16(offset) == 0
              offset += 2
              c += 1
              if c == col # end of row
                c = 0
                r += 1
              end
            end
          when TDT_ERROR        # Error enum, 2 bytes per field
            (cb/2).times do
              @table_data[r][c] = "Error:#{data.get_int16(offset)}"
              offset += 2
              c += 1
              if c == col # end of row
                c = 0
                r += 1
              end
            end
          when TDT_BLANK        # Number of blank cells, 2 bytes per field
            (cb/2).times do
              blanks = data.get_int16(offset)
              offset += 2
              blanks.times do
                @table_data[r][c] = ""
                c += 1
                if c == col # end of row
                  c = 0
                  r += 1
                end
              end
            end
          when TDT_INT        # Int, 2 bytes per field
            (cb/2).times do
              @table_data[r][c] = data.get_int16(offset) == 0
              offset += 2
              c += 1
              if c == col # end of row
                c = 0
                r += 1
              end
            end
          else
            return nil
        end
      end
#TODO:	free FFI::Pointer ?  delete []data;                          // Возвращаем память системе
	  true      # Data aquisition successful
    end
  end
end
