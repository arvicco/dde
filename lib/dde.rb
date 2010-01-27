require 'win_gui'

module DDE
  # Define constants:
  CP_WINANSI = 1004
  APPCLASS_STANDARD = 0
  DNS_REGISTER = 1
  DNS_UNREGISTER = 2

  XTYP_CONNECT = 0x60
  XTYP_DISCONNECT = 0xC0
  XTYP_POKE = 0x90
  XTYP_ERROR = 0x00

  DMLERR_NO_ERROR = 0x00
  DMLERR_DLL_USAGE = 0x4004
  DMLERR_INVALIDPARAMETER = 0x4006
  DMLERR_SYS_ERROR = 0x400f

  class Server

    attr_accessor :name, :id
    
    def initialize(name = 'excel')
      @name = name
      @id = 0
    end

    def connect &callback

    return false unless register_clipboard_format_a("XlTable")

	# Инициализация DDEML библиотеки (returns 0 if everything OK)
    @id, status = dde_initialize(APPCLASS_STANDARD, 0) &callback
	return false unless status == DMLERR_NO_ERROR

	# Получение идентификаторов строк для сервиса, раздела и элемента данных
	hsz_service = dde_create_string_handle_a(@id, @name, CP_WINANSI)

	return false if(hszService == 0) 

    # Регистрация сервиса
	return false unless dde_name_service(@id, hsz_service, NULL, DNS_REGISTER)

	return true
    end
  end
end