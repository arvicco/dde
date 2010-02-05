
module DDE
  class Server
    include DDE::Functions
    
    attr_accessor :name, :id

    def initialize(name = 'excel')
      @name = name
      @id = 0
    end

    def connect &callback

    return false unless register_clipboard_format("XlTable")

	# Инициализация DDEML библиотеки (returns 0 if everything OK)
    @id, status = dde_initialize APPCLASS_STANDARD, &callback
	return false unless status == DMLERR_NO_ERROR

	# Получение идентификаторов строк для сервиса, раздела и элемента данных
	service = dde_create_string_handle_a(@id, @name, CP_WINANSI)

	return false unless service

    # Регистрация сервиса
	return false unless dde_name_service(@id, hsz_service, NULL, DNS_REGISTER)

	return true
    end
  end
end