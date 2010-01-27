#include <cstdlib>
#include "XlTable.h"

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::XlTable()
//  Назначение:    Конструктор класса (инициализирует класс начальными значениями)
//  Вход:          нет
//  Выход:         нет
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
XlTable::XlTable()
{
	table_data.clear();
	row_deliver.clear();
	col_deliver.clear();
	row = 0;
	col = 0;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::IsData()
//  Назначение:    Определяет есть ли данные в таблице или нет
//  Вход:          нет
//  Выход:         true - таблица содержит данные
//				   false - таблица не содержит данные или они не корректны
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool XlTable::IsData()
{
	if (!table_data.empty())
	{
		if ((row > 0) && (col >0)) 
		{
			if (row != table_data.size()) return false; 
			if (col != table_data[0].size()) return false;
			return true;
		}
		else return false;
	}
	else 
		return false;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::InitTable(WORD r, WORD c, string s)
//  Назначение:    Инициализирует таблицу
//  Вход:          r - кол-во строк, c - кол-во колонок, s - строка инициализации
//  Выход:         нет
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::InitTable(WORD r, WORD c, string s)
{
	table_data.assign(r, std::vector<std::string>(c,s));
	row = r;
	col = c;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::GetBuf(string b)
//  Назначение:    Инициализирует имя топика и итема
//  Вход:          b - имя топика и итема пришедших данных
//  Выход:         нет
//  Примечание:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::GetBuf(string b)
{
	buf = b;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::DrawTable()
//  Назначение:    Функция вывода таблицы с учетом всех разделителей
//  Вход:          нет
//  Выход:         true - таблица выведена успешно
//				   false - таблица пуста и не может быть выведена
//  Примечание:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool XlTable::DrawTable()
{
	if (table_data.empty()) return false;
	string top;
	string it;
	size_t pos;

	if (topit_deliver.length() > 0)	// Если указан разделитель между разными таблицами, то определяем 
	{								// имя топика и итема ( у нас они хранятся в виде [topic]item
		pos = buf.find("]");
		top = buf.substr(1, pos-1);	// top содержит имя топика
		it = buf.substr(pos+1, buf.length() - 1);	// it содержит имя итема
	}

	if (data_deliver.length() > 0)	// Если указан разделитель данных, то выводим его
		std::cout<<data_deliver.c_str()<<"\n";	// Выводим разделитель данных

	for (int i = 0; i < row; i++)	// Выводим таблицу
	{
		if (topit_deliver.length()>0)
			std::cout<<top.c_str()<<topit_deliver.c_str()<<it.c_str()<<topit_deliver.c_str();	
		for (int j = 0; j < col; j++)
		{
			std::cout<<table_data[i][j].c_str();
			if ( j != (col - 1)) std::cout<<col_deliver.c_str();
		}
		std::cout<<row_deliver.c_str();
	}
	return true;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::SetDelivers(string row_d, string col_d, string data_d, string topit_d)
//  Назначение:    Устанавливаем разделители
//  Вход:          row_d - разделитель строк, col_d - разделитель колонок,
//				   data_d - разделитель блоков данных, topit_d - разделитель данных разных таблиц
//  Выход:         нет
//  Примечание:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::SetDelivers(string row_d, string col_d, string data_d, string topit_d)
{
	row_deliver = row_d;
	col_deliver = col_d;
	data_deliver = data_d;
	topit_deliver = topit_d;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::Delete()
//  Назначение:    Очищает таблицу
//  Вход:          нет
//  Выход:         нет
//  Примечание:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::Delete()
{
	table_data.clear();
	row = 0;
	col = 0;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           bool XlTable::GetData(HDDEDATA hData)
//  Назначение:    Получает данные от клиента и на их основе формирует таблицу данных и типов
//  Вход:          идентификатор области памяти (передается функции обратного вызова системой)          
//  Выход:         0 - при успешном выполнении, 1 - при ошибке
//  Примечание:    доступна пользователю
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool XlTable::GetData(HDDEDATA hData)
{
	 WORD row_n	= 0;
	 WORD col_n = 0;
     DWORD length = 0;                       // Длина данных
     UINT offset  = 0;                       // Позиция в массиве данных
     PBYTE data   = NULL;                    // Указатель на принимаемые данные
	 WORD cnum = 0, rnum = 0;                // Переменные для перебора строк и столбцов таблицы
	 WORD cb = 0;                            // Переменная для хранения cb
	 WORD size = 0;                          // Переменная для хранения размера данных
	 WORD type = 0;                          // Переменная для хранения типа данных
	 string str;
          
     union                                   // Объединение для преобразования данных в различные типы
     {
           WORD w; 
           double d; 
           BYTE b[8];
     } conv;  

     length = DdeGetData(hData, NULL, 0, 0); // Определяем объем поступивших от клиента данных
     if (length < 8) return 1;               

     data = new BYTE[length];                // Выделяем буфер требуемой длины
     DdeGetData(hData,data,length,0);        // Получаем данные

     // Смотрим, чтобы первый блок был tdtTable
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
	 if (conv.w != tdtTable) return 1;
   
     // Смотрим, чтобы cb равнялось 4
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
	 if (conv.w != 4) return 1;

     // Получаем количество строк таблицы
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
     row_n = conv.w;

     // Получаем количество столбцов таблицы
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
     col_n = conv.w;

	 if (!col_n || !row_n)                       // Если кол-во строк или столбцов равно нулю, то ошибка
     {
         delete []data;
         return 1;
     }    

	 InitTable(row_n,col_n,"");
  
     // Заполняем обе таблицы данными
	 while (offset < length)
	 {
         // Считываем тип данных  
		 conv.b[0] = data[offset++];
	     conv.b[1] = data[offset++];
		 type = conv.w;

		 switch(type)
		 {
		 case tdtFloat:                      // Пришли данные типа Double
			 {
	 			conv.b[0] = data[offset++];
				conv.b[1] = data[offset++];
				cb = conv.w;
				if (cb%8) 
				{
					delete []data; 
					return 1;
				}
				for (int tmp=0, i=0; i < cb; i++, tmp++)
				{
					conv.b[tmp] = data[offset++];
					if (tmp == 7)  // набрали одно число
					{
						 tmp = -1;
						 str = DoubleToString(conv.d);
						 if (!str.compare("")) 
						 {
							 delete []data;
							 return 1;
						 }
						 
						 table_data[rnum][cnum] = str;
						 cnum++;
						 if (cnum == col)
						 {
							 cnum = 0;
							 rnum++;
						 }
					 }
				}
			 }
			 break;

		 case tdtString:                     // Пришли данные типа String
			 {
                 // Считываем cb
			     conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;

        		 for (WORD i = 0; i < cb; )
			     {	 // Считываем длину строки (она без завершающнго нуля)
				    size = data[offset++];
					
				    // Заполняем строку символами 
				    for(int j = 0; j < size; j++)
                        table_data[rnum][cnum].push_back(data[offset++]);
				    i += 1 + size;
				    cnum++;
				    if (cnum == col)
				    {
					   cnum = 0;
					   rnum++;
			        }
                 }
			 }
			 break;

		 case tdtBool:                       // Пришли данные типа Bool
			 {
			     conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb%2) 
			     {
				     delete []data; 
				     Delete();
 				     return 1;
		         }
			     for (int tmp=0, i=0; i < cb; i++, tmp++)
			     {
				     conv.b[tmp] = data[offset++];
				     if (tmp == 1)  // набрали одно число
				     {
					    tmp = -1;
					    if (conv.w) table_data[rnum][cnum] = "true";
					    else table_data[rnum][cnum] = "false";
					    cnum++;
					    if (cnum == col)
					    {
						   cnum = 0;
						   rnum++;
                        }
                     }
                 }
			 }
			 break;

		 case tdtError:                      // Пришли данные типа Error
			 {
                 conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb%2) 
			     {
				     delete []data; 
				     Delete();
				     return 1;
                 }
			     for (int tmp=0, i=0; i < cb; i++, tmp++)
			     {
				     conv.b[tmp] = data[offset++];
				     if (tmp == 1)  // набрали одно число
				     {
					    tmp = -1;
					    table_data[rnum][cnum] = "error";
					    cnum++;
					    if (cnum == col)
					    {
						    cnum = 0;
						    rnum++;
				        }
				     }
		         }
			 }
			 break;

		 case tdtBlank:                      // Пришли данные типа Blank
			 {
                 conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb != 2)
			     {
				     delete []data; 
				     Delete();
				     return 1;
		         }
			     conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     size = conv.w;
			     for (int i = 0; i<size; i++)
			     {
				     table_data[rnum][cnum] = "";
				     cnum++;
				     if (cnum == col)
				     {
					     cnum = 0;
					     rnum++;
			         }
	             }
			 }
			 break;

		 case tdtInt:	                     // Пришли данные типа Int
			 {
				 char mas[20];
                 conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb%2) 
			     {
				     delete []data; 
				     Delete();
				     return 1;
		         }
			     for (int tmp=0, i=0; i < cb; i++, tmp++)
			     {
				     conv.b[tmp] = data[offset++];
				     if (tmp == 1)  // набрали одно число
				     {
					     tmp = -1;
						 _itoa_s(conv.w,mas,10);
					     table_data[rnum][cnum] = mas; 
					     cnum++;
					     if (cnum == col)
					     {
						     cnum = 0;
						     rnum++;
				         }
		             }
			     }
			 }
			 break;

		 case tdtSkip:                       // Не поддерживается
		     {
				 delete []data; 
				 Delete();
				 return 1;
             }
		 }
	 }

	 delete []data;                          // Возвращаем память системе  
	 return 0;                               // Функция выполнена успешно
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           XlTable::DoubleToString(double num)
//  Назначение:    Преобразует вещественное число в строку
//  Вход:          num - вещественное число типа double
//  Выход:         строка-представление вещественного числа, либо пустая строка в случае ошибки
//  Примечание:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
string XlTable::DoubleToString(double num)
{
string str;
	char buf[_CVTBUFSIZE];
	int decimal;
	int sign;
	int size;
	
	// преобразуем число
	if (_ecvt_s(buf, _CVTBUFSIZE, num, 15, &decimal, &sign))
	{
		return "";
	}
		
	str = buf;

	// преобразуем строку в удобный формат
	if (decimal > 0)
	{
		str.insert(decimal,".");
	}

	if (decimal <= 0)
	{
		decimal = -decimal;
		for(int i = 0; i < decimal; i++)
		{
			str.insert(0,"0");	
		}
		str.insert(0,"0.");
	}

	if (sign < 0)
	{
		str.insert(0,"-");
	}
	
	size = str.size();
	while (str[size-1] == '0')
	{
		str.erase(size-1,1);
		size = str.size();
	}

	size = str.size();
	if (str[size-1] == '.') 
	{
		str.erase(size-1, 1);
		return str;
	}

	return str;
}