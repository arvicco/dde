/*
Класс XlTable предназначен для представления, хранения, обработки получаемых данных.

!!! Внимание: конструктору класса передавать ТОЛЬКО формальные параметры функции main (так как метод предполагает, что эти 
    данные валидные, т.е. значение переменной argc соответствует кол-ву элементов в массиве argv)
*/
#ifndef XLTABLE_H_
#define XLTABLE_H_

#include <windows.h>
#include <iostream>
#include <cstring>
#include <ddeml.h>
#include <vector>

using std::string;
using std::vector;
// Возможные типы получаемых данных
enum ddt {tdtFloat=1, tdtString, tdtBool, tdtError, tdtBlank, tdtInt, tdtSkip, tdtTable = 16}; 

//////////////////////////////////////////////////////////////////////////////////////////////

class XlTable
{
public:
	#ifdef GTEST_ON	// Если определен макрос GTEST_ON, то вставить объявления тестов
    FRIEND_TEST(XlTable, Constructor);			// Тест конструктора
	FRIEND_TEST(XlTable, IsData);				// Тест функции Isdata
	FRIEND_TEST(XlTable, InitTable);			// Тест функции InitTable
	FRIEND_TEST(XlTable, SetDelivers);			// Тест функции SetDelivers
	FRIEND_TEST(XlTable, Delete);				// Тест функции Delete
	FRIEND_TEST(XlTable, DoubleToString);		// Тест функции DoubleToString
	FRIEND_TEST(XlTable, GetBuf);				// Тест функции GetBuf
	#endif

	bool IsData();						// Функция определения доступности данных
    bool GetData(HDDEDATA hData);		// Функция получения данных
	bool DrawTable();					// Функция вывода полученных данных	
    XlTable();							// Конструктор
	void SetDelivers(string row_d, string col_d, string data_d, string topit_d);	// Функция получения разделителей
	void Delete();						// Функция очистки таблицы
	void GetBuf(string b);				// Функция получения имени топика и раздела
private:
    WORD col;										// Количество столбцов в таблице
    WORD row;										// Количество строк в таблице
	string buf;										// Имя топика и раздела
	string row_deliver;								// Поля соответствуют полям класса CommandLineParser
	string col_deliver;								//				--//--	
	string data_deliver;							//				--//--
	string topit_deliver;							//				--//--
	void InitTable(WORD r, WORD c, string s);		// Функция инициализации таблицы нуль-значением
	vector <vector<string>> table_data;				// Указатель на таблицу данных
	string DoubleToString(double num);				// Функция перевода вещественных чисел в строку
};
#endif // XLTABLE_H_
