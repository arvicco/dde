/*
Класс CommandLineParser предназначен для разбора параметров командной строки.

!!! Внимание: конструктору класса передавать ТОЛЬКО формальные параметры функции main (так как метод предполагает, что эти 
    данные валидные, т.е. значение переменной argc соответствует кол-ву элементов в массиве argv)
*/
#ifndef COMMAND_LINE_PARSER_H_
#define COMMAND_LINE_PARSER_H_

#include <iostream>
#include <cstring>

using std::string;
using std::cout;

enum status_type {CML_OK, CML_HELP};		// Статус класса
											// CML_OK - разбор прошел успешно
											// CML_HELP - прервать работу и вывести помощь

class CommandLineParser
{
public:

	#ifdef GTEST_ON	// Если определен макрос GTEST_ON, то вставить объявления тестов
    FRIEND_TEST(CommandLineParser, Constructor_with_parameters);	// Тест конструктора с параметрами
	FRIEND_TEST(CommandLineParser, Constructor_without_parameters); // Тест конструктора без параметров
	FRIEND_TEST(CommandLineParser, GetServiceName);					// Тест функции GetServiceName
	FRIEND_TEST(CommandLineParser, GetColDeliver);					// Тест функции GetColDeliver
	FRIEND_TEST(CommandLineParser, GetRowDeliver);					// Тест функции GetRowDeliver
	FRIEND_TEST(CommandLineParser, GetDataDeliver);					// Тест функции GetDataDeliver
	FRIEND_TEST(CommandLineParser, GetTopItDeliver);				// Тест функции GetTopItDeliver
	FRIEND_TEST(CommandLineParser, GetStatus);						// Тест функции GetStatus
	#endif

	CommandLineParser(int argc, char *argv[]);	// Конструктор. Инициализирует поля переданными данными
	CommandLineParser();			// Конструктор без параметров. Инициализирует поля класса значениями по умолчанию
	string GetServiceName();		// Функция возвращает имя переданного сервиса
	string GetColDeliver();			// Функция возвращает разделитель между колонками
	string GetRowDeliver();			// Функция возвращает разделитель между строками
	string GetDataDeliver();		// Функция возвращает разделитель между порциями данных
	string GetTopItDeliver();		// Функция возвращает разделитель между данными от разных таблиц
	status_type GetStatus();		// Функция возвращает статус разбора коммандной строки
	void Help();					// Функция выводит помощь

private:
	status_type status;				// Статус разбора параметров командной строки
	string service;					// Имя сервиса DDE
	string col_deliver;				// Разделитель колонок в таблице
	string row_deliver;				// Разделитель строк в таблице
	string data_deliver;			// Разделитель данных
	string topit_deliver;			// Разделитель между данными от разных таблиц
};
#endif