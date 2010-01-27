/*
Класс DDE представляет собой менеджер DDE-сервера.
Он предназначен для создания, удаления сервера, хранения получаемых данных
*/
#ifndef DDE_SERVER_H_
#define DDE_SERVER_H_

#include <windows.h>
#include <cstring>
#include <ddeml.h>
#include "XlTable.h"

//Указатель на функцию обратного вызова
typedef HDDEDATA (CALLBACK *CallBack) (UINT uType,	UINT uFmt, HCONV hConv,	HSZ hsz1, HSZ hsz2, 
												HDDEDATA hData, DWORD dwData1, DWORD dwData2);	

class DDE
{
public:
	#ifdef GTEST_ON	// Если определен макрос GTEST_ON, то вставить объявления тестов
	FRIEND_TEST(DdeServer, Constructor_with_parameters);	// Тест конструктора с параметрами
	FRIEND_TEST(DdeServer, Constructor_without_parameters);	// Тест конструктора без параметров
	FRIEND_TEST(DdeServer, DdeInit);						// Тест функции DdeInit
	#endif
	DDE(std::string service);				// Конструтор с параметром. 
	DDE();									// Конструтор без параметров.
	~DDE();                                 // Деструктор. Освобождает ресурсы, занятые сервером                            
	void DdeInit(std::string service);		// Функция инициализации 	
	bool Connect(CallBack DdeCallback);		// Функция регистрирует сервис в DDEML библиотеке и проводит инициализацию
    void Disconnect();                      // Функция отменяет регистрацию сервиса в библиотеке DDEML        
	HSZ GetName();							// Функция возвращает имя зарегистрированного сервиса
	DWORD GetId();							// Функция возвращает идентификатор зарегистрированного сервиса
	XlTable xltable;						// Данные, получаемые сервером
private:
	std::string chService;                  // Хранит имя сервиса                           
	DWORD idInst;							// Идентификатор зарегистрированного приложения
	HSZ hszService;                         // Идентификатор сервиса
};

#endif // DE_SERVER_H_
