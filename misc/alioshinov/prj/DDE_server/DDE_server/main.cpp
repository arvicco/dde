#include <iostream>
#include <windows.h>
#include <conio.h>
#include <queue>
#include "command_line_parser.h"
#include "dde_server.h"
#define _CRT_SECURE_NO_WARNINGS		// дает возможность использовать функцию _ecvt_s

DDE server;							// сервер
std::queue<XlTable> q;				// очередь, содержит таблицы, которые необходимо вывести на консоль
XlTable xlt;						// содержит выводимую таблицу

HANDLE hMutex;						// Синхронизирует вывод на консоль для различных потоков
HANDLE hMutex1;						// Синхронизирует работу с очередью
HANDLE hSemaphore;					// Синхронизирует работу потоков
HANDLE hThread;						// Хэндл потока, выводящего таблицу на консоль

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           Draw(LPVOID NoUse)
//  Назначение:    Выводит таблицы, находящиеся в очереди, на консоль
//  Вход:          не используется
//  Выход:         нет
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
DWORD WINAPI Draw(LPVOID NoUse)
{
	while(1)
	{	// Ждем данных для вывода
		WaitForSingleObject(hSemaphore, INFINITE);
		// Данные пришли, получаем эксклюзивный доступ к очереди
		WaitForSingleObject(hMutex1, INFINITE);
		if (!q.empty())
		{	// если очередь не пуста, то извлекаем таблицу и открываем доступ к очереди другим потокам
			xlt = q.front();
			q.pop();
			ReleaseMutex(hMutex1);
			// Выводим таблицу на консоль
			WaitForSingleObject(hMutex,INFINITE);
				xlt.DrawTable();
			ReleaseMutex(hMutex);
		}
		else ReleaseMutex(hMutex1);
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, 
//							   HSZ hsz2,HDDEDATA hData, DWORD dwData1, DWORD dwData2)
//  Назначение:    Функция обратного вызова дляобработки DDE-запростов
//  Вход:          uType - тип сообщения, остальные параметры зависят от сообщения.
//  Выход:         результат
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
HDDEDATA CALLBACK DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, HSZ hsz2,
									HDDEDATA hData, DWORD dwData1, DWORD dwData2)
{
	switch (uType)
	{
	//----------------------------------------------------------------------
	case XTYP_CONNECT:			// Создание канала передачи данных
		if (hsz2 == server.GetName())
		    return (HDDEDATA)TRUE;	// Сервер поддерживает Service
		else
		{	// канал передачи не создан
			WaitForSingleObject(hMutex, INFINITE);
				std::cerr<<"Канал передачи данных не создан\n";
			ReleaseMutex(hMutex);
            return FALSE;
        }			// Сервер не поддерживает Service
	//----------------------------------------------------------------------
	case XTYP_POKE:				// Пришли данные от DDE-клиента
		{
			BOOL flag;
			CHAR buf[200];
			// Получаем данные от клиента
			flag = server.xltable.GetData(hData);
			
			if (!flag)
			{
				// Преобразуем HSZ в string (название топика и итема в формате [topic]item)
				DdeQueryStringA(server.GetId(),hsz1,buf,200,CP_WINANSI);
				server.xltable.GetBuf(buf);
				// Помещаем таблицу в очередь
				WaitForSingleObject(hMutex1,INFINITE);
					q.push(server.xltable);
				ReleaseMutex(hMutex1);
				// Позволяем запустится потоку вывода таблицы на экран
				ReleaseSemaphore(hSemaphore,1,NULL);
					
	            return((HDDEDATA)DDE_FACK);  // Признак успешного завершения транзакции                 
			}
			return (HDDEDATA)TRUE;			 // Признак неудачной транзакции 	
		}
	case XTYP_DISCONNECT:		// Отключение DDE-клиента	
		{
			server.xltable.Delete(); 
			break;
		}
	case XTYP_ERROR:			// Ошибка
		{
			WaitForSingleObject(hMutex, INFINITE);
				std::cerr<<"DDE error.\n";
			ReleaseMutex(hMutex);
			break;
		}

	//----------------------------------------------------------------------
	default:
		return (HDDEDATA)NULL;
	//----------------------------------------------------------------------
	}
	return((HDDEDATA)NULL);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           CtrlHandler(DWORD type)
//  Назначение:    Обработчик Ctr-C
//  Вход:          название события (нас интересует только событие CTRL_C_EVENT)
//  Выход:         нет
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool CtrlCHandler(DWORD type)
{
	switch(type) 
	{
	case CTRL_C_EVENT: 
		// завершаем приложение, освобождаем все ресурсы
		WaitForSingleObject(hMutex, INFINITE);
			std::cout<<"DDE-сервер завершен нажатием комбинации клавиш CTRL+C\n";
		ReleaseMutex(hMutex);
		CloseHandle(hMutex);
		CloseHandle(hMutex1);
		CloseHandle(hSemaphore);
		CloseHandle(hThread);
		ExitProcess(1);
		return true;

	case CTRL_CLOSE_EVENT: 
		return false; 
 
	case CTRL_BREAK_EVENT: 
		return true; 
 
	case CTRL_LOGOFF_EVENT: 
		return false; 
 
	case CTRL_SHUTDOWN_EVENT: 
		return false; 
 
    default: 
		return false; 
	} 
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Имя:           main(int argc, char *argv[])
//  Назначение:    главная функция
//  Вход:          командная строка
//  Выход:         возвращаемое программой значение
//  Примечание:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
int main (int argc, char *argv[])
{
	MSG msg;
	// Создаем мьютексы и семофор для синхронизации потоков
	hMutex = CreateMutex(NULL, FALSE, NULL);
	if (hMutex == NULL)	return GetLastError();

	hMutex1 = CreateMutex(NULL, FALSE, NULL);
	if (hMutex1 == NULL) return GetLastError();

	hSemaphore = CreateSemaphore(NULL,0, 1000, NULL);
	if (hSemaphore == NULL)	return GetLastError();
	// создаем поток для вывода таблицы на консоль
	hThread = CreateThread(NULL,0, Draw, NULL, 0, NULL);
	if (hThread == NULL)	return GetLastError();
	// Устанавливаем русскую локаль
	setlocale(LC_CTYPE, "Russian_Russia.1251");	
	// Устанавливаем свой обработчик нажатия клавиш Ctrl-C
	if(!SetConsoleCtrlHandler((PHANDLER_ROUTINE)CtrlCHandler, true)) 
		std::cerr<<"Не удалось установить обработчик нажатия комбинации клавиш CTRL+C\n";
	// Считываем параметры из командной строки
	CommandLineParser clp(argc, argv);				
	if (CML_OK != clp.GetStatus())
	{
		clp.Help();
		return 1;
	}
	// инициализируем DDE-сервер
	server.DdeInit(clp.GetServiceName());
	server.xltable.SetDelivers(clp.GetRowDeliver(), clp.GetColDeliver(), clp.GetDataDeliver(), clp.GetTopItDeliver());	

	if (!server.Connect(DdeCallback)) 
	{
		std::cerr<<"Connect error!!!\n";
		return 1;
	}
	// Запускаем цикл обработки сообщений (необходим для обработки DDE)
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	// Освобождаем все ресурсы
	CloseHandle(hMutex);
	CloseHandle(hMutex1);
	CloseHandle(hSemaphore);
	CloseHandle(hThread);
	return 0;
}

