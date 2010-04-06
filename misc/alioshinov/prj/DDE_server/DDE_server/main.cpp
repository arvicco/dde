#include <iostream>
#include <windows.h>
#include <conio.h>
#include <queue>
#include "command_line_parser.h"
#include "dde_server.h"
#define _CRT_SECURE_NO_WARNINGS		// ���� ����������� ������������ ������� _ecvt_s

Dde server;							// ������
std::queue<XlTable> q;				// �������, �������� �������, ������� ���������� ������� �� �������
XlTable xlt;						// �������� ��������� �������

HANDLE hMutex;						// �������������� ����� �� ������� ��� ��������� �������
HANDLE hMutex1;						// �������������� ������ � ��������
HANDLE hSemaphore;					// �������������� ������ �������
HANDLE hThread;						// ����� ������, ���������� ������� �� �������

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           Draw(LPVOID NoUse)
//  ����������:    ������� �������, ����������� � �������, �� �������
//  ����:          �� ������������
//  �����:         ���
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
DWORD WINAPI Draw(LPVOID NoUse)
{
	while(1)
	{	// ���� ������ ��� ������
		WaitForSingleObject(hSemaphore, INFINITE);
		// ������ ������, �������� ������������ ������ � �������
		WaitForSingleObject(hMutex1, INFINITE);
		if (!q.empty())
		{	// ���� ������� �� �����, �� ��������� ������� � ��������� ������ � ������� ������ �������
			xlt = q.front();
			q.pop();
			ReleaseMutex(hMutex1);
			// ������� ������� �� �������
			WaitForSingleObject(hMutex,INFINITE);
				xlt.DrawTable();
			ReleaseMutex(hMutex);
		}
		else ReleaseMutex(hMutex1);
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, 
//							   HSZ hsz2,HDdeDATA hData, DWORD dwData1, DWORD dwData2)
//  ����������:    ������� ��������� ������ ������������ Dde-���������
//  ����:          uType - ��� ���������, ��������� ��������� ������� �� ���������.
//  �����:         ���������
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
HDdeDATA CALLBACK DdeCallback(UINT uType, UINT uFmt, HCONV hConv, HSZ hsz1, HSZ hsz2,
									HDdeDATA hData, DWORD dwData1, DWORD dwData2)
{
	switch (uType)
	{
	//----------------------------------------------------------------------
	case XTYP_CONNECT:			// �������� ������ �������� ������
		if (hsz2 == server.GetName())
		    return (HDdeDATA)TRUE;	// ������ ������������ Service
		else
		{	// ����� �������� �� ������
			WaitForSingleObject(hMutex, INFINITE);
				std::cerr<<"����� �������� ������ �� ������\n";
			ReleaseMutex(hMutex);
            return FALSE;
        }			// ������ �� ������������ Service
	//----------------------------------------------------------------------
	case XTYP_POKE:				// ������ ������ �� Dde-�������
		{
			BOOL flag;
			CHAR buf[200];
			// �������� ������ �� �������
			flag = server.xltable.GetData(hData);
			
			if (!flag)
			{
				// ����������� HSZ � string (�������� ������ � ����� � ������� [topic]item)
				DdeQueryStringA(server.GetId(),hsz1,buf,200,CP_WINANSI);
				server.xltable.GetBuf(buf);
				// �������� ������� � �������
				WaitForSingleObject(hMutex1,INFINITE);
					q.push(server.xltable);
				ReleaseMutex(hMutex1);
				// ��������� ���������� ������ ������ ������� �� �����
				ReleaseSemaphore(hSemaphore,1,NULL);
					
	            return((HDdeDATA)Dde_FACK);  // ������� ��������� ���������� ����������
			}
			return (HDdeDATA)TRUE;			 // ������� ��������� ����������
		}
	case XTYP_DISCONNECT:		// ���������� Dde-�������
		{
			server.xltable.Delete(); 
			break;
		}
	case XTYP_ERROR:			// ������
		{
			WaitForSingleObject(hMutex, INFINITE);
				std::cerr<<"Dde error.\n";
			ReleaseMutex(hMutex);
			break;
		}

	//----------------------------------------------------------------------
	default:
		return (HDdeDATA)NULL;
	//----------------------------------------------------------------------
	}
	return((HDdeDATA)NULL);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           CtrlHandler(DWORD type)
//  ����������:    ���������� Ctr-C
//  ����:          �������� ������� (��� ���������� ������ ������� CTRL_C_EVENT)
//  �����:         ���
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool CtrlCHandler(DWORD type)
{
	switch(type) 
	{
	case CTRL_C_EVENT: 
		// ��������� ����������, ����������� ��� �������
		WaitForSingleObject(hMutex, INFINITE);
			std::cout<<"Dde-������ �������� �������� ���������� ������ CTRL+C\n";
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
//  ���:           main(int argc, char *argv[])
//  ����������:    ������� �������
//  ����:          ��������� ������
//  �����:         ������������ ���������� ��������
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
int main (int argc, char *argv[])
{
	MSG msg;
	// ������� �������� � ������� ��� ������������� �������
	hMutex = CreateMutex(NULL, FALSE, NULL);
	if (hMutex == NULL)	return GetLastError();

	hMutex1 = CreateMutex(NULL, FALSE, NULL);
	if (hMutex1 == NULL) return GetLastError();

	hSemaphore = CreateSemaphore(NULL,0, 1000, NULL);
	if (hSemaphore == NULL)	return GetLastError();
	// ������� ����� ��� ������ ������� �� �������
	hThread = CreateThread(NULL,0, Draw, NULL, 0, NULL);
	if (hThread == NULL)	return GetLastError();
	// ������������� ������� ������
	setlocale(LC_CTYPE, "Russian_Russia.1251");	
	// ������������� ���� ���������� ������� ������ Ctrl-C
	if(!SetConsoleCtrlHandler((PHANDLER_ROUTINE)CtrlCHandler, true)) 
		std::cerr<<"�� ������� ���������� ���������� ������� ���������� ������ CTRL+C\n";
	// ��������� ��������� �� ��������� ������
	CommandLineParser clp(argc, argv);				
	if (CML_OK != clp.GetStatus())
	{
		clp.Help();
		return 1;
	}
	// �������������� Dde-������
	server.DdeInit(clp.GetServiceName());
	server.xltable.SetDelivers(clp.GetRowDeliver(), clp.GetColDeliver(), clp.GetDataDeliver(), clp.GetTopItDeliver());	

	if (!server.Connect(DdeCallback)) 
	{
		std::cerr<<"Connect error!!!\n";
		return 1;
	}
	// ��������� ���� ��������� ��������� (��������� ��� ��������� Dde)
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	// ����������� ��� �������
	CloseHandle(hMutex);
	CloseHandle(hMutex1);
	CloseHandle(hSemaphore);
	CloseHandle(hThread);
	return 0;
}

