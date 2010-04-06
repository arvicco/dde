/*
����� Dde ������������ ����� �������� Dde-�������.
�� ������������ ��� ��������, �������� �������, �������� ���������� ������
*/
#ifndef Dde_SERVER_H_
#define Dde_SERVER_H_

#include <windows.h>
#include <cstring>
#include <ddeml.h>
#include "XlTable.h"

//��������� �� ������� ��������� ������
typedef HDdeDATA (CALLBACK *CallBack) (UINT uType,	UINT uFmt, HCONV hConv,	HSZ hsz1, HSZ hsz2,
												HDdeDATA hData, DWORD dwData1, DWORD dwData2);

class Dde
{
public:
	#ifdef GTEST_ON	// ���� ��������� ������ GTEST_ON, �� �������� ���������� ������
	FRIEND_TEST(DdeServer, Constructor_with_parameters);	// ���� ������������ � �����������
	FRIEND_TEST(DdeServer, Constructor_without_parameters);	// ���� ������������ ��� ����������
	FRIEND_TEST(DdeServer, DdeInit);						// ���� ������� DdeInit
	#endif
	Dde(std::string service);				// ���������� � ����������.
	Dde();									// ���������� ��� ����������.
	~Dde();                                 // ����������. ����������� �������, ������� ��������
	void DdeInit(std::string service);		// ������� ������������� 	
	bool Connect(CallBack DdeCallback);		// ������� ������������ ������ � DdeML ���������� � �������� �������������
    void Disconnect();                      // ������� �������� ����������� ������� � ���������� DdeML
	HSZ GetName();							// ������� ���������� ��� ������������������� �������
	DWORD GetId();							// ������� ���������� ������������� ������������������� �������
	XlTable xltable;						// ������, ���������� ��������
private:
	std::string chService;                  // ������ ��� �������                           
	DWORD idInst;							// ������������� ������������������� ����������
	HSZ hszService;                         // ������������� �������
};

#endif // DE_SERVER_H_
