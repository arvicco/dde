/*
����� XlTable ������������ ��� �������������, ��������, ��������� ���������� ������.

!!! ��������: ������������ ������ ���������� ������ ���������� ��������� ������� main (��� ��� ����� ������������, ��� ��� 
    ������ ��������, �.�. �������� ���������� argc ������������� ���-�� ��������� � ������� argv)
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
// ��������� ���� ���������� ������
enum ddt {tdtFloat=1, tdtString, tdtBool, tdtError, tdtBlank, tdtInt, tdtSkip, tdtTable = 16}; 

//////////////////////////////////////////////////////////////////////////////////////////////

class XlTable
{
public:
	#ifdef GTEST_ON	// ���� ��������� ������ GTEST_ON, �� �������� ���������� ������
    FRIEND_TEST(XlTable, Constructor);			// ���� ������������
	FRIEND_TEST(XlTable, IsData);				// ���� ������� Isdata
	FRIEND_TEST(XlTable, InitTable);			// ���� ������� InitTable
	FRIEND_TEST(XlTable, SetDelivers);			// ���� ������� SetDelivers
	FRIEND_TEST(XlTable, Delete);				// ���� ������� Delete
	FRIEND_TEST(XlTable, DoubleToString);		// ���� ������� DoubleToString
	FRIEND_TEST(XlTable, GetBuf);				// ���� ������� GetBuf
	#endif

	bool IsData();						// ������� ����������� ����������� ������
    bool GetData(HDdeDATA hData);		// ������� ��������� ������
	bool DrawTable();					// ������� ������ ���������� ������	
    XlTable();							// �����������
	void SetDelivers(string row_d, string col_d, string data_d, string topit_d);	// ������� ��������� ������������
	void Delete();						// ������� ������� �������
	void GetBuf(string b);				// ������� ��������� ����� ������ � �������
private:
    WORD col;										// ���������� �������� � �������
    WORD row;										// ���������� ����� � �������
	string buf;										// ��� ������ � �������
	string row_deliver;								// ���� ������������� ����� ������ CommandLineParser
	string col_deliver;								//				--//--	
	string data_deliver;							//				--//--
	string topit_deliver;							//				--//--
	void InitTable(WORD r, WORD c, string s);		// ������� ������������� ������� ����-���������
	vector <vector<string>> table_data;				// ��������� �� ������� ������
	string DoubleToString(double num);				// ������� �������� ������������ ����� � ������
};
#endif // XLTABLE_H_
