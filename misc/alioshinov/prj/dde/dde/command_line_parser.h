/*
����� CommandLineParser ������������ ��� ������� ���������� ��������� ������.

!!! ��������: ������������ ������ ���������� ������ ���������� ��������� ������� main (��� ��� ����� ������������, ��� ��� 
    ������ ��������, �.�. �������� ���������� argc ������������� ���-�� ��������� � ������� argv)
*/
#ifndef COMMAND_LINE_PARSER_H_
#define COMMAND_LINE_PARSER_H_

#include <iostream>
#include <cstring>

using std::string;
using std::cout;

enum status_type {CML_OK, CML_HELP};		// ������ ������
											// CML_OK - ������ ������ �������
											// CML_HELP - �������� ������ � ������� ������

class CommandLineParser
{
public:

	#ifdef GTEST_ON	// ���� ��������� ������ GTEST_ON, �� �������� ���������� ������
    FRIEND_TEST(CommandLineParser, Constructor_with_parameters);	// ���� ������������ � �����������
	FRIEND_TEST(CommandLineParser, Constructor_without_parameters); // ���� ������������ ��� ����������
	FRIEND_TEST(CommandLineParser, GetServiceName);					// ���� ������� GetServiceName
	FRIEND_TEST(CommandLineParser, GetColDeliver);					// ���� ������� GetColDeliver
	FRIEND_TEST(CommandLineParser, GetRowDeliver);					// ���� ������� GetRowDeliver
	FRIEND_TEST(CommandLineParser, GetDataDeliver);					// ���� ������� GetDataDeliver
	FRIEND_TEST(CommandLineParser, GetTopItDeliver);				// ���� ������� GetTopItDeliver
	FRIEND_TEST(CommandLineParser, GetStatus);						// ���� ������� GetStatus
	#endif

	CommandLineParser(int argc, char *argv[]);	// �����������. �������������� ���� ����������� �������
	CommandLineParser();			// ����������� ��� ����������. �������������� ���� ������ ���������� �� ���������
	string GetServiceName();		// ������� ���������� ��� ����������� �������
	string GetColDeliver();			// ������� ���������� ����������� ����� ���������
	string GetRowDeliver();			// ������� ���������� ����������� ����� ��������
	string GetDataDeliver();		// ������� ���������� ����������� ����� �������� ������
	string GetTopItDeliver();		// ������� ���������� ����������� ����� ������� �� ������ ������
	status_type GetStatus();		// ������� ���������� ������ ������� ���������� ������
	void Help();					// ������� ������� ������

private:
	status_type status;				// ������ ������� ���������� ��������� ������
	string service;					// ��� ������� Dde
	string col_deliver;				// ����������� ������� � �������
	string row_deliver;				// ����������� ����� � �������
	string data_deliver;			// ����������� ������
	string topit_deliver;			// ����������� ����� ������� �� ������ ������
};
#endif