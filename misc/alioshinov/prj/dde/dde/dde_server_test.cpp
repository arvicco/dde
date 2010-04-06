//
//					������������ ������� ������ Dde
//
#include <gtest/gtest.h>

#define GTEST_ON		// �������� ����������� ������������ �����

#include "dde_server.h"

// ���� ������������ � �����������
TEST(DdeServer, Constructor_with_parameters)
{
	string str = "MyService";
	Dde srv(str);
	ASSERT_EQ(0, srv.idInst);
	ASSERT_EQ(str, srv.chService);
	ASSERT_EQ(NULL, srv.hszService);

	str = "dsfsdfsd";
	Dde server(str);
	ASSERT_EQ(0, server.idInst);
	ASSERT_EQ(str, server.chService);
	ASSERT_EQ(NULL, server.hszService);
}

// ���� ������������ ��� ����������
TEST(DdeServer, Constructor_without_parameters)
{
	string str = "";
	Dde srv;
	ASSERT_EQ(0, srv.idInst);
	ASSERT_EQ(str, srv.chService);
	ASSERT_EQ(NULL, srv.hszService);
}

// ���� ������� DdeInit
TEST(DdeServer, DdeInit)
{
	Dde srv;
	string str = "My service";
	srv.DdeInit(str);
	EXPECT_EQ(0, srv.idInst);
	EXPECT_EQ(str, srv.chService);
	EXPECT_EQ(NULL, srv.hszService);

	str = "Testing";
	Dde server("Hello");
	server.DdeInit(str);
	EXPECT_EQ(0, server.idInst);
	EXPECT_EQ(str, server.chService);
	EXPECT_EQ(NULL, server.hszService);
}