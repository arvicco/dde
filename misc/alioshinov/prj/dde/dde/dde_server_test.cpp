//
//					Тестирование функций класса DDE
//
#include <gtest/gtest.h>

#define GTEST_ON		// включаем возможность использовать тесты

#include "dde_server.h"

// Тест конструктора с параметрами
TEST(DdeServer, Constructor_with_parameters)
{
	string str = "MyService";
	DDE srv(str);
	ASSERT_EQ(0, srv.idInst);
	ASSERT_EQ(str, srv.chService);
	ASSERT_EQ(NULL, srv.hszService);

	str = "dsfsdfsd";
	DDE server(str);
	ASSERT_EQ(0, server.idInst);
	ASSERT_EQ(str, server.chService);
	ASSERT_EQ(NULL, server.hszService);
}

// Тест конструктора без параметров
TEST(DdeServer, Constructor_without_parameters)
{
	string str = "";
	DDE srv;
	ASSERT_EQ(0, srv.idInst);
	ASSERT_EQ(str, srv.chService);
	ASSERT_EQ(NULL, srv.hszService);
}

// Тест функции DdeInit
TEST(DdeServer, DdeInit)
{
	DDE srv;
	string str = "My service";
	srv.DdeInit(str);
	EXPECT_EQ(0, srv.idInst);
	EXPECT_EQ(str, srv.chService);
	EXPECT_EQ(NULL, srv.hszService);

	str = "Testing";
	DDE server("Hello");
	server.DdeInit(str);
	EXPECT_EQ(0, server.idInst);
	EXPECT_EQ(str, server.chService);
	EXPECT_EQ(NULL, server.hszService);
}