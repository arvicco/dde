#include <gtest/gtest.h>

#define GTEST_ON		// включаем возможность использовать тесты

#include "XlTable.h"

// Тест конструтора
TEST(XlTable, Constructor)
{
	XlTable xlt;

	ASSERT_EQ(true,xlt.table_data.empty());
	ASSERT_EQ(0,xlt.row);
	ASSERT_EQ(0,xlt.col);
	ASSERT_EQ(true,xlt.row_deliver.empty());
	ASSERT_EQ(true,xlt.row_deliver.empty());
}

// Тест функции IsData
TEST(XlTable, IsData)
{
	XlTable xlt;
	vector<string> v;

	EXPECT_EQ(false,xlt.IsData());				// После вызова конструктора таблица должна быть пустой

	v.push_back("row1 col1");
	v.push_back("row1 col2");
	xlt.table_data.push_back(v);				// Если таблица содержит данные, но кол-во строк равно 0
	xlt.row = 0;
	xlt.col = 2;
	EXPECT_EQ(false,xlt.IsData());

	xlt.row = 1;								// Все корректно
	EXPECT_EQ(true,xlt.IsData());

	xlt.row = 2;								// Кол-во строк не соответствует действительности
	EXPECT_EQ(false,xlt.IsData());

	v.clear();
	v.push_back("row2 col1");
	v.push_back("row2 col2");
	xlt.table_data.push_back(v);

	xlt.col = 3;								// Кол-во столбцов не соответствует действительности
	EXPECT_EQ(false,xlt.IsData());

	xlt.col = 2;								// Все корректно
	EXPECT_EQ(true,xlt.IsData());

	xlt.col = 0;								// Если таблица содержит данные, но кол-во столбцов равно 0
	EXPECT_EQ(false,xlt.IsData());

	xlt.table_data.pop_back();					// Если таблица пуста, но кол-во строк и кол-во столбцов отличны от нуля
	xlt.table_data.pop_back();
	EXPECT_EQ(false,xlt.IsData());
}

// Тест функции InitTable
TEST (XlTable, InitTable)
{
	XlTable xlt;
	string str = "TeSt";
	xlt.InitTable(3,4,str);
	ASSERT_EQ(3,xlt.row);
	ASSERT_EQ(4,xlt.col);
	for (int i = 0; i < 3; i++)
		for (int j = 0; j < 4; j++)
			ASSERT_EQ(str, xlt.table_data[i][j]);
}

// Тест функции Delete
TEST (XlTable, Delete)
{
	XlTable xlt;
	xlt.InitTable(4,4,"nnn");
	xlt.Delete();
	EXPECT_EQ(0,xlt.row);
	EXPECT_EQ(0,xlt.col);
	EXPECT_EQ(true,xlt.table_data.empty());
}

// Тест функции SetDelivers
TEST(XlTable, SetDelivers)
{
	XlTable xlt;
	string str_row = "..sdf";
	string str_col = "sdsdsd sd";
	string str_data = "HKJjhjhkhjk";
	string str_topit = "ООпаарывпы";
	xlt.SetDelivers(str_row, str_col, str_data, str_topit);
	EXPECT_EQ(str_row, xlt.row_deliver);
	EXPECT_EQ(str_col, xlt.col_deliver);
	EXPECT_EQ(str_data, xlt.data_deliver);
	EXPECT_EQ(str_topit, xlt.topit_deliver);

	str_row = "";
	str_col = "___";
	str_data = "пррппрппп";
	str_topit = "ддлжтргпнг";
	xlt.SetDelivers(str_row, str_col, str_data, str_topit);
	EXPECT_EQ(str_row, xlt.row_deliver);
	EXPECT_EQ(str_col, xlt.col_deliver);
	EXPECT_EQ(str_data, xlt.data_deliver);
	EXPECT_EQ(str_topit, xlt.topit_deliver);
}

// Тест функции DoubleToString
TEST(XlTable, DoubleToString)
{
	XlTable xlt;
	double dbl = 1.2;
	string str = "1.2";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = 0.00001;
	str = "0.00001";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = 103430;
	str = "103430";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = 0.103430;
	str = "0.10343";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = 0.100004;
	str = "0.100004";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = 1001.1001;
	str = "1001.1001";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = .2003;
	str = "0.2003";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));

	dbl = 0.;
	str = "0";
	EXPECT_EQ(str,xlt.DoubleToString(dbl));
}

// Тест функции GetBuf
TEST(XlTable, GetBuf)
{
	XlTable xlt;
	string topit = "Topic[item]";
	xlt.GetBuf(topit);
	EXPECT_EQ(topit, xlt.buf);
	topit = "[ывывыв]ааааа";
	xlt.GetBuf(topit);
	EXPECT_EQ(topit, xlt.buf);
}