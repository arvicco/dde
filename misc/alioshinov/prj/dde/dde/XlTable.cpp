#include <cstdlib>
#include "XlTable.h"

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::XlTable()
//  ����������:    ����������� ������ (�������������� ����� ���������� ����������)
//  ����:          ���
//  �����:         ���
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
XlTable::XlTable()
{
	table_data.clear();
	row_deliver.clear();
	col_deliver.clear();
	row = 0;
	col = 0;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::IsData()
//  ����������:    ���������� ���� �� ������ � ������� ��� ���
//  ����:          ���
//  �����:         true - ������� �������� ������
//				   false - ������� �� �������� ������ ��� ��� �� ���������
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool XlTable::IsData()
{
	if (!table_data.empty())
	{
		if ((row > 0) && (col >0)) 
		{
			if (row != table_data.size()) return false; 
			if (col != table_data[0].size()) return false;
			return true;
		}
		else return false;
	}
	else 
		return false;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::InitTable(WORD r, WORD c, string s)
//  ����������:    �������������� �������
//  ����:          r - ���-�� �����, c - ���-�� �������, s - ������ �������������
//  �����:         ���
//  ����������:    
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::InitTable(WORD r, WORD c, string s)
{
	table_data.assign(r, std::vector<std::string>(c,s));
	row = r;
	col = c;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::GetBuf(string b)
//  ����������:    �������������� ��� ������ � �����
//  ����:          b - ��� ������ � ����� ��������� ������
//  �����:         ���
//  ����������:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::GetBuf(string b)
{
	buf = b;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::DrawTable()
//  ����������:    ������� ������ ������� � ������ ���� ������������
//  ����:          ���
//  �����:         true - ������� �������� �������
//				   false - ������� ����� � �� ����� ���� ��������
//  ����������:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool XlTable::DrawTable()
{
	if (table_data.empty()) return false;
	string top;
	string it;
	size_t pos;

	if (topit_deliver.length() > 0)	// ���� ������ ����������� ����� ������� ���������, �� ���������� 
	{								// ��� ������ � ����� ( � ��� ��� �������� � ���� [topic]item
		pos = buf.find("]");
		top = buf.substr(1, pos-1);	// top �������� ��� ������
		it = buf.substr(pos+1, buf.length() - 1);	// it �������� ��� �����
	}

	if (data_deliver.length() > 0)	// ���� ������ ����������� ������, �� ������� ���
		std::cout<<data_deliver.c_str()<<"\n";	// ������� ����������� ������

	for (int i = 0; i < row; i++)	// ������� �������
	{
		if (topit_deliver.length()>0)
			std::cout<<top.c_str()<<topit_deliver.c_str()<<it.c_str()<<topit_deliver.c_str();	
		for (int j = 0; j < col; j++)
		{
			std::cout<<table_data[i][j].c_str();
			if ( j != (col - 1)) std::cout<<col_deliver.c_str();
		}
		std::cout<<row_deliver.c_str();
	}
	return true;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::SetDelivers(string row_d, string col_d, string data_d, string topit_d)
//  ����������:    ������������� �����������
//  ����:          row_d - ����������� �����, col_d - ����������� �������,
//				   data_d - ����������� ������ ������, topit_d - ����������� ������ ������ ������
//  �����:         ���
//  ����������:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::SetDelivers(string row_d, string col_d, string data_d, string topit_d)
{
	row_deliver = row_d;
	col_deliver = col_d;
	data_deliver = data_d;
	topit_deliver = topit_d;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::Delete()
//  ����������:    ������� �������
//  ����:          ���
//  �����:         ���
//  ����������:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
void XlTable::Delete()
{
	table_data.clear();
	row = 0;
	col = 0;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           bool XlTable::GetData(HDdeDATA hData)
//  ����������:    �������� ������ �� ������� � �� �� ������ ��������� ������� ������ � �����
//  ����:          ������������� ������� ������ (���������� ������� ��������� ������ ��������)          
//  �����:         0 - ��� �������� ����������, 1 - ��� ������
//  ����������:    �������� ������������
////////////////////////////////////////////////////////////////////////////////////////////////////////////
bool XlTable::GetData(HDdeDATA hData)
{
	 WORD row_n	= 0;
	 WORD col_n = 0;
     DWORD length = 0;                       // ����� ������
     UINT offset  = 0;                       // ������� � ������� ������
     PBYTE data   = NULL;                    // ��������� �� ����������� ������
	 WORD cnum = 0, rnum = 0;                // ���������� ��� �������� ����� � �������� �������
	 WORD cb = 0;                            // ���������� ��� �������� cb
	 WORD size = 0;                          // ���������� ��� �������� ������� ������
	 WORD type = 0;                          // ���������� ��� �������� ���� ������
	 string str;
          
     union                                   // ����������� ��� �������������� ������ � ��������� ����
     {
           WORD w; 
           double d; 
           BYTE b[8];
     } conv;  

     length = DdeGetData(hData, NULL, 0, 0); // ���������� ����� ����������� �� ������� ������
     if (length < 8) return 1;               

     data = new BYTE[length];                // �������� ����� ��������� �����
     DdeGetData(hData,data,length,0);        // �������� ������

     // �������, ����� ������ ���� ��� tdtTable
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
	 if (conv.w != tdtTable) return 1;
   
     // �������, ����� cb ��������� 4
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
	 if (conv.w != 4) return 1;

     // �������� ���������� ����� �������
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
     row_n = conv.w;

     // �������� ���������� �������� �������
     conv.b[0] = data[offset++];
     conv.b[1] = data[offset++];
     col_n = conv.w;

	 if (!col_n || !row_n)                       // ���� ���-�� ����� ��� �������� ����� ����, �� ������
     {
         delete []data;
         return 1;
     }    

	 InitTable(row_n,col_n,"");
  
     // ��������� ��� ������� �������
	 while (offset < length)
	 {
         // ��������� ��� ������  
		 conv.b[0] = data[offset++];
	     conv.b[1] = data[offset++];
		 type = conv.w;

		 switch(type)
		 {
		 case tdtFloat:                      // ������ ������ ���� Double
			 {
	 			conv.b[0] = data[offset++];
				conv.b[1] = data[offset++];
				cb = conv.w;
				if (cb%8) 
				{
					delete []data; 
					return 1;
				}
				for (int tmp=0, i=0; i < cb; i++, tmp++)
				{
					conv.b[tmp] = data[offset++];
					if (tmp == 7)  // ������� ���� �����
					{
						 tmp = -1;
						 str = DoubleToString(conv.d);
						 if (!str.compare("")) 
						 {
							 delete []data;
							 return 1;
						 }
						 
						 table_data[rnum][cnum] = str;
						 cnum++;
						 if (cnum == col)
						 {
							 cnum = 0;
							 rnum++;
						 }
					 }
				}
			 }
			 break;

		 case tdtString:                     // ������ ������ ���� String
			 {
                 // ��������� cb
			     conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;

        		 for (WORD i = 0; i < cb; )
			     {	 // ��������� ����� ������ (��� ��� ������������ ����)
				    size = data[offset++];
					
				    // ��������� ������ ��������� 
				    for(int j = 0; j < size; j++)
                        table_data[rnum][cnum].push_back(data[offset++]);
				    i += 1 + size;
				    cnum++;
				    if (cnum == col)
				    {
					   cnum = 0;
					   rnum++;
			        }
                 }
			 }
			 break;

		 case tdtBool:                       // ������ ������ ���� Bool
			 {
			     conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb%2) 
			     {
				     delete []data; 
				     Delete();
 				     return 1;
		         }
			     for (int tmp=0, i=0; i < cb; i++, tmp++)
			     {
				     conv.b[tmp] = data[offset++];
				     if (tmp == 1)  // ������� ���� �����
				     {
					    tmp = -1;
					    if (conv.w) table_data[rnum][cnum] = "true";
					    else table_data[rnum][cnum] = "false";
					    cnum++;
					    if (cnum == col)
					    {
						   cnum = 0;
						   rnum++;
                        }
                     }
                 }
			 }
			 break;

		 case tdtError:                      // ������ ������ ���� Error
			 {
                 conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb%2) 
			     {
				     delete []data; 
				     Delete();
				     return 1;
                 }
			     for (int tmp=0, i=0; i < cb; i++, tmp++)
			     {
				     conv.b[tmp] = data[offset++];
				     if (tmp == 1)  // ������� ���� �����
				     {
					    tmp = -1;
					    table_data[rnum][cnum] = "error";
					    cnum++;
					    if (cnum == col)
					    {
						    cnum = 0;
						    rnum++;
				        }
				     }
		         }
			 }
			 break;

		 case tdtBlank:                      // ������ ������ ���� Blank
			 {
                 conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb != 2)
			     {
				     delete []data; 
				     Delete();
				     return 1;
		         }
			     conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     size = conv.w;
			     for (int i = 0; i<size; i++)
			     {
				     table_data[rnum][cnum] = "";
				     cnum++;
				     if (cnum == col)
				     {
					     cnum = 0;
					     rnum++;
			         }
	             }
			 }
			 break;

		 case tdtInt:	                     // ������ ������ ���� Int
			 {
				 char mas[20];
                 conv.b[0] = data[offset++];
			     conv.b[1] = data[offset++];
			     cb = conv.w;
			     if (cb%2) 
			     {
				     delete []data; 
				     Delete();
				     return 1;
		         }
			     for (int tmp=0, i=0; i < cb; i++, tmp++)
			     {
				     conv.b[tmp] = data[offset++];
				     if (tmp == 1)  // ������� ���� �����
				     {
					     tmp = -1;
						 _itoa_s(conv.w,mas,10);
					     table_data[rnum][cnum] = mas; 
					     cnum++;
					     if (cnum == col)
					     {
						     cnum = 0;
						     rnum++;
				         }
		             }
			     }
			 }
			 break;

		 case tdtSkip:                       // �� ��������������
		     {
				 delete []data; 
				 Delete();
				 return 1;
             }
		 }
	 }

	 delete []data;                          // ���������� ������ �������  
	 return 0;                               // ������� ��������� �������
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  ���:           XlTable::DoubleToString(double num)
//  ����������:    ����������� ������������ ����� � ������
//  ����:          num - ������������ ����� ���� double
//  �����:         ������-������������� ������������� �����, ���� ������ ������ � ������ ������
//  ����������:
////////////////////////////////////////////////////////////////////////////////////////////////////////////
string XlTable::DoubleToString(double num)
{
string str;
	char buf[_CVTBUFSIZE];
	int decimal;
	int sign;
	int size;
	
	// ����������� �����
	if (_ecvt_s(buf, _CVTBUFSIZE, num, 15, &decimal, &sign))
	{
		return "";
	}
		
	str = buf;

	// ����������� ������ � ������� ������
	if (decimal > 0)
	{
		str.insert(decimal,".");
	}

	if (decimal <= 0)
	{
		decimal = -decimal;
		for(int i = 0; i < decimal; i++)
		{
			str.insert(0,"0");	
		}
		str.insert(0,"0.");
	}

	if (sign < 0)
	{
		str.insert(0,"-");
	}
	
	size = str.size();
	while (str[size-1] == '0')
	{
		str.erase(size-1,1);
		size = str.size();
	}

	size = str.size();
	if (str[size-1] == '.') 
	{
		str.erase(size-1, 1);
		return str;
	}

	return str;
}