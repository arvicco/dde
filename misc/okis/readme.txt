���������� ������� ������� ������������� � ����������� ������������� � ������������ Dde �������
������ �� ������ ���� ��������� Dde ������(�) � ������(��) �� ������� ��� ����� Delphi,VB,VBS,JS,PERL,Python,Rexx,Net
��� ����� ������, ����������� ������������� ���������� ������, ������� � ����� ��������� ������, ��������� ADO,ODBC,�������

����������� ������� �� VBS ��� �������� ������ � ��� ���� ������

� ���� ����������� ActiveX ������ ������, � �����- ��������� ���������� ��� ����� �������� � ������
�������� ��������������

������ �� legkome@mail.ru


�����!!! dde.exe ����� �������� �� 01.07.2010
��� �� ������������� �����, � ��� ����� ���� ������.
��� ������� (vbs,wsc) ������ ��� �������, �� ������ ��
������ ������ ������ ���.
��� �� ��� ��������� � � FINDHTA68.exe

������� ����� legkome@mail.ru

���� � ��� ����� �� ���� �������� ��������� � �������� 
�����
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text
"Format"="Delimited(;)"
��� ��������� ���� jet40.reg


1. ������� ����� (����� SERVER) � ��� ����� ��� ����� info.exe
 
2. ��������������� ���� �����
   (   ��� ���� ���� ������������.vbs. 
   ��� ���� ��� �������, ��������� ��� (2 ���� ������). ���������
   ���������� ����, �� �������� ��� �������� ����, ��������� �������(-�)
   ��� ������ �� ��� (Dde Server=excel). � ���� ����� ������������
   ������ �� ������ �����. )

3. ������������ � ��. �������(2 ���� ������) ���� ..\SERVER\QUIK.UDL
   �� ������� ����������� � 1 �������� �� ..\SERVER\QUIK.MDB
   [��������� �����������]
   [OK]
4. � �������� (notepad.exe) ������� START.VBS
   ��������� �������� (���������� � ['])

5. � �������� (notepad.exe) ������� Dde.VBS
   ��������� �������� (���������� � ['])

6. � �������� (notepad.exe) ������� Dde.WSC
   ��������� �������� (���������� � ['])

7. � �������� (notepad.exe) ������� ANALYS.WSC
   ��������� �������� (���������� � ['])
8. ��������� ��. ��� ����� ������� ���� quik.udl (2 ���� ������)
   ���� �� ������ �������� � ��� ���������� �� ������� 2(Connection)
   ������� [...] � �������� ���� quik.mdb. 
   ���� MS SQL ������ �������� ���������� �� ������� ���������
   � ����� Connection ��������� ����������� ����
   
   [Test Connection] � [OK]

8. ��������� START.VBS (2 ���� ����)
   ��������� ������� �����, �������� ��� �������, �������� ����,
   ��������� � ����e ������� � ��� (�� � start.vbs �����������)
9.  � �������� ..\SERVER\ARXIV\2009\MM\YY ����� ����� csv ��� �������.
    �������� ��, ���������� ��� � ����
   ������ �� ������ �� ����� ������� ������� � ��
   
10 ��������� CSV2SQL.VBS (2 ���� ����) �������� ����� ��� ����� ����� csv.
   ��������� ������ ������� ������� ������ ���� ��� ������ 2� �����.  

11 ��������� � 8 ������ ������ ����� �������� ��� � ��. 
   (�� dde.vbs)



GearBox.qpl - ����� ������ ������� ������ ����� � ��������� ����� ������� �� �����
GearBox.wsc - com ��� ������ � ������� � ��������� �����
ML_EXAMPLE.vbs - ������ ������ GearBox.wsc

TEST.vbs - ��� ������� � ������� ������� � ��������� � �������� ������
           (�������� ������� �����, �������� dde.exe)
ANALYS.vbs - ������ ��� ��������� ������� ANALYS.WSC
           (������ �������� ����� � ������ �������)  


