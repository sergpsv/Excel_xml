--------------------------------------------------------------------
/*  ������ ������������� � PL\SQL ���� �� ����� */
declare
  l_conn  UTL_TCP.connection;
  res     number;
begin
  xml_excel.gpv_sql := 'select * from q1';--'select table_name, temporary, pct_free, num_rows, last_analyzed from user_tables';
  xml_excel.gpv_sheetName := '�����������';
  xml_excel.makeexcelxml;

	-- ������ �������� �� ����� �� myEmail@megafonkavkaz.ru
	res := emailsender.sendemailwithattach('sergey.parshukov@megafonkavkaz.ru','Excel xml test',
	                      to_clob('�������� ������ �������'||chr(13)||chr(10)||'XML ����� ������������ � Excel'),
												'Test_file.xml', xml_excel.gpv_xls_clob);

end;

[1]: (Error): 
ORA-31000: Resource '' is not an XDB schema document 
ORA-06512: at "XDB.DBMS_XMLDOM", line 4564 ORA-06512: at "XDB.DBMS_XMLDOM", line 4578 
ORA-06512: at "SPARSHUKOV.XML_EXCEL", line 379 
ORA-00081: address range [0x60000000000A8700, 0x60000000000A8704) is not readable 
ORA-00600: internal error code, arguments: [17108], [0x000000000], [], [], [], [], [], [] 
ORA-06512: at line 7

--------------------------------------------------------------------
/*  ������ ������������� � PL\SQL ���� �� FTP ������*/
declare
  res number;
  msg varchar2(200);
  l_conn  UTL_TCP.connection;
  
begin
  xml_excel.makeexcelxml('select sysdate from dual','������ 1');
  xml_excel.AddSheet('select * from dual','sheet_2');
  xml_excel.AddSheet('select * from user_tables','��� ��������');
  xml_excel.writetoclob;
  
  l_conn := ftp.login('10.61.14.168', '21', 'ftpuser', 'ftp21gfcc');
  ftp.put_remote_ascii_data(l_conn, 'test2.xml', xml_excel.gpv_xls_clob);
  ftp.logout(l_conn);
  utl_tcp.close_all_connections;  
exception
  when others then
    dbms_output.put_line(SQLCODE ||' '||SQLERRM);
end;

--------------------------------------------------------------------
/*  ������ ������������� � SQL ���� ���������� LOB*/
select xml_excel.getexcelxml_l('select table_name, pct_free, num_rows, last_analyzed from user_tables','��� �������') from dual


--------------------------------------------------------------------
/*  ������ ������������� � SQL ���� ���������� VARCHAR(4000)      */
/*  ���� ��������� ����� 4000 �������� - ���������� XML � ������� */
select xml_excel.getexcelxml_c('select table_name, pct_free, num_rows, last_analyzed from user_tables','��� �������') from dual

