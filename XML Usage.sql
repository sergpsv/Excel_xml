--------------------------------------------------------------------
/*  1 пример использования в PL\SQL коде по почте */
declare
  l_conn  UTL_TCP.connection;
  res     number;
begin
  xml_excel.gpv_sql := 'select sysdate, user from dual';--'select table_name, temporary, pct_free, num_rows, last_analyzed from user_tables';
  xml_excel.gpv_sheetName := 'пример';
  xml_excel.makeexcelxml;

	-- пример отправки по почте на Your_email@mail.ru
	res := emailsender.sendemailwithattach('Your_email@mail.ru', -- адрес
	                                       'Excel xml test',     -- тема письма
	                                       to_clob('тестовый пример посылки'||chr(13)||chr(10)||'XML файла совместимого с Excel'), -- тело письма
										   'Test_file.xml',         -- имя для вложения
										   xml_excel.gpv_xls_clob); -- содержимое файла

end;


--------------------------------------------------------------------
/* 2 пример использования в PL\SQL коде на FTP сервер*/
declare
  res number;
  msg varchar2(200);
  l_conn  UTL_TCP.connection;
  
begin
  xml_excel.makeexcelxml('select sysdate from dual','лист1');
  xml_excel.AddSheet('select * from dual','sheet_2');
  xml_excel.AddSheet('select * from user_tables','мои таблички');
  xml_excel.writetoclob;
  
  l_conn := ftp.login('192.168.1.1', '21', 'ftp_login', 'ftp_password');
  ftp.put_remote_ascii_data(l_conn, 'test2.xml', xml_excel.gpv_xls_clob);
  ftp.logout(l_conn);
  utl_tcp.close_all_connections;  
exception
  when others then
    dbms_output.put_line(SQLCODE ||' '||SQLERRM);
end;

--------------------------------------------------------------------
/* 3 пример использования в SQL коде возвращает LOB*/
select xml_excel.getexcelxml_l('select table_name, pct_free, num_rows, last_analyzed from user_tables','мои таблицы') from dual


--------------------------------------------------------------------
/* 4 пример использования в SQL коде возвращает VARCHAR(4000)      */
/*  если результат более 4000 символов - возвращает XML с ошибкой */
select xml_excel.getexcelxml_c('select table_name, pct_free, num_rows, last_analyzed from user_tables','мои таблицы') from dual

