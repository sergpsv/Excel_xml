CREATE OR REPLACE PACKAGE xml_excel
authid current_user
/*
    author  = sparshukov
    goal    = выгрузка селекта в слов. оформтированный по правилам Excel без 
              привлечения XMLType
*/
is
  gpv_version   constant varchar2(10) := '1.0';

  gpv_sql       varchar2(32000) := '';
  gpv_sheetName varchar2(140)   := '';
  gpv_xls_clob  clob;
  
  function GetVersion return varchar2;

  procedure makeExcelXml;
  procedure makeExcelXml(p_sql varchar2, p_sheetName varchar2);
  procedure AddSheet(p_sql varchar2, p_sheetName varchar2);
  procedure writeToClob;

  function GetExcelXml_l(p_sql varchar2, p_sheetName varchar2) return clob;
  function GetExcelXml_c(p_sql varchar2, p_sheetName varchar2) return varchar2;

end;
/

CREATE OR REPLACE PACKAGE BODY xml_excel
IS
  doc       xmldom.DOMDocument;
  main_node xmldom.DOMNode;  -- весь XML document
  root_node xmldom.DOMNode;  -- первый элемент в списке = содержит весь документ
  root_elmt xmldom.DOMElement; --

  book_node   xmldom.DOMNode;
  styles_node xmldom.DOMNode;
  sheet_node  xmldom.DOMNode;
  table_node  xmldom.DOMNode;
  row_node    xmldom.DOMNode;
  cell_node   xmldom.DOMNode;
  data_node   xmldom.DOMNode;
  pi_node     xmldom.DOMProcessingInstruction;

  user_node xmldom.DOMNode;

  item_node xmldom.DOMNode;
  item_elmt xmldom.DOMElement;
  item_text xmldom.DOMText;

  l_clob    clob;
  pos       pls_integer;
  prevpos   pls_integer;
  width     pls_integer;


--//////////////////////////////////////////////////////////////////////////////
procedure writeToClob
is
begin
  if xmldom.isnull(doc) then
    Raise_application_error(-20001, 'The target XML document is null.');
  end if;

  l_clob := empty_clob(); l_clob :=' ';
  xmldom.writeToClob(doc, l_clob, 'Windows-1251');
  pos := 1; prevpos := 1;
  while pos <> 0
  loop
    pos := dbms_lob.instr(l_clob, 'xmlns=""',pos);
    if pos <>0 then
      dbms_lob.write(l_clob,8,pos,'        ');
    end if;
  end loop;

  gpv_xls_clob := empty_clob();
  gpv_xls_clob := to_clob('<?xml version="1.0" encoding="Windows-1251"?>'||chr(13)||chr(10));
  dbms_lob.append(gpv_xls_clob, l_clob);
  l_clob := empty_clob();

end;

--//////////////////////////////////////////////////////////////////////////////
function GetExcelXml_l(p_sql varchar2, p_sheetName varchar2) return clob
is
begin
  if nvl(p_sql,'') = '' then return to_clob(''); end if;

  if xmldom.isnull(doc) then
    makeExcelXml(p_sql, nvl(p_sheetname,'sheet_1'));
  end if;
  writeToClob;
  return gpv_xls_clob;
end;

--//////////////////////////////////////////////////////////////////////////////
function GetExcelXml_c(p_sql varchar2, p_sheetName varchar2) return varchar2
is
begin
  if nvl(p_sql,'') = '' then return ''; end if;

  if xmldom.isnull(doc) then
    makeExcelXml(p_sql, nvl(p_sheetname,'sheet_1'));
  end if;
  writeToClob;

  if dbms_lob.getlength(gpv_xls_clob)>4000 then
    return '<?xml version="1.0" ?><Result>XML result is too long (more that 4000 symbols)</Result>';
  else
    return substr(to_char(gpv_xls_clob),1,4000);
  end if;
end;


--//////////////////////////////////////////////////////////////////////////////
procedure MakeStyleTable
is
begin
  if xmldom.isnull(doc) then
    Raise_application_error(-20001, 'The target XML document is null.');
  end if;
    item_elmt := xmldom.createElement(doc, 'Styles');
    styles_node := xmldom.appendChild(book_node, xmldom.makeNode(item_elmt));

    -- style N1
    item_elmt := xmldom.createElement(doc, 'Style');
    xmldom.setAttribute(item_elmt, 'ss:ID', 'Default');
    xmldom.setAttribute(item_elmt, 'ss:Name', 'Normal');
    user_node := xmldom.appendChild(styles_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Alignment');
    xmldom.setAttribute(item_elmt, 'ss:Vertical', 'Bottom');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Font');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'NumberFormat');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));

    -- style N2 -- title of columns
    item_elmt := xmldom.createElement(doc, 'Style');
    xmldom.setAttribute(item_elmt, 'ss:ID', 'BoldTitle');
    user_node := xmldom.appendChild(styles_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Alignment');
    xmldom.setAttribute(item_elmt, 'ss:Horizontal', 'Left');
    xmldom.setAttribute(item_elmt, 'ss:Vertical', 'Top');
    xmldom.setAttribute(item_elmt, 'ss:WrapText', '1');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Font');
    xmldom.setAttribute(item_elmt, 'ss:Size', '9');
    xmldom.setAttribute(item_elmt, 'ss:Color', '#000000');
    xmldom.setAttribute(item_elmt, 'ss:Bold', '1');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'NumberFormat');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));

    -- style N3 -- STRING of columns
    item_elmt := xmldom.createElement(doc, 'Style');
    xmldom.setAttribute(item_elmt, 'ss:ID', 'SC');
    user_node := xmldom.appendChild(styles_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Alignment');
    xmldom.setAttribute(item_elmt, 'ss:Horizontal', 'Left');
    xmldom.setAttribute(item_elmt, 'ss:Vertical', 'Top');
    xmldom.setAttribute(item_elmt, 'ss:WrapText', '1');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Font');
    xmldom.setAttribute(item_elmt, 'ss:Size', '8');
    xmldom.setAttribute(item_elmt, 'ss:Color', '#000000');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'NumberFormat');
    xmldom.setAttribute(item_elmt, 'ss:Format', '@');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));

    -- style N4 -- DATE of columns
    item_elmt := xmldom.createElement(doc, 'Style');
    xmldom.setAttribute(item_elmt, 'ss:ID', 'DC');
    user_node := xmldom.appendChild(styles_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Alignment');
    xmldom.setAttribute(item_elmt, 'ss:Horizontal', 'Left');
    xmldom.setAttribute(item_elmt, 'ss:Vertical', 'Top');
    xmldom.setAttribute(item_elmt, 'ss:WrapText', '1');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Font');
    xmldom.setAttribute(item_elmt, 'ss:Size', '8');
    xmldom.setAttribute(item_elmt, 'ss:Color', '#000000');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'NumberFormat');
--RVS 18/08/2008    xmldom.setAttribute(item_elmt, 'ss:Format', 'dd/mm/yyyy');
    xmldom.setAttribute(item_elmt, 'ss:Format', 'General Date');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));

    -- style N5 -- NUMBER of columns
    item_elmt := xmldom.createElement(doc, 'Style');
    xmldom.setAttribute(item_elmt, 'ss:ID', 'NC');
    user_node := xmldom.appendChild(styles_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Alignment');
    xmldom.setAttribute(item_elmt, 'ss:Horizontal', 'Right');
    xmldom.setAttribute(item_elmt, 'ss:Vertical', 'Top');
    xmldom.setAttribute(item_elmt, 'ss:WrapText', '1');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'Font');
    xmldom.setAttribute(item_elmt, 'ss:Size', '8');
    xmldom.setAttribute(item_elmt, 'ss:Color', '#000000');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
    item_elmt := xmldom.createElement(doc, 'NumberFormat');
    item_node := xmldom.appendChild(user_node, xmldom.makeNode(item_elmt));
end;

--//////////////////////////////////////////////////////////////////////////////
function GetCellType(p_type number ) return varchar2
is
begin
  return
    case p_type
      when 1 then 'SC'
      when 2 then 'NC'
      when 12 then 'DC'
    else 'SC' end;
end;

function GetCellDataType(p_type number ) return varchar2
is
begin
  return
    case p_type
      when 1 then 'String'
      when 2 then 'Number'
--      when 12 then 'String' --'DateTime'
--RVS 18/08/2008
      when 12 then 'DateTime'
    else 'String' end;
end;

--//////////////////////////////////////////////////////////////////////////////
procedure AddSheet(p_sql varchar2, p_sheetName varchar2)
is
   l_descTbl		dbms_sql.desc_tab;       -- таблица описаний
   l_rowCounter	number   := 0;
   l_theCursor		number   ;
   l_colCnt			number   :=0;            -- кол-во колонок
   l_status			number   :=0;            -- результат выполнения запроса
   l_colValue		varchar2(500)  :='';     -- значение столбца
   l_colDateValue	date;
   type TColRec is record
   (
     column_idx    pls_integer,
     column_node   xmldom.DOMNode,
     column_maxLen pls_integer := 0
   );

   type TColumnList  is table of TColRec index by pls_integer;
   colList           TColumnList;
begin
  if xmldom.isnull(doc) then
    Raise_application_error(-20001, 'The target XML document is null. Call makeExcelXml in first.');
  end if;

  if colList is not null then colList.delete; end if;

  -- открываем курсор
  l_theCursor := dbms_sql.open_cursor;
  -- анализируем запрос
  dbms_sql.parse(l_theCursor, p_sql, dbms_sql.native);
  -- получаем описание результатов запроса
  dbms_sql.describe_columns(l_theCursor, l_colCnt, l_descTbl);

  item_elmt := xmldom.createElement(doc,'Worksheet');
  xmldom.setAttribute(item_elmt, 'ss:Name', nvl(p_sheetName,'Sheet 1'));
  sheet_node := xmldom.appendChild(book_node, xmldom.makeNode(item_elmt));
    item_elmt  := xmldom.createElement(doc, 'Table');
    table_node := xmldom.appendChild(sheet_node, xmldom.makeNode(item_elmt));
    for i in 1..l_colCnt
    loop
      item_elmt := xmldom.createElement(doc, 'Column');
--      xmldom.setAttribute(item_elmt, 'ss:AutoFitWidth', '0');
      if l_descTbl(i).col_type = 1 then -- varchar2
        width := l_descTbl(i).col_name_len * 10;
      elsif (l_descTbl(i).col_type = 2) then -- number
        width := 80;
      elsif (l_descTbl(i).col_type = 8) then -- long
        width := 120;
      elsif (l_descTbl(i).col_type = 11) then -- rowid
        width := 120;
      elsif (l_descTbl(i).col_type = 12) then -- date
        width := 100;
      else
        width := 70;
      end if;
      xmldom.setAttribute(item_elmt, 'ss:Width', replace(to_char(width),',','.') );
      colList(i).column_idx  := i;
      colList(i).column_node := xmldom.appendChild(table_node, xmldom.makeNode(item_elmt));
      colList(i).column_maxLen := l_descTbl(i).col_name_len;
      -- связываем солбцы курсора с переменной, в которую будем таскать значения
--RVS 18/08/2008 start
      if l_descTbl(i).col_type = 12 then
      	dbms_sql.define_column(l_theCursor, i, l_colDateValue);
      else
--RVS 18/08/2008 end
      	dbms_sql.define_column(l_theCursor, i, l_colValue, 500);
      end if;
    end loop;

    -- формируем заголовок - первая строка Excel
    item_elmt  := xmldom.createElement(doc, 'Row');
    row_node := xmldom.appendChild(table_node, xmldom.makeNode(item_elmt));
    for i in 1..l_colCnt
    loop
      item_elmt  := xmldom.createElement(doc, 'Cell');
      xmldom.setAttribute(item_elmt, 'ss:StyleID', 'BoldTitle');
      cell_node := xmldom.appendChild(row_node, xmldom.makeNode(item_elmt));
        item_elmt  := xmldom.createElement(doc, 'Data');
        xmldom.setAttribute(item_elmt, 'ss:Type', 'String');
        data_node := xmldom.appendChild(cell_node, xmldom.makeNode(item_elmt));
            item_text := xmldom.createTextNode(doc, l_descTbl(i).col_name );
            item_node := xmldom.appendChild(data_node, xmldom.makeNode(item_text));
    end loop;

    -- выполняем запрос
    l_status := dbms_sql.execute(l_theCursor);
    -- извлекаем результаты
	while (dbms_sql.fetch_rows(l_theCursor) > 0 )
	loop
		item_elmt  := xmldom.createElement(doc, 'Row');
		row_node := xmldom.appendChild(table_node, xmldom.makeNode(item_elmt));
		for i in 1..l_colCnt
		loop
			item_elmt  := xmldom.createElement(doc, 'Cell');
			xmldom.setAttribute(item_elmt, 'ss:StyleID', GetCellType(l_descTbl(i).col_type));
			cell_node := xmldom.appendChild(row_node, xmldom.makeNode(item_elmt));
--RVS 18/08/2008 start
			if l_descTbl(i).col_type = 12 then
				dbms_sql.column_value(l_theCursor, i, l_colDateValue);
				if l_colDateValue < to_date('01/01/1900','dd/mm/yyyy') then l_colDateValue:=null; end if;
				if l_colDateValue > to_date('31/12/2999','dd/mm/yyyy') then l_colDateValue:=to_date('31/12/2999','dd/mm/yyyy'); end if;
				l_colValue := to_char(l_colDateValue, 'yyyy-mm-dd"T"hh24:mi:ss');
			else
--RVS 18/08/2008 end
				dbms_sql.column_value(l_theCursor, i, l_colValue);
				l_colValue := trim(l_colValue);
			end if;
         if (l_descTbl(i).col_type <> 12 or (l_descTbl(i).col_type = 12 and l_colDateValue is not null))
         then
				item_elmt  := xmldom.createElement(doc, 'Data');
				xmldom.setAttribute(item_elmt, 'ss:Type', GetCellDataType(l_descTbl(i).col_type));
				data_node := xmldom.appendChild(cell_node, xmldom.makeNode(item_elmt));
				if l_descTbl(i).col_type = 2 then 
					l_colValue := replace(l_colValue,',','.');
				end if;
				if colList(i).column_maxLen < length(l_colValue) then
					colList(i).column_maxLen := length(l_colValue);
				end if;
				item_text := xmldom.createTextNode(doc, l_colValue );
				item_node := xmldom.appendChild(data_node, xmldom.makeNode(item_text));
      	end if;
		end loop;
		l_rowCounter := l_rowCounter +1;
	end loop;

  dbms_sql.close_cursor(l_theCursor);

  for i in 1..l_colCnt
  loop
    case
      when colList(i).column_maxLen < 12 then
          xmldom.setAttribute(xmldom.makeelement(colList(i).column_node), 'ss:Width',
                        replace(to_char(colList(i).column_maxLen * 6),',','.' ));
      when colList(i).column_maxLen < 24 then
          xmldom.setAttribute(xmldom.makeelement(colList(i).column_node), 'ss:Width',
                        replace(to_char(colList(i).column_maxLen * 5.5),',','.' ));
      when colList(i).column_maxLen < 50 then
          xmldom.setAttribute(xmldom.makeelement(colList(i).column_node), 'ss:Width',
                        replace(to_char(colList(i).column_maxLen * 5),',','.' ));
      else
       xmldom.setAttribute(xmldom.makeelement(colList(i).column_node), 'ss:Width', 250);
      end case;
  end loop;

exception
  when others then
    dbms_sql.close_cursor(l_theCursor);
    raise;
end;


--//////////////////////////////////////////////////////////////////////////////
procedure makeExcelXml(p_sql varchar2, p_sheetName varchar2)
is
begin
  gpv_sql       := p_sql;
  gpv_sheetName := p_sheetName;
  makeExcelXml;
end;

--//////////////////////////////////////////////////////////////////////////////
function GetVersion return varchar2
is
begin
  return gpv_version;
end;

--//////////////////////////////////////////////////////////////////////////////
procedure makeExcelXml
is
BEGIN
    if not xmldom.isnull(doc) then
      xmldom.freeDocument(doc);
      gpv_xls_clob := empty_clob();
    end if;

  doc := xmldom.newDOMDocument;
  main_node := xmldom.makeNode(doc);

  pi_node := xmldom.createProcessingInstruction(doc,'mso-application','progid="Excel.Sheet"');
  root_node := xmldom.appendChild(main_node , xmldom.makeNode(pi_node));
--  xmldom.setVersion(doc, '1.0');  xmldom.setCharset(doc, 'Windows-1251');  xmldom.setStandalone(doc, 'yes');

  root_elmt := xmldom.createElement(doc , 'Workbook' );
  xmldom.setAttribute(root_elmt, 'xmlns',     'urn:schemas-microsoft-com:office:spreadsheet');
  xmldom.setAttribute(root_elmt, 'xmlns:o',   'urn:schemas-microsoft-com:office:office');
  xmldom.setAttribute(root_elmt, 'xmlns:x',   'urn:schemas-microsoft-com:office:excel');
  xmldom.setAttribute(root_elmt, 'xmlns:ss',  'urn:schemas-microsoft-com:office:spreadsheet');
  xmldom.setAttribute(root_elmt, 'xmlns:html','http://www.w3.org/TR/REC-html40');

  book_node := xmldom.appendChild(main_node , xmldom.makeNode(root_elmt) );

  MakeStyleTable;
  AddSheet(gpv_sql, gpv_sheetName);

  writeToClob;

exception
  when others then
    dbms_output.put_line('Исключение при makeExcelXml: '||SQLERRM);
/*    if not xmldom.isnull(doc) then
      xmldom.freeDocument(doc);
    end if;*/
    raise;
end;


BEGIN
  gpv_xls_clob := empty_clob();
END;
/
