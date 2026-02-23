// Copyright (C) 2012-2024 Engineer Mareev Enterprises

//	Create создает необходимые поля в структуре ERPData.
Procedure Create(ERPData) Export

	//	Путь до конфигурационного файла ERPEngine.xml.
	ERPData.Insert("Config", "C:\inetpub\ERPWebInterface\Settings\ERPEngine.xml");
	
	//	Конфигурационный файл ERPEngine.xml.
	ERPData.Insert("ERPEngine", "");
	
	//	Загруженный конфигурационный файл ERPEngine.xml.
	ERPData.Insert("DOMDocument", Undefined);
	
	//	COM-объект подключения к SQL-базе данных ADODB.Connection 
	ERPData.Insert("Connection", New COMObject("ADODB.Connection"));
	
	//	Строка подключения к SQL-базе данных 
	ERPData.Insert("ConnStr", "");
	
	//	Уровень вложенности трнзакции
	ERPData.Insert("TransLevel", 0);
	
	//	Параметры сообщения
	ERPData.Insert("Source", "");
	ERPData.Insert("Target", "");
	ERPData.Insert("Name", "");

	//	Таблица заголовков и дочерние таблицы
	ERPData.Insert("HeaderTable", New ValueTable);
	ERPData.Insert("ChildTables", New Array);
	
	//	Строка данных таблицы заголовков
	ERPData.Insert("header", Undefined);
	
	//	Номер теущий строки таблицы заголовов
	ERPData.Insert("HeaderLine", -1);
	
	//	Краткие имена дочерних таблиц
	ERPData.Insert("ChildNames", New Array);
	
	//	Индекс дочерней таблицы
	ERPData.Insert("ChildIndex", -1);
	
	//	Номера текущих строк дочерних таблиц
	ERPData.Insert("ChildLines", New Array);
	
	//	Полное имя таблицы заголовка и дочерних таблиц
	ERPData.Insert("HeaderTableName", "");
	ERPData.Insert("ChildTableNames", New Array);
	
	//	Индексы колонок заголовка с контрольной информацией (кол-во строк в дочерней таблице)
	ERPData.Insert("TotalIndexes", New Array);
	
	//	#AMP Попов 10.12.2024 В режиме экспорта уникальных сообщений, старые сообщения можно не удалять,
	//	а это очень хорошо, так как команда DELETE лочит всю таблицы. Предусмотрим для этого специальный флаг.
	//	Этот флаг в некоторых сообщениях можно сбрасывать (например, сообщение REMAINS).
	ERPData.Insert("IsSkipDeleteOldMessages", False);
	
EndProcedure

//	ReadConfig читает конфигурационный файл
Function ReadConfig(ERPData) Export
	
	DOMBuilder = New DOMBuilder;
	XMLReader = New XMLReader;
	
	Try
		If ERPData.ERPEngine = "" ИЛИ ERPData.ERPEngine = Неопределено Then
			XMLReader.OpenFile(ERPData.Config);
		Else
			XMLReader.SetString(ERPData.ERPEngine);
		EndIf;
	 	ERPData.DOMDocument = DOMBuilder.Read(XMLReader);
		XMLReader.Close();
	Except
		Return "Ошибка чтения конфигурационного файла <" + ERPData.Config + ">:" + Chars.CR + Chars.LF + ErrorDescription();
	EndTry;
	
	EngineNode = ERPData.DOMDocument.DocumentElement;
	
	If EngineNode.NodeName <> "ERPEngine" Then
		Return "Недопустимый формат описателя интерфейса ERP-WMS (отсутствует корневой элемент ERPEngine)";
	EndIf;
	
	Return "";

EndFunction

//	Connect подключает модуль к SQL-базе данных.
Function Connect(ERPData) Export
	
	Error = ReadConfig(ERPData);
	If Error <> "" Then
		Return Error;
	EndIf;

	EngineNode = ERPData.DOMDocument.DocumentElement;
	
	If EngineNode.NodeName <> "ERPEngine" Then
		Return "Недопустимый формат описателя интерфейса ERP-WMS (отсутствует корневой элемент ERPEngine)";
	EndIf;
	
	//	Переберем атрибуты интерфейса ERP-EME.WMS
	For Each EngineAttrib In EngineNode.Attributes Do
		If EngineAttrib.Name = "connection" Then
			ERPData.ConnStr = Mid(EngineAttrib.Value, StrLen("ODBC:") + 1);
		EndIf
	EndDo;
	
	Try
		//	Открываем соединение
		ERPData.Connection.Open(ERPData.ConnStr);
	Except
		Return "Ошибка подключения к базе данных <" + ERPData.ConnStr + ">:" + Chars.CR + Chars.LF + ErrorDescription();
	EndTry;
		
	//	Устанавливаем уровень изоляции adXactReadCommitted
	ERPData.Connection.IsolationLevel = 4096;
	
	Return "";
	
EndFunction

//	Disconnect отключает модуль к SQL-базе данных.
Procedure Disconnect(ERPData) Export
	
	//	Закрываем соединение
	ERPData.Connection.Close();
	
EndProcedure

//  BeginExport начинает экспорт сообщения <Name> c источником <Source> и приемником <Target>.
Procedure BeginExport(ERPData, Val Source, Val Target, Val Name) Export
	
	Markup(ERPData, Source, Target, Name);
	ERPData.TransLevel = 0;

	 // Проиницализируем текущие строки
 	ERPData.HeaderLine = -1;
 	For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
 		ERPData.ChildLines.Add();
  		ERPData.ChildLines[ChildIndex] = -1;
	EndDo;

EndProcedure

//  CommitExport фиксирует изменения после экспорта сообщений в SQL-базе данных.
Procedure CommitExport(ERPData) Export
	
	If ERPData.ConnStr = "" Then
		Return;
	EndIf;
			
	//	Начинаем Sql-транзакцию
	ERPData.TransLevel = ERPData.Connection.BeginTrans();

	//	Удаляем старые сообщения
	DeleteOldMessages(ERPData);
		
	//	Добавляем новые сообщения
	InsertNewMessages(ERPData);

	//	Завершаем Sql-транзакцию
	ERPData.Connection.CommitTrans();
	ERPData.TransLevel = 0;
	
EndProcedure

//  RollbackExport откатывает изменения после экспорта сообщений в SQL-базе данных.
Procedure RollbackExport(ERPData) Export
	
	If ERPData.ConnStr = "" Then
		Return;
	EndIf;
	
	//	Откатываем Sql-транзакцию
	If ERPData.TransLevel > 0 Then
		ERPData.Connection.RollbackTrans();
		ERPData.TransLevel = 0;
	EndIf

EndProcedure

//  UploadToXml выгружает сообщения в XML.
Procedure UploadToXML(ERPData, XMLWriter) Export
	
	XMLWriter.SetString("windows-1251");

	XMLWriter.WriteXMLDeclaration();
	XMLWriter.WriteStartElement("ERPMessage");
	XMLWriter.WriteAttribute("source", ERPData.Source);
	XMLWriter.WriteAttribute("target", ERPData.Target);
	XMLWriter.WriteAttribute("name", ERPData.Name);
	
	XMLWriter.WriteStartElement("ERPTable");
	XMLWriter.WriteAttribute("name", ERPData.HeaderTableName);
	XMLWriter.WriteAttribute("type", "header");
	UploadTableToXML(ERPData, XMLWriter, ERPData.HeaderTable);
	XMLWriter.WriteEndElement();
	
	For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
		XMLWriter.WriteStartElement("ERPTable");
		XMLWriter.WriteAttribute("name", ERPData.ChildTableNames[ChildIndex]);
		XMLWriter.WriteAttribute("type", "child");
		XMLWriter.WriteAttribute("child", ERPData.ChildNames[ChildIndex]);
		UploadTableToXML(ERPData, XMLWriter, ERPData.ChildTables[ChildIndex]);
		XMLWriter.WriteEndElement();
	EndDo;
	
	XMLWriter.WriteEndElement();
	
EndProcedure

Procedure UploadTableToXML(ERPData, XMLWriter, Table)
	
	XMLWriter.WriteStartElement("ERPFields");
	For ColumnIndex = 0 To Table.Columns.Count() - 1 Do
		XMLWriter.WriteStartElement("ERPField");
		Name = Table.Columns[ColumnIndex].Name;
		Type = Table.Columns[ColumnIndex].ValueType;
		XMLWriter.WriteAttribute("name", Name);
		XMLWriter.WriteAttribute("type", GetSqlType(Type));
		XMLWriter.WriteEndElement();
	EndDo;
	XMLWriter.WriteEndElement();
	
	XMLWriter.WriteStartElement("ERPLines");
	For Each TableRow In Table Do
		XMLWriter.WriteStartElement("ERPLine");
		For ColumnIndex = 0 To Table.Columns.Count() - 1  Do
			
			Name = Table.Columns[ColumnIndex].Name;
			Value = TableRow[ColumnIndex];
			Type = Table.Columns[ColumnIndex].ValueType;
			
			XMLWriter.WriteAttribute(Name, AsString(Value, Type));
			
		EndDo;
		XMLWriter.WriteEndElement();
	EndDo;
	XMLWriter.WriteEndElement();
	
EndProcedure

//  AppendHeaderLine добавляет строку в таблицу заголовка сообщения.
Procedure AppendHeaderLine(ERPData) Export
	
	ERPData.header = ERPData.HeaderTable.Add();
	ERPData.header.id = "";
	ERPData.header.what_to_do = "MOD";
	ERPData.header.state = "NEW";
	ERPData.header.created_at = CurrentDate();
	ERPData.header.processed_at = '00010101000000';
	ERPData.header.error_code = "";
	
	//	Обнулим контрольную информацию по количеству строк в дочерних таблицах
	For Each TotalIndex In ERPData.TotalIndexes Do
		ERPData.header[TotalIndex] = 0;
	EndDo
	
EndProcedure

//  SelectChild задает в качестве текущей дочерней таблицы сообщения таблицу <ChildName>.
Procedure SelectChild(ERPData, Val ChildName) Export
	
	ChildIndex = ERPData.ChildNames.Find(ChildName);
	
	If ChildIndex = Undefined Then
		Raise "В описателе интерфейса ERP-WMS отсутствует таблица " + ChildName;
	EndIf;
	
	ERPData.ChildIndex = ChildIndex;
	ERPData.ChildLines[ChildIndex] = -1;
	
EndProcedure

//  AppendChildLine добавляет строку в текущую дочернюю таблицу сообщения. 
Procedure AppendChildLine(ERPData) Export

	ChildIndex = ERPData.ChildIndex;
	If ChildIndex < 0 Then
		Raise "Не задана дочерняя таблица";
	EndIf;
	
	ChildRow = ERPData.ChildTables[ChildIndex].Add();
	ERPData.Insert(ERPData.ChildNames[ChildIndex], ChildRow);
	
	ChildRow.id = "";
	ChildRow.header_id = ERPData.header.id;
    ChildRow.error_code = "";
	
	//	Увеличим на 1 контрольную информацию по кол-ву строк в дочерней таблице
	TotalIndex = ERPData.TotalIndexes[ChildIndex];
	ERPData.header[TotalIndex] = ERPData.header[TotalIndex] + 1;
	
EndProcedure

//  BeginImport начинает импорт сообщения <Name> c источником <Source> и приемником <Target>.
Procedure BeginImport(ERPData, Val Source, Val Target, Val Name, Val HeaderIds = Null) Export
	
	Markup(ERPData, Source, Target, Name);
	ERPData.TransLevel = 0;

	//	Начинаем Sql-транзакцию
	ERPData.TransLevel = ERPData.Connection.BeginTrans();
	
	//	Прочитаем заголовки сообщений
	QueryText = "SELECT * FROM """ + ERPData.HeaderTableName + """ WHERE (""state""='NEW' OR ""state""='WRN')";
	If HeaderIds <> Null Then
		QueryText = QueryText + " AND ""id"" IN(";
		For Each HeaderId In HeaderIds Do
			HeaderId = StrReplace(HeaderId, "'", "''");
			If Right(QueryText, 1) <> "(" Then
				QueryText = QueryText + ", ";
			EndIf;
			QueryText = QueryText + "'" + HeaderId + "'";
		EndDo;
		QueryText = QueryText + ")";
	EndIf;
	//	#AMP Попов 10.12.2024 Так как в EME.WMS возможен режим уникальных сообщений
	//	(к примеру, на один товар может быть несколько сообщений GOODS),
	//  импортировать сообщнеия будем по хронологии (FIFO).
	//QueryText = QueryText + " ORDER BY ""id"";";
	QueryText = QueryText + " ORDER BY ""created_at"" ASC, ""id"";";
	
	SelectLines(ERPData.Connection, QueryText, ERPData.HeaderTable);
		
	//	Проиницализируем текущие строки
	ERPData.HeaderLine = -1;
	For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
		ERPData.ChildLines.Add();
		ERPData.ChildLines[ChildIndex] = -1;
	EndDo;
	
EndProcedure

//  CommitImport фиксирует изменения после импорта сообщений в SQL-базе данных.
Procedure CommitImport(ERPData) Export
	
	//	Обновим последнюю строку
	If ERPData.HeaderLine <> -1 Then
		UpdateMessage(ERPData);
	EndIf;
	
	//	Завершаем Sql-транзакцию
	ERPData.Connection.CommitTrans();
	ERPData.TransLevel = 0;
	
EndProcedure

//  RollbackImport откатывает изменения после импорта сообщений в SQL-базе данных.
Procedure RollbackImport(ERPData) Export
	
	//	Откатываем Sql-транзакцию
	If ERPData.TransLevel > 0 Then
		ERPData.Connection.RollbackTrans();
		ERPData.TransLevel = 0;
	EndIf

EndProcedure

//  NextHeaderLine переставляет указатель на следующую строку в таблице заголовка сообщения.
Function NextHeaderLine(ERPData) Export
	
	If ERPData.HeaderLine + 1 < ERPData.HeaderTable.Count() Then
		
		If ERPData.HeaderLine <> -1 Then
			UpdateMessage(ERPData);
		EndIf;
		
		ERPData.HeaderLine = ERPData.HeaderLine + 1;
		
		ERPData.header = ERPData.HeaderTable[ERPData.HeaderLine];
		
		//	Прочитаем строки сообщений
		For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
			
			SelectLines(
				ERPData.Connection,
				"SELECT * FROM """ + ERPData.ChildTableNames[ChildIndex] + """ WHERE ""header_id"" = '" + ERPData.header.id + "';",
				ERPData.ChildTables[ChildIndex]);
		EndDo;
			
		For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
			
			//	Проверим контрольную информацию
			If ERPData.ChildTables[ChildIndex].Count() <> ERPData.HeaderTable[ERPData.HeaderLine][ERPData.TotalIndexes[ChildIndex]] Then
				ErrorHeader(ERPData, "TOTAL" + (ChildIndex + 1));
			EndIf;

			ERPData.ChildLines[ChildIndex] = -1;
		EndDo;

		//	#AMP Попов 10.12.2024 Проверим, а не выгрузили ли нам новое сообщение?
		CreatedAtTable = New ValueTable;
		AddField(CreatedAtTable, "datetime", "created_at");

		SelectLines(ERPData.Connection,
			"SELECT ""created_at"" FROM """ + ERPData.HeaderTableName + """ WHERE ""id"" = '" + ERPData.header.id + "';",
			CreatedAtTable);
			
		If CreatedAtTable.Count() = 0 Then
			Raise "После загрузки удалено сообщение с id=" + ERPData.header.id;
		Else
			CreatedAt = CreatedAtTable.Get(0).created_at;
			If CreatedAt <> ERPData.header.created_at Then
				Raise "Для сообщения с id=" + ERPData.header.id + " не совпадает created_at до и после загрузки сообщения: " +
                ERPData.header.created_at + "->" + CreatedAt;
		 	EndIf;
		EndIf;
		
		Return True;
	Else
		Return False;
	EndIf

EndFunction

//  NextChildLine переставляет указатель на следующую строку в текущей дочерней таблице сообщения.
Function NextChildLine(ERPData) Export
	
	ChildIndex = ERPData.ChildIndex; 
	ChildName = ERPData.ChildNames[ChildIndex];
	ChildLine = ERPData.ChildLines[ChildIndex];
	
	If ChildLine + 1 < ERPData.ChildTables[ChildIndex].Count() Then
		ERPData.ChildLines[ChildIndex] = ChildLine + 1;
		ERPData.Insert(ChildName, ERPData.ChildTables[ChildIndex][ERPData.ChildLines[ChildIndex]]);
		
		Return True;
	Else
		Return False;
	EndIf

EndFunction

//  ErrorHeader фиксирует ошибку <ErrorCode> в текущей строке таблицы заголовка сообщения.
Procedure ErrorHeader(ERPData, Val ErrorCode) Export
	
	//If HasErrors(ERPData) Then
	//	Return
	//EndIf;
		
	ERPData.header.state = "ERR";
	ERPData.header.processed_at = CurrentDate();
	ERPData.header.error_code = ErrorCode;
	
EndProcedure

//  ErrorChild фиксирует ошибку <ErrorCode> в текущей строке текущей дочерней таблице заголовка сообщения.
Procedure ErrorChild(ERPData, Val ErrorCode) Export
	
	//If HasErrors(ERPData) Then
	//	Return
	//EndIf;
		
	ChildIndex = ERPData.ChildIndex;
	ErrorHeader(ERPData, "CHILD" + (ChildIndex + 1));
	ERPData.ChildTables[ChildIndex][ERPData.ChildLines[ChildIndex]].error_code = ErrorCode;
		
EndProcedure

//  HasErrors возвращает true, если при импорте сообщения были зафиксированы ошибки, иначе возвращает false.
Function HasErrors(ERPData) Export
	
	Return ERPData.header.state = "ERR";
	
EndFunction

//  WarningHeader фиксирует предупреждение <WarningCode> в текущей строке таблицы заголовка сообщения.
Procedure WarningHeader(ERPData, Val WarningCode, Val WarningMore = "") Export
	
	If HasErrors(ERPData) Then
		Return
	EndIf;
		
	ERPData.header.state = "WRN";
	ERPData.header.processed_at = CurrentDate();
	ERPData.header.error_code = WarningCode;
	ERPData.header.error_more = WarningMore;
	
EndProcedure

//  WarningChild фиксирует предупреждение <WarningCode> в текущей строке текущей дочерней таблице заголовка сообщения.
Procedure WarningChild(ERPData, Val WarningCode) Export
	
	If HasErrors(ERPData) Then
		Return
	EndIf;
		
	ChildIndex = ERPData.ChildIndex;
	WarningHeader(ERPData, "CHILD" + (ChildIndex + 1));
	ERPData.ChildTables[ChildIndex][ERPData.ChildLines[ChildIndex]].error_code = WarningCode;
		
EndProcedure

//  HasWarnings возвращает true, если при импорте сообщения были зафиксированы предупреждения, иначе возвращает false.
Function HasWarnings(ERPData) Export
	
	Return ERPData.header.state = "WRN";
	
EndFunction

//  Success явно подтверждает удачный импорт сообщений
Procedure Success(ERPData) Export
	
	If HasErrors(ERPData) Then
		Raise "Нельзя подтвердить успешный импорт сообщения в котором уже есть ошибки";
	EndIf;
	
	ERPData.header.state = "OK";
	ERPData.header.processed_at = CurrentDate();
	
EndProcedure

//  UndoImport откатывает импорт сообщений <Name> c источником <Source> и приемником <Target> в идентификаторами HeaderIds.
Function UndoImport(ERPData, Val Source, Val Target, Val Name, HeaderIds) Export
	
	Markup(ERPData, Source, Target, Name);
	ERPData.Connection.BeginTrans();
	Try
		For Each HeaderId In HeaderIds Do
			UndoImportCore(ERPData, HeaderId);
		EndDo;
		ERPData.Connection.CommitTrans();
	Except
		ERPData.Connection.RollbackTrans();
		Return ErrorDescription();
	EndTry;
	
	Return "";

EndFunction

//  UndoImportCore откатывает импорт сообщения <HeaderId>.
Procedure UndoImportCore(ERPData, Val HeaderId)
	
	HeaderId = StrReplace(HeaderId, "'", "''");
    UndoCommand = ReplaceMacro(ERPData, "UPDATE [$HEADER$] SET [$STATE$]='NEW', [$PROCESSED_AT$]=NULL, [$ERROR_CODE$]=NULL WHERE [$ID$]='" + HeaderId + "';");
    ERPData.Connection.Execute(UndoCommand);

	For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
		ERPData.ChildIndex = ChildIndex;
		UndoCommand = ReplaceMacro(ERPData, "UPDATE [$CHILD$] SET [$ERROR_CODE$]=NULL WHERE [$HEADER_ID$]='" + HeaderId + "';");
		ERPData.Connection.Execute(UndoCommand);
    EndDo

EndProcedure

//	GetStatistics заполняет таблицу <Table> статистикой сообщений.
Procedure GetStatistics(ERPData, Table) Export

	Table.Columns.Clear();
	AddField(Table, "varchar(4)", "source");
	AddField(Table, "varchar(4)", "target");
	AddField(Table, "varchar(32)", "name");
	AddField(Table, "varchar(4)", "state");
	AddField(Table, "integer", "count");
	
	EngineNode = ERPData.DOMDocument.DocumentElement;
	
	//	Переберем сообщения интерфейса ERP-EME.WMS
	For Each MessageNode In EngineNode.ChildNodes Do
		
		If MessageNode.NodeName <> "ERPMessage" Then
			Continue
		EndIf;
		
		MessageName = "";
		MessageSource = "";
		MessageTarget = "";
		
		//	Переберем атрибуты сообщения интерфейса ERP-EME.WMS
		For Each MessageAttrib In MessageNode.Attributes Do
			
			If MessageAttrib.Name = "name" Then
				MessageName = MessageAttrib.Value;
			ElsIf MessageAttrib.Name = "source" Then
				MessageSource = MessageAttrib.Value;
			ElsIf MessageAttrib.Name = "target" Then
				MessageTarget = MessageAttrib.Value;
			EndIf
			
		EndDo;
		
		GetMessageCount(ERPData, MessageSource, MessageTarget, MessageName, "NEW", Table);
		GetMessageCount(ERPData, MessageSource, MessageTarget, MessageName, "WRN", Table);
		GetMessageCount(ERPData, MessageSource, MessageTarget, MessageName, "ERR", Table);
		GetMessageCount(ERPData, MessageSource, MessageTarget, MessageName, "OK", Table);
		
	EndDo;
	
EndProcedure

//	GetMessageCount возвращает количество сообщений.
Procedure GetMessageCount(ERPData, Val Source, Val Target, Val Name, Val State, Table)
	
	CountTable = New ValueTable;
	AddField(CountTable, "integer", "count");
	
	SqlTableName = Source + "_" + Target + "_" + Name + "_header";
	SelectLines(
		ERPData.Connection,
		"SELECT COUNT(id) AS count FROM """ + SqlTableName + """ WHERE ""state""='" + State + "';",
		CountTable);
		
	CountRow = CountTable[0];
	Count = CountRow.count;
	If Count <> 0 Then
		Row = Table.Add();
		Row.source = Source;
		Row.target = Target;
		Row.name = Name;
		Row.state = State;
		Row.count = Count;
	EndIf
	
EndProcedure

//	GetHeaders заполняет таблицу <Table> заголовками сообщения <Name> c источником <Source> и приемником <Target> на статусе <State>.
Procedure GetHeaders(ERPData, Val Source, Val Target, Val Name, Val State, Table) Export

	MessageNode = GetMessageNode(ERPData, Source, Target, Name);
	TableNode = GetTableNode(MessageNode, "header");
	MarkupTable(TableNode, Table);
	
	//	Имя SQL-таблицы заголовков сообщений
	SqlTableName = TableNode.Attributes.GetNamedItem("name").Value;
	
	//	Сформируем условие отбора по статусу сообщения
	StateCondition = "";
	States = StrReplace(State, ",", Chars.LF);
	For Index = 1 To StrLineCount(States) Do
		If StateCondition <> "" Then
			StateCondition = StateCondition + " OR ";
		EndIf;
		
        StateCondition = StateCondition + """state""='" + StrGetLine(States, Index) + "'";
	EndDo;
		
	//	Прочитаем заголовки сообщений
	SelectLines(
		ERPData.Connection,
		"SELECT * FROM """ + SqlTableName + """ WHERE " + StateCondition + " ORDER BY ""id"";",
		Table);
	
EndProcedure

//	GetChilds заполняет массив <ChildsArray> заголовками сообщения <Name> c источником <Source> и приемником <Target>.
Procedure GetChilds(ERPData, Val Source, Val Target, Val Name, ChildsArray) Export

	MessageNode = GetMessageNode(ERPData, Source, Target, Name);
	
	//	Переберем таблицы сообщения интерфейса ERP-EME.WMS
	For Each TableNode In MessageNode.ChildNodes Do
		
		If TableNode.NodeName <> "ERPTable" Then
			Continue
		EndIf;

		TableType = "";
		TableChild = "";
		
		//	Переберем атрибуты таблицы сообщения интерфейса ERP-EME.WMS
		For Each TableAttrib In TableNode.Attributes Do
			If TableAttrib.Name = "type" Then
				TableType = TableAttrib.Value;
			ElsIf TableAttrib.Name = "child" Then
				TableChild = TableAttrib.Value;
			EndIf
		EndDo;
		
		If (TableType = "child") Then
			ChildsArray.Add(TableChild);
		EndIf
		
	EndDo;
	
EndProcedure

//	GetChildLines заполняет таблицу <Table> строками <Child> сообщения <Name> с идентификатором <HeaderId> c источником <Source> и приемником <Target> на статусе <State>.
Procedure GetChildLines(ERPData, Val Source, Val Target, Val Name, Val Child, Val HeaderId, Table) Export

	MessageNode = GetMessageNode(ERPData, Source, Target, Name);
	TableNode = GetTableNode(MessageNode, Child);
	MarkupTable(TableNode, Table);
	
	//	Имя SQL-таблицы заголовков сообщений
	SqlTableName = TableNode.Attributes.GetNamedItem("name").Value;
	
	//	Прочитаем строки сообщения
	SelectLines(
		ERPData.Connection,
		"SELECT * FROM """ + SqlTableName + """ WHERE ""header_id""='" + StrReplace(HeaderId, "'", "''") + "';",
		Table);
	
EndProcedure

//  Markup размечает набор данных под сообщение <Name> c источником <Source> и приемником <Target>.
Procedure Markup(ERPData, Val Source, Val Target, Val Name);
	
	ERPData.Source = Source;
	ERPData.Target = Target;
	ERPData.Name = Name;

	//	Очистим структуру ERPData от предыдущей разметки
	ERPData.HeaderTable.Columns.Clear();
	ERPData.ChildTables.Clear();
	ERPData.header = Undefined;
	ERPData.HeaderLine = -1;
	For Each ChildName In ERPData.ChildNames Do
		ERPData.Delete(ChildName);
	EndDo;
	ERPData.ChildNames.Clear();
	ERPData.ChildIndex = -1;
	ERPData.ChildLines.Clear();
	ERPData.HeaderTableName = "";
	ERPData.ChildTableNames.Clear();
	ERPData.TotalIndexes.Clear();

	MessageNode = GetMessageNode(ERPData, Source, Target, Name);
	
	ChildCount = 0;
			
	//	Переберем таблицы сообщения интерфейса ERP-EME.WMS
	For Each TableNode In MessageNode.ChildNodes Do
		
		If TableNode.NodeName <> "ERPTable" Then
			Continue
		EndIf;

		TableName = "";
		TableType = "";
		TableChild = "";
		
		//	Переберем атрибуты таблицы сообщения интерфейса ERP-EME.WMS
		For Each TableAttrib In TableNode.Attributes Do
			If TableAttrib.Name = "name" Then
				TableName = TableAttrib.Value;
			ElsIf TableAttrib.Name = "type" Then
				TableType = TableAttrib.Value;
			ElsIf TableAttrib.Name = "child" Then
				TableChild = TableAttrib.Value;
			EndIf
		EndDo;
		
		If TableType = "header" Then
			ERPData.HeaderTableName = TableName;
			Table = ERPData.HeaderTable;
		ElsIf TableType = "child" Then
			ERPData.ChildTables.Add(New ValueTable);
			ERPData.ChildNames.Add(TableChild);
			ERPData.ChildTableNames.Add(TableName);
			ERPData.Insert(TableChild, Undefined);
			Table = ERPData.ChildTables[ChildCount];
			ChildCount = ChildCount + 1;
		Else
			Continue
		EndIf;
		
		MarkupTable(TableNode, Table);
		
	EndDo;
	
	//	Найдем в таблице заголовков колонки с итоговыми количествами строк в дочерних таблицах
	For Each ChildName In ERPData.ChildNames Do
		TotalColumnName = "total_" + ChildName;
		
		TotlaIndex = Undefined;
		For ColumnIndex = 0 To ERPData.HeaderTable.Columns.Count() - 1 Do
			If ERPData.HeaderTable.Columns[ColumnIndex].Name = TotalColumnName Then
				TotlaIndex = ColumnIndex;
				Break
			EndIf
		EndDo;
		
		If TotlaIndex = Undefined Then
			Raise "В таблице заголовков отсутствует колонка " + TotalColumnName + " с контрольной информацией и кол-ве строк в дочерней таблиц";
		EndIf;
		
		ERPData.TotalIndexes.Add(TotlaIndex);
	EndDo;
	
EndProcedure

//  GetMessageNode возвращает XML-описатель сообщения <Name> c источником <Source> и приемником <Target>.
Function GetMessageNode(ERPData, Val Source, Val Target, Val Name)
	
	EngineNode = ERPData.DOMDocument.DocumentElement;
	
	//	Переберем сообщения интерфейса ERP-EME.WMS
	For Each MessageNode In EngineNode.ChildNodes Do
		
		If MessageNode.NodeName <> "ERPMessage" Then
			Continue
		EndIf;
		
		MessageName = "";
		MessageSource = "";
		MessageTarget = "";
		
		//	Переберем атрибуты сообщения интерфейса ERP-EME.WMS
		For Each MessageAttrib In MessageNode.Attributes Do
			
			If MessageAttrib.Name = "name" Then
				MessageName = MessageAttrib.Value;
			ElsIf MessageAttrib.Name = "source" Then
				MessageSource = MessageAttrib.Value;
			ElsIf MessageAttrib.Name = "target" Then
				MessageTarget = MessageAttrib.Value;
			EndIf
			
		EndDo;
		
		If MessageName = Name And MessageSource = Source And MessageTarget = Target Then
			Return MessageNode;
		EndIf
	EndDo;
	
	Raise "В описателе интерфейса ERP-WMS отсутствует сообщение " + Name + " (откуда-" + Source + ", куда-" + Target + ")";
	Return Undefined;
	
EndFunction

//  GetMessageNode возвращает XML-описатель таблицы <TableName> в XML-описателе сообщения <MessageNode>.
Function GetTableNode(MessageNode, Val TableName)
	
	//	Переберем таблицы сообщения интерфейса ERP-EME.WMS
	For Each TableNode In MessageNode.ChildNodes Do
		
		If TableNode.NodeName <> "ERPTable" Then
			Continue
		EndIf;

		TableType = "";
		TableChild = "";
		
		//	Переберем атрибуты таблицы сообщения интерфейса ERP-EME.WMS
		For Each TableAttrib In TableNode.Attributes Do
			If TableAttrib.Name = "type" Then
				TableType = TableAttrib.Value;
			ElsIf TableAttrib.Name = "child" Then
				TableChild = TableAttrib.Value;
			EndIf
		EndDo;
		
		If (TableType = "header" И TableName = "header") Или (TableType = "child" И TableName = TableChild) Then
			Return TableNode;
		EndIf
		
	EndDo;
	
	Raise "В описателе интерфейса ERP-WMS отсутствует таблица " + TableName;
	Return Undefined;
	
EndFunction

//  MarkupTable по XML-описателю таблицы <TableNode> размечает таблцу <Table>.
Procedure MarkupTable(TableNode, Table)
	
	//	Переберем поля таблицы сообщения интерфейса ERP-EME.WMS
	For Each FieldNode In TableNode.ChildNodes Do
		
		If FieldNode.NodeName <> "ERPField" Then
			Continue;
		EndIf;

		FieldName = "";
		FieldType = "";
		
		//	Переберем атрибуты поля таблицы сообщения интерфейса ERP-EME.WMS
		For Each FieldAttrib In FieldNode.Attributes Do
			If FieldAttrib.Name = "name" Then
				 FieldName = FieldAttrib.Value;
			ElsIf FieldAttrib.Name = "type" Then
				 FieldType = FieldAttrib.Value;
			EndIf
		EndDo;
		
		AddField(Table, FieldType, FieldName);
		
	EndDo
	
EndProcedure
	
//	DeleteOldMessages удаляет старые сообщения.
Procedure DeleteOldMessages(ERPData)
	
    //	#AMP Попов 10.12.2024 В режиме экспорта уникальных сообщений старые сообщения не удаляем
	If ERPData.IsSkipDeleteOldMessages Then
		Return;
	EndIf;
	
	For Each TableRow In ERPData.HeaderTable Do
		
		//	Удаляем старый заголовок сообщения
		ERPData.Connection.Execute("DELETE FROM """ + ERPData.HeaderTableName + """ WHERE ""id""='" + TableRow.id + "';");
		
		//	Удаляем старые строки сообщения
		For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
			ERPData.Connection.Execute("DELETE FROM """ + ERPData.ChildTableNames[ChildIndex] + """ WHERE ""header_id""='" + TableRow.id + "';");
		EndDo
		
	EndDo;
		
EndProcedure

//	InsertNewMessages добавляет новые сообщения.
Procedure InsertNewMessages(ERPData)
	
	//	Добавляем новые заголовки сообщений
	InsertLines(ERPData.Connection, ERPData.HeaderTableName, ERPData.HeaderTable);
	
	//	Добавляем новые строки сообщений
	For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
		InsertLines(ERPData.Connection, ERPData.ChildTableNames[ChildIndex], ERPData.ChildTables[ChildIndex]);
	EndDo

EndProcedure

//  UpdateMessage обновляет проимпортированные сообщения.
Procedure UpdateMessage(ERPData)
	
	//	Таблица с ответом в заголовке сообщения
	AnswerHeader = New ValueTable;
	AddField(AnswerHeader, "varchar(36)", "id");
	AddField(AnswerHeader, "varchar(4)", "state");
	//	#AMP Попов 10.12.2024 Используем поле created_at в ключах ответа, чтобы отсечь обвноление сообщения
	AddField(AnswerHeader, "datetime", "created_at");
	AddField(AnswerHeader, "datetime", "processed_at");
	AddField(AnswerHeader, "varchar(8)", "error_code");
	AddField(AnswerHeader, "varchar(256)", "error_more");
	
	//	Таблица с ответом в строках сообщения
	AnswerChild = New ValueTable;
	AddField(AnswerChild, "varchar(36)", "header_id");
	AddField(AnswerChild, "varchar(36)", "id");
	AddField(AnswerChild, "varchar(8)", "error_code");
	
	AnswerHeaderRow = AnswerHeader.Add();
	AnswerHeaderRow.id = ERPData.header.id;
	AnswerHeaderRow.state = ERPData.header.state;
	//	#AMP Попов 10.12.2024 Используем поле created_at в ключах ответа, чтобы отсечь обвноление сообщения
	AnswerHeaderRow.created_at = ERPData.header.created_at;
	AnswerHeaderRow.processed_at = ERPData.header.processed_at;
	AnswerHeaderRow.error_code = ERPData.header.error_code;
	AnswerHeaderRow.error_more = ERPData.header.error_more;
	
	UpdateLines(ERPData.Connection, ERPData.HeaderTableName, AnswerHeader);

	//	#AMP Попов 10.12.2024 Проверим, а не выгрузили ли нам новое сообщение?
	CreatedAtTable = New ValueTable;
	AddField(CreatedAtTable, "datetime", "created_at");

	SelectLines(ERPData.Connection,
		"SELECT ""created_at"" FROM """ + ERPData.HeaderTableName + """ WHERE ""id"" = '" + ERPData.header.id + "';",
		CreatedAtTable);
		
	If CreatedAtTable.Count() = 0 Then
		WriteLogEvent("Интеграция 1C-EME. Импорт. " + ERPData.HeaderTableName,
			EventLogLevel.Error,,,"После загрузки удалено сообщение с id=" + ERPData.header.id);			
		Return;
	Else
		CreatedAt = CreatedAtTable.Get(0).created_at;
		If CreatedAt <> ERPData.header.created_at Then
			WriteLogEvent("Интеграция 1C-EME. Импорт. " + ERPData.HeaderTableName,
				EventLogLevel.Error,,,"Для сообщения с id=" + ERPData.header.id +
				" не совпадает created_at до и после загрузки сообщения: " +
            	ERPData.header.created_at + "->" + CreatedAt);			
			Return;
	 	EndIf;
	EndIf;
	
	For ChildIndex = 0 To ERPData.ChildTables.Count() - 1 Do
		AnswerChild.Clear();
		For ChildLine = 0 To ERPData.ChildTables[ChildIndex].Count() - 1 Do
			ChildRow = ERPData.ChildTables[ChildIndex][ChildLine];
			If ChildRow.error_code <> "" Then
				AnswerChildRow = AnswerChild.Add();
				AnswerChildRow.header_id = ChildRow.header_id;
				AnswerChildRow.id = ChildRow.id;
				AnswerChildRow.error_code = ChildRow.error_code;
			EndIf
		EndDo;
		UpdateLines(ERPData.Connection, ERPData.ChildTableNames[ChildIndex], AnswerChild);
	EndDo		
	
EndProcedure

//	AddField добавляет поле FieldName с типом FieldType в таблицу Table.
Procedure AddField(Table, Val FieldType, Val FieldName)
	
	FieldType = Lower(FieldType);
	
	If FieldType = "integer" Or FieldType = "int" Then
		
		IntegerQualifier = New NumberQualifiers(10, 0, AllowedSign.Any);
		TypeDescription = New TypeDescription("Number", , IntegerQualifier);
		Table.Columns.Add(FieldName, TypeDescription);
		
	ElsIf FieldType = "date" Then
		
		DateQualifier = New DateQualifiers(DateFractions.Date);
		TypeDescription = New TypeDescription("Date", , DateQualifier);
		Table.Columns.Add(FieldName, TypeDescription);
		
	ElsIf FieldType = "time" Then
		
		TimeQualifier = New DateQualifiers(DateFractions.Time);
		TypeDescription = New TypeDescription("Date", , TimeQualifier);
		Table.Columns.Add(FieldName, TypeDescription);
		
	ElsIf FieldType = "datetime" Then
		
		DateTimeQualifier = New DateQualifiers(DateFractions.DateTime);
		TypeDescription = New TypeDescription("Date", , DateTimeQualifier);
		Table.Columns.Add(FieldName, TypeDescription);
		
	ElsIf FieldType = "float" Then
		
		FloatQualifier = New NumberQualifiers(15, 6, AllowedSign.Any);
		TypeDescription = New TypeDescription("Number", , FloatQualifier);
		Table.Columns.Add(FieldName, TypeDescription);
		
	Else
		
		Length = StrLen(FieldType);
		If Left(FieldType, 8) = "varchar(" And Length > 9 And Right(FieldType, 1) = ")" Then
			Size = Mid(FieldType, 9);
			Size = Number(Лев(Size, StrLen(Size) - 1));
			StringQualifiers = New StringQualifiers(Size, AllowedLength.Variable);
			TypeDescription = New TypeDescription("String", , StringQualifiers);
			Table.Columns.Add(FieldName, TypeDescription);
		Else
			Raise "Неподдерживаемый тип " + FieldType + " у поля " + FieldName;
		EndIf
		
	EndIf
		
EndProcedure

//	GetSqlType возвращает SQL-тип колонки таблицы.
Function GetSqlType(Val Type)
	
	SqlType = "";

	If Type.ContainsType(Type("String")) Then
		
		SqlType = "varchar(" + Type.StringQualifiers.Length + ")";
		
	ElsIf Type.ContainsType(Type("Number")) Then
		
		If Type.NumberQualifiers.FractionDigits = 0 Then
			SqlType = "integer";
		Else
			SqlType = "float";
		EndIf
			
	ElsIf Type.ContainsType(Type("Date")) Then
		
		If Type.DateQualifiers.DateFractions = DateFractions.Time Then
			SqlType = "time";
		ElsIf  Type.DateQualifiers.DateFractions = DateFractions.DateTime Then
			SqlType = "datetime";
		ElsIf Type.DateQualifiers.DateFractions = DateFractions.Date Then
			SqlType = "date";
		Else
			Raise "Неподдерживаемый тип даты" + Type.DateQualifiers.DateFractions;
		EndIf
		
	Else
		
		Raise "Неподдерживаемый тип " + Type;
		
	EndIf;
	
	Return SqlType;

EndFunction

//	AsSqlValue возвращает значение в форме, пригодной для использования в SQL-выражениях.
Function AsSqlValue(Val Value, Val Type)
	
	If Value = Undefined Then
		SqlValue = "NULL";
	Else
		If Type.ContainsType(Type("String")) Then
			
			SqlValue = "'" + StrReplace(Value, "'", "''") + "'";
			
		ElsIf Type.ContainsType(Type("Number")) Then
			
			If Type.NumberQualifiers.FractionDigits = 0 Then
				SqlValue = Format(Value, "NZ=0;NG=0");
			Else
				SqlValue = Format(Value, "NZ=0;NG=0;NDS=.");
			EndIf
				
		ElsIf Type.ContainsType(Type("Date")) Then
			
			If Type.DateQualifiers.DateFractions = DateFractions.Time Then
				SqlValue = Format(Value, "DF=""HH:mm:ss.000""");
				SqlValue = "{ t '" + SqlValue + "' }";
			Else
				If Value < '19000101' Then
					SqlValue = "NULL";
				Else
					If Type.DateQualifiers.DateFractions = DateFractions.DateTime Then
						SqlValue = Format(Value, "DF=""yyyy-MM-dd HH:mm:ss.000""");
						SqlValue = "{ ts '" + SqlValue + "' }";
					ElsIf Type.DateQualifiers.DateFractions = DateFractions.Date Then
						SqlValue = Format(Value, "DF=""yyyy-MM-dd""");
						SqlValue = "{ d '" + SqlValue + "' }";
					Else
						Raise "Неподдерживаемый тип даты" + Type.DateQualifiers.DateFractions;
					EndIf
				EndIf
			EndIf;
			
		Else
			
			Raise "Неподдерживаемый тип " + Type;
			
		EndIf;
	EndIf;
	
	Return SqlValue;

EndFunction

//	AsString возвращает значение в виде строки.
Function AsString(Val Value, Val Type)
	
	If Value = Undefined Then
		SqlValue = "";
	Else
		If Type.ContainsType(Type("String")) Then
			
			SqlValue = Value;
			
		ElsIf Type.ContainsType(Type("Number")) Then
			
			If Type.NumberQualifiers.FractionDigits = 0 Then
				SqlValue = Format(Value, "NZ=0;NG=0");
			Else
				SqlValue = Format(Value, "NZ=0;NG=0;NDS=.");
			EndIf
				
		ElsIf Type.ContainsType(Type("Date")) Then
			
			If Type.DateQualifiers.DateFractions = DateFractions.Time Then
				SqlValue = Format(Value, "DF=""HH:mm:ss""");
			Else
				If Value < '19000101' Then
					SqlValue = "0100-01-01T00:00:00";
				Else
					If Type.DateQualifiers.DateFractions = DateFractions.DateTime Then
						SqlValue = Format(Value, "DF=""yyyy-MM-ddTHH:mm:ss""");
					ElsIf Type.DateQualifiers.DateFractions = DateFractions.Date Then
						SqlValue = Format(Value, "DF=""yyyy-MM-dd""");
					Else
						Raise "Неподдерживаемый тип даты" + Type.DateQualifiers.DateFractions;
					EndIf
				EndIf
			EndIf;
			
		Else
			
			Raise "Неподдерживаемый тип " + Type;
			
		EndIf;
	EndIf;
	
	Return SqlValue;

EndFunction

//	CastCase приводит имя <Name> к верхнему регистру в зависимости от настроек. Возвращает приведенное имя.
Function CastCase(ERPData, Val Name)
	Return Name;
EndFunction

//	ReplaceMacro производит контекстную замену макросов в строке SQL-запроса. Возвращает преобразованную строку SQL-запроса.
Function ReplaceMacro(ERPData, Val SqlQuery)
	
	Count = ERPData.ChildTableNames.Count();
	Index = ERPData.ChildIndex;
	If Index >= 0 And Index < Count Then
		SqlQuery = StrReplace(SqlQuery, "$CHILD$", ERPData.ChildTableNames[ERPData.ChildIndex]);
	EndIf;
	SqlQuery = StrReplace(SqlQuery, "$HEADER$", ERPData.HeaderTableName);
	SqlQuery = StrReplace(SqlQuery, "[", """");
	SqlQuery = StrReplace(SqlQuery, "]", """");
	SqlQuery = StrReplace(SqlQuery, "$STATE$", CastCase(ERPData, "state"));
	SqlQuery = StrReplace(SqlQuery, "$ID$", CastCase(ERPData, "id"));
	SqlQuery = StrReplace(SqlQuery, "$HEADER_ID$", CastCase(ERPData, "header_id"));
	SqlQuery = StrReplace(SqlQuery, "$CREATED_AT$", CastCase(ERPData, "created_at"));
	SqlQuery = StrReplace(SqlQuery, "$PROCESSED_AT$", CastCase(ERPData, "processed_at"));
	SqlQuery = StrReplace(SqlQuery, "$ERROR_CODE$", CastCase(ERPData, "error_code"));
	
	Return SqlQuery;
	
EndFunction

//	InsertLines добавляет строки из 1C-таблицы Table в Sql-таблицу TableName.
Procedure InsertLines(Connection, Val TableName, Table)
	
	Connection.CommandTimeout = 600;
	
	Columns = "(";
	For ColumnIndex = 0 To Table.Columns.Count() - 1 Do
		If ColumnIndex <> 0 Then
			Columns = Columns + ", ";
		EndIf;
		Columns = Columns + """" + Table.Columns[ColumnIndex].Name + """";
	EndDo;
	Columns = Columns + ")";
	
	Count = 0;
	Total = 0;
	Command = "";
	For Each TableRow In Table Do
		Values = "(";
		For ColumnIndex = 0 To Table.Columns.Count() - 1  Do
			
			If ColumnIndex <> 0 Then
				Values = Values + ", ";
			EndIf;
			
			Value = TableRow[ColumnIndex];
			Type = Table.Columns[ColumnIndex].ValueType;
			
			Values = Values + AsSqlValue(Value, Type);
			
		EndDo;
		Values = Values + ")";
		Command = Command + "INSERT INTO """ + TableName + """ " + Columns + " VALUES " + Values + ";";
		
		Count = Count + 1;
		If Count = 256 Then
			Connection.Execute(Command);
			Command = "";
			Total = Total + Count;
			Count = 0;
			WriteLogEvent("Интеграция 1C-EME. Экспорт", EventLogLevel.Information,,,"Insert " + Total + " records into " + TableName);
		EndIf

	EndDo;
	
	If Command <> "" Then
		Connection.Execute(Command);
	EndIf

EndProcedure

//	SelectLines выбирает строки запроса QueryText и помещает их в 1C-таблицу значений Table.
Procedure SelectLines(Connection, Val QueryText, Table)
	
	//	Очистим 1С-таблицу значений
	Table.Clear();
	
	//	Создадим рекордсет с параметрами adOpenForwardOnly и adLockReadOnly
	Recordset = New COMObject("ADODB.Recordset");
	Recordset.Open(QueryText, Connection, 0, 1, -1);
	
	//	Найдем соответствие полей рекордсета и колонок 1С-таблицы значений
	SQLTo1C = New Array;
	FieldsCount = Recordset.Fields.Count;
	For FieldIndex = 0 To FieldsCount - 1 Do
		SQLTo1C.Add();
		SQLTo1C[FieldIndex] = -1;
		
		FieldColumn = Table.Columns.Find(Recordset.Fields(FieldIndex).Name);
		If FieldColumn <> Undefined Then
			SQLTo1C[FieldIndex] = Table.Columns.IndexOf(FieldColumn);
		Else
			SQLTo1C[FieldIndex] = -1;
		EndIf

	EndDo;
	
	//	Скопируем данные из рекордсета в 1С-таблицу значений
	If Not Recordset.EOF Then
		Recordset.MoveFirst();
		While Not Recordset.EOF Do
			Row = Table.Add();
			For FieldIndex = 0 To FieldsCount - 1 Do
				If SQLTo1C[FieldIndex] <> -1 Then
					Row[SQLTo1C[FieldIndex]] = Recordset.Fields(FieldIndex).Value;
				EndIf
			EndDo;
			Recordset.MoveNext();
		EndDo;
	EndIf;
	
	//	Закроем рекордсет
	Recordset.Close();
	
EndProcedure

//	UpdateLines обновляет строки в Sql-таблице TableName по данным в 1С-таблице значений.
Procedure UpdateLines(Connection, Val TableName, Table)
	
	For Each TableRow In Table Do
		
		SetExpression = "";
		WhereExpression = "";
		
		For ColumnIndex = 0 To Table.Columns.Count() - 1  Do
			
			Name = Table.Columns[ColumnIndex].Name;
			Type = Table.Columns[ColumnIndex].ValueType;
			Value = TableRow[ColumnIndex];
			
			//	#AMP Попов 10.12.2024 Используем поле created_at в ключах ответа, чтобы отсечь обновление сообщения
			//If Name = "id" Or Name = "header_id" Then
			If Name = "id" Or Name = "header_id" Or Name = "created_at" Then
				If WhereExpression <> "" Then
					WhereExpression = WhereExpression + " AND ";
				EndIf;
				WhereExpression = WhereExpression + """" + Name + """=" + AsSqlValue(Value, Type);
			Else
				If SetExpression <> "" Then
					SetExpression = SetExpression + ", ";
				EndIf;
				SetExpression = SetExpression + """" + Name + """=" + AsSqlValue(Value, Type);
			EndIf
				
		EndDo;
		
		Connection.Execute("UPDATE """ + TableName + """ SET " + SetExpression + " WHERE " + WhereExpression + ";");
		
	EndDo;
	
EndProcedure

//	Convert преобразует текст OldCodeLines, использующий движок ERPEngine.dll,
//	в текст NewCodeLines, использующий общий модуль EmeWmsERPEgine;
Procedure Convert(OldCodeLines, NewCodeLines) Export
	
	//	Имя объекта движка
	ERPEngine = "ERPEngine";
	
	//	Свойства
	Properties = New Array;
	Properties.Add("Config");
	Properties.Add("EnableLog");
	Properties.Add("LogFile");
	
	//	Методы движка ERPEngine.dll
	Methods = New Array;
	Methods.Add("Connect");
	Methods.Add("Disconnect");
	Methods.Add("BeginExport");
	Methods.Add("CommitExport");
	Methods.Add("RollbackExport");
	Methods.Add("AppendHeaderLine");
	Methods.Add("PutHeaderData");
	Methods.Add("PutHeaderDataAsText");
	Methods.Add("SelectChild");
	Methods.Add("AppendChildLine");
	Methods.Add("PutChildData");
	Methods.Add("PutChildDataAsText");
	Methods.Add("BeginImport");
	Methods.Add("CommitImport");
	Methods.Add("RollbackImport");
	Methods.Add("NextHeaderLine");
	Methods.Add("GetHeaderData");
	Methods.Add("GetHeaderDataAsText");
	Methods.Add("NextChildLine");
	Methods.Add("GetChildData");
	Methods.Add("GetChildDataAsText");
	Methods.Add("ErrorHeader");
	Methods.Add("ErrorChild");
	Methods.Add("HasErrors");
	Methods.Add("WarningHeader");
	Methods.Add("WarningChild");
	Methods.Add("HasWarnings");
	Methods.Add("Success");
	Methods.Add("Log");

	Child = "";
	
	For Index = 0 To OldCodeLines.Count() - 1 Do
		OldLine = OldCodeLines.Get(Index);
		NewLine = "";
		
		//	COM-объект
		If Find(OldLine, "EME.ERPEngine") > 0 Then
			NewCodeLines.Add(Символ(9) + "ERPData = Новый Структура;");
			NewCodeLines.Add(Символ(9) + "EmeWmsERPEngine.Create(ERPData);");
			Continue;
		EndIf;
		
		//	Свойства
		Property = CheckKeyWord(OldLine, Properties);
		If Property <> "" Then
			If Find(OldLine, ERPEngine + "." + Property) Then
				NewLine = StrReplace(OldLine, ERPEngine + "." + Property, "ERPData." + Property);
			Endif
		EndIf;
		
		//	Методы
		If NewLine = "" Then
			
			Method = CheckKeyWord(OldLine, Methods);
			If Method <> "" Then
				CallMethod = ERPEngine + "." + Method + "(";
				Pos = Find(OldLine, CallMethod);
				If Find(OldLine, CallMethod) > 0 Then
					NewLine = NewLine + Left(OldLine, Pos);
					If Find(Method, "Data") > 0 Then
						
						If Find(OldLine, CallMethod) > 0 Then
							//	Таблица данных
							Table = Lower(Child);
							If Find(Method, "Header") > 0 Then
								Table = "header";
							EndIf;
								
							//	Поле данных
							Field = ExtractString(OldLine);
							If Find(OldLine, """" + Field + """,") > 0 Then
								OldLine = StrReplace(OldLine, """" + Field + """,", "");
							Else
								OldLine = StrReplace(OldLine, """" + Field + """", "");
							EndIf;
							
							//	Оператор присваивания
							Assignment = "";
							If Find(Method, "Put") > 0 Then
								Assignment = " =";
							EndIf;
							
							NewLine = StrReplace(OldLine, CallMethod, "ERPData." + Table + "." + Lower(Field) + Assignment);
							
							//	Удалим закрывающую скобку
							NewLine = StrReplace(NewLine, ");", ";");
						EndIf
						
					Else
						If Find(OldLine, CallMethod + ")") > 0 Then
							NewLine = StrReplace(OldLine, CallMethod, "EmeWmsERPEngine." + Method + "(ERPData");
						Else
							NewLine = StrReplace(OldLine, CallMethod, "EmeWmsERPEngine." + Method + "(ERPData, ");
						EndIf;
						If Method = "SelectChild" Then
							Child = ExtractString(OldLine);
						EndIf
					EndIf
				EndIf
			EndIf
		EndIf;

		If NewLine = "" Then
			NewLine = OldLine;
		EndIf;
		
		NewLine = StrReplace(NewLine, "(" + ERPEngine, "(ERPData");
		
		NewCodeLines.Add(NewLine);
	EndDo
EndProcedure

//	Convert преобразует текст OldCodeLines, использующий движок ERPEngine.dll,
//	в текст NewCodeLines, использующий общий модуль EmeWmsERPEgine;
Procedure ConvertNew(OldCodeLines, NewCodeLines) Export
	
	//	Имя объекта движка
	ERPEngine = "ERPEngine";
	
	//	Свойства
	Properties = New Array;
	Properties.Add("Config");
	Properties.Add("EnableLog");
	Properties.Add("LogFile");
	
	//	Методы движка ERPEngine.dll
	Methods = New Array;
	Methods.Add("Connect");
	Methods.Add("Disconnect");
	Methods.Add("BeginExport");
	Methods.Add("CommitExport");
	Methods.Add("RollbackExport");
	Methods.Add("AppendHeaderLine");
	Methods.Add("PutHeaderData");
	Methods.Add("PutHeaderDataAsText");
	Methods.Add("SelectChild");
	Methods.Add("AppendChildLine");
	Methods.Add("PutChildData");
	Methods.Add("PutChildDataAsText");
	Methods.Add("BeginImport");
	Methods.Add("CommitImport");
	Methods.Add("RollbackImport");
	Methods.Add("NextHeaderLine");
	Methods.Add("GetHeaderData");
	Methods.Add("GetHeaderDataAsText");
	Methods.Add("NextChildLine");
	Methods.Add("GetChildData");
	Methods.Add("GetChildDataAsText");
	Methods.Add("ErrorHeader");
	Methods.Add("ErrorChild");
	Methods.Add("HasErrors");
	Methods.Add("WarningHeader");
	Methods.Add("WarningChild");
	Methods.Add("HasWarnings");
	Methods.Add("Success");
	Methods.Add("Log");

	Child = "";
	
	For Index = 0 To OldCodeLines.Count() - 1 Do
		OldLine = OldCodeLines.Get(Index);
		NewLine = "";
		
		//	COM-объект
		If Find(OldLine, "EME.ERPEngine") > 0 Then
			NewCodeLines.Add(Символ(9) + "ERPData = Новый Структура;");
			NewCodeLines.Add(Символ(9) + "EmeWmsERPEngine.Create(ERPData);");
			Continue;
		EndIf;
		
		//	Свойства
		Property = CheckKeyWord(OldLine, Properties);
		If Property <> "" Then
			If Find(OldLine, ERPEngine + "." + Property) Then
				NewLine = StrReplace(OldLine, ERPEngine + "." + Property, "ERPData." + Property);
			Endif
		EndIf;
		
		//	Методы
		If NewLine = "" Then
			While OldLine <> "" Do
				
				For Each Method In Methods Do
					CallMethod = ERPEngine + "." + Method + "(";
					Pos = Find(OldLine, CallMethod);
					If (Pos > 0) Then
						Break
					EndIf
				EndDo;
				
				If Pos > 0 Then
					NewLine = NewLine + Left(OldLine, Pos - 1);
					OldLine = Mid(OldLine, Pos + StrLen(CallMethod));
					If Find(Method, "Data") > 0 Then
						NewLine = NewLine + "ERPData.";
						
						//	Таблица данных
						If Find(Method, "Header") > 0 Then
							NewLine = NewLine + "header";
						Else
							NewLine = NewLine + Child;
						EndIf;
						
						NewLine = NewLine + ".";
						
						//	Поле данных
						FieldName = ExtractString(OldLine);
						NewLine = NewLine + FieldName;
						If Find(OldLine, """" + FieldName + """,") > 0 Then
							OldLine = StrReplace(OldLine, """" + FieldName + """,", "");
						Else
							OldLine = StrReplace(OldLine, """" + FieldName + """", "");
						EndIf;
							
						//	Оператор присваивания
						If Find(Method, "Put") > 0 Then
							NewLine = NewLine + " =";
						EndIf;
						
						//	Найдем и удалим закрывающую скобку
						BraceCount = 1;
						While OldLine <> "" Do
							Symbol = Left(OldLine, 1);
							OldLine = Mid(OldLine, 2);
							If Symbol = "(" Then
								BraceCount = BraceCount + 1;
							ElsIf Symbol = ")" Then
								BraceCount = BraceCount - 1;
							EndIf;
							If BraceCount = 0 Then
								Break
							Else
								NewLine = NewLine + Symbol;
							EndIf
						EndDo
						
					Else
						NewLine = NewLine + "EmeWmsERPEngine." + Method + "(ERPData";
						
						If Left(OldLine, 1) <> ")" Then
							NewLine = NewLine + ", ";
						EndIf;
						
						If Method = "SelectChild" Then
							Child = ExtractString(OldLine);
						EndIf
					EndIf
				Else
					NewLine = NewLine + OldLine;
					OldLine = "";
				EndIf;
			EndDo
		EndIf;

		If NewLine = "" Then
			NewLine = OldLine;
		EndIf;
		
		NewLine = StrReplace(NewLine, "(" + ERPEngine, "(ERPData");
		
		NewCodeLines.Add(NewLine);
	EndDo
EndProcedure

//	CheckKeyWord проверяет наличие в строке Line ключевого слова из массива ключевых слов KeyWords.
//	Если ключевое слово найдено, то возвращает его, иначе возвращает пустую строку "".
Function CheckKeyWord(Val Line, KeyWords)
		
	For Each KeyWord In KeyWords Do
		If Find(Line, KeyWord) > 0 Then
			Return KeyWord;
		EndIf;
	EndDo;		
		
	Return "";
	
EndFunction

//	 ExtractString извлекает из строки <Line> подстроку ограниченную кавычками.
Function ExtractString(Val Line)
	FirstPos = Find(Line, """");
	If FirstPos > 0 Then
		LastPos = Find(Mid(Line, FirstPos + 1), """");
		If LastPos > 0 Then
			Return Mid(Line, FirstPos + 1, LastPos - 1);
		EndIf
	EndIf;
	Return "???";
EndFunction
