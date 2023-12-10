'класс для работы с таплицами опрделенного формата
'где сначала идут строки с дополнительной информацией к заголовкам
'далее идут сами заголовки с разделением на первую и вторую половины 
'(разделитель - заголовок "break")
'(один из заголовков первой половины обязательно именуется - "№")
'после идут данные по номерам 
'(номера не должны повторяться. Повторение ведет к потере данных)
'
'
'
'
'TODO чтобы сами данные могли быть не просто значением, а форматированной записью с несколькими значениями для псевдо-многомерности
'
'
'функционал:
'-ввести имя листа
'- получить данные записи по номеру (отдельный класс - данные, поля, метаданные)
' -получить данные накладной (отдельный класс) - только если упоминается предметная область - сделать шаблон товарной накладной
'- в записи с определенным номером изменить значение определенного заголовка первой / второй половины (учесть возможность форматирования)
'- получить значение записи с определенным номером по определенному заголовку первой / второй половины (учесть возможность форматирования)
'* данные второй половины - только не пустые
'
'
'Usage:
'Dim ft__ As New FormatTable
'ft__.collect_format_mapptings "some_list"
'Dim zap_num_data As Dictionary
'Set zap_num_data = ft__.get_num_data(10)
'
'
' NOTE необходимо в зависимостях указать Microsoft Scripting Runtime (класс Dictionary)
'
'


Private __names As Dictionary
Private __format_mapping As Dictionary
Private top_row_splitter As String
Private half_splitter As String



Private Sub class_init ()
	top_row_splitter = "заголовки" ' альтернатива "данныеСтроки"
	half_splitter = "break"
End Sub


' проходит по шапке, разделяет половины 
Private Sub GetNames(top_row As Integer)
	'собирает заголовки указанные в шапке таблицы.
	' входной параметр обозначает что таблица смещена вниз
	' на выхде словарь с парами: наименование - номер столбца
	is_data_st = False
	is_fields_st = True
	Dim data As Dictionary, fields As Dictionary, meta As Dictionary
	Set data = New Dictionary
	Set fields = New Dictionary
	Set meta = New Dictionary

	el_cntr=0
	Dim last_row As Integer
	For Each row_cell In Range(CStr(top_row) + ":" + CStr(top_row))
		If row_cell.Value = Empty And is_data_st Then
			Exit For
		End If
		If Cells(top_row, row_cell.Column).Value = half_splitter Then
			is_fields_st = False
		End If

		If is_fields_st Then
			fields(row_cell.Value) = row_cell.Column
		End If

		If is_data_st Then
			last_row = row_cell.Column
			data(row_cell.Column) = row_cell.Value
		End If

		If Celis(top_row, row_cell.Column).Value = half_splitter Then
			is_data_st = True
			meta("BreakCol") = row_cell.Column
		End If
	Next row_cell
	meta("LastCol") = last_row

	Dim res As Dictionary
	Set res = New Dictionary
	Set res("méta") = meta
	Set res("fields") = fields
	Set res("data") = data

	Set __names = res
End Sub

' обработка самого офрмата
' собирает таблицы соответствия по столбцам (с разделением на половины) 
' и по строкам (с разделением на дополнительную информацию, заголовки и сами данные)
' проихводит перрегрупировку
Public Sub collect_format_mapptings(list_name As String)
	Worksheets(list_name).Activate

	Dim res As Dictionary, names As Dictionary, _
		numbers As Dictionary, _
		metadata As Dictionary, _
		group_metadata As Dictionary

	Set res = New Dictionary
	Set names = New Dictionary
	names("indication") = 0
	Set numbers = New Dictionary
	Set metadata = New Dictionary
	Set group_metadata = New Dictionary
	Dim data_st As Integer

	For Each itm In Range("A:A")
		If data_st Then
			If Cells(itm.Row, 2).Value = Empty Then
				Exit For
			End If
		End If
		If itm.Value = Empty And data_st = Empty Then
			If itm.Row <2 Then
				MsgBox "Лист (" + ActiveSheet.Name + ") неправильного формата"
			End If
			Exit For
		ElseIf itm.Value = top_row_splitter Then
			Call GetNames(top_row:=itm.Row)
			Set names = __names
			data_st = itm.Row + 1
		ElseIf itm.Row >= data_st Then
			if Not (names("fields").Exists("№")) Then
				Exit For
			End If
			Num_col = names("fields")("№")
			num = Cells(itm.Row, Num_col).Text
			numbers(num) = itm.Row
			names("indication") = 1
		ElseIf Not( itm.Value = Empty) Then
			metadata(itm.Value) = itm.Row
		End If
	Next itm

	If names("indication") = 0 Then
		MsgBox "Допущена ошибка в обработке файла"
		Exit Sub
	End If

	For Each key_ In metadata
		'собирает словари по каждой строке до начала таблицы (особые данные к данным таблицы)
		' как словарь 
		'{
		'	название_данных : {наименование_данных_из_шапки : значение} 
		'}
		Dim cargo As Dictionary
		Set cargo = get_num_data_by_num( _
			row_num:=metadata(key_), _
			names:=names)
		Set metadata(key_) = cargo("data")
	Next key_

	For Each col_num_ In names("data")
		name_=names("data")(col_num_)
		Set group_metadata(name_) = New Dictionary
		For Each param_name In metadata
			' перегруппирует особыне данные к данным таблицы из целого словаря 
			'{
			'	название_данных : {наименование_данных_из_шапки : значение} 
			'}
			' в 
			'{
			'	наименование_данных_из_шапки : {название_данных : значение} 
			'}
			group_metadata(name_)(param_name) = metadata(param_name)(name_)
		Next param_name
	Next col_num_

	Set res("names") = names
	Set res("numbers") = numbers
	Set res("grmeta") = group_metadata

	Set __format_mapping = res
End Sub


Private Function fields_autoreplacement(param_ As String, val_ )
	fields_autoreplacement = val_
End Function 


Private Function data_autoreplacement(param_ As String, val_ )
	data_autoreplacement = val_
End Function 

' создает отдельную мини таблицу только с дынными строки целевого номера
' дополнительны правила по автозаполнению отсутствующих данных
Public Function get_num_data(num As Integer) As Dictionary
	' срабатывают правила по автозаполнению 
	' не пропущенных (пустых считанных полей) или зашифрованных данных 
	' они вынесены в отдельные разделы для удобства fields_autoreplacement, data_autoreplacement
	'
	'сбор данных по конкретному номеру накладной. (соответствующей строки)
	'возврат словаря следущего формата
	'{
	'	"fields":{"поле для оформления документа":str}
	'	"data":{"наименование материальной ценности":int}
	'}
	
	If not (preprocessed_data("numbers").Exists(nacl_num) )Then
		MsgBox "Записи №" + CStr(nacl_num) + " - не существует"
		Exit Sub
	End If
	
	Set preprocessed_data = __format_mapping
	
	Dim names_ As Dictionary, row_num_ As Integer
	row_num_ = preprocessed_data("numbers")(nacl_num)
	Set names_ = preprocessed_data("names")
	
	Dim res As New Dictionary

	Dim fields As Dictionary, data As Dictionary
	Set fields = New Dictionary
	Set data = New Dictionary

	For Each field In names("fields").Keys()
		fields(field) = fields_autoreplacement( _
			param_ = field, _
			val_ = Cells(row_num, names("fields")(field)).Value)
	Next field

	
	is_st = False
	last_col = names("meta")("LastCol")
	For Each row_cell In Range(CStr(row_num) + ":" + CStr(row_num))
		If row_cell.Column > last_col Then
			Exit For
		End If
		If is_st Then
			If Not (row_cell.Value = Empty) Then
				cur_name = names("data")(row_cell.Column)
				data(cur_name) = data_autoreplacement( _
					param_ = cur_name, _
					val_ = row_cell.Value)
			End If'
		End If
		If row_cell.Column = names("meta")("BreakCol") Then
			s_st = True
		End If
	Next row_cell

	Set res("fields") = fields
	Set res("data") = data

	Set get_num_data = res
End Sub


