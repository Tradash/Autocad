Public Const AttrName = "Номер"
Public Const Connect_Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Dsk3main\PROD\Obmen\Словарь_изделий_ДСК3.mdb"
Public Sub start_script()
UserForm1.Show
End Sub

Public Sub set_attr(Name_Attr As String, oBlockRef As IAcadBlockReference2, value As Variant)
'' Устанавливает значение заданного атрибута в определенном элементе
Dim i As Integer
Dim varAttributes As Variant
varAttributes = oBlockRef.GetAttributes
For i = LBound(varAttributes) To UBound(varAttributes)
If varAttributes(i).TagString = Name_Attr Then
varAttributes(i).TextString = value
End If
Next
End Sub

Public Function get_attr(Name_Attr As String, oBlockRef As IAcadBlockReference2) As Variant
'' Возвращает значание заданного атрибута в определенном элементе
Dim i As Integer
Dim varAttributes As Variant
varAttributes = oBlockRef.GetAttributes
For i = LBound(varAttributes) To UBound(varAttributes)
If varAttributes(i).TagString = Name_Attr Then
get_attr = varAttributes(i).TextString
End If
Next
End Function

Public Static Sub get_array(kod() As String, counter() As Integer)
'' Возвращает в массиве все handler на Blockref и их номер монтажа
Dim i As Integer, i_max As Integer
Dim ObjEnt As Variant
Dim objent1 As AcadBlockReference

i_max = Get_Count_BlockRef()

ReDim kod(i_max)
ReDim counter(i_max)
i = 0
For Each ObjEnt In ThisDrawing.ModelSpace
If ObjEnt.ObjectName = "AcDbBlockReference" Then
i = i + 1
kod(i) = ObjEnt.Handle
Set objent1 = ObjEnt
counter(i) = get_attr(AttrName, objent1)
End If
Next
End Sub

Public Static Sub sort_array(kod() As String, counter() As Integer)
Dim i As Integer, j As Integer
Dim tmp_int As Integer, tmp_str As String

For i = UBound(kod) - 1 To 1 Step -1
For j = 1 To i
If counter(j) > counter(j + 1) Then
tmp_int = counter(j)
tmp_str = kod(j)
counter(j) = counter(j + 1)
kod(j) = kod(j + 1)
counter(j + 1) = tmp_int
kod(j + 1) = tmp_str
End If
Next
Next

End Sub

Public Function Get_Count_BlockRef() As Integer
Dim i As Integer
Dim ObjEnt As Variant
Get_Count_BlockRef = 0
For Each ObjEnt In ThisDrawing.ModelSpace
If ObjEnt.ObjectName = "AcDbBlockReference" Then
Get_Count_BlockRef = Get_Count_BlockRef + 1
End If
Next
End Function
Public Static Sub create_tbl(tbl() As String, Attr_Name() As String)
Dim Count_Attr As Integer
Dim ObjEnt As AcadBlockReference
Dim Obj_A As Variant
Dim varAttributes As Variant
Dim i As Integer, j As Integer, k As Integer
Dim z As Variant

Count_Attr = 0
''Ищем количество аттрибутов
For Each Obj_A In ThisDrawing.ModelSpace
If Obj_A.ObjectName = "AcDbBlockReference" Then
Set ObjEnt = Obj_A
varAttributes = ObjEnt.GetAttributes
If Count_Attr < UBound(varAttributes) + 1 Then Count_Attr = UBound(varAttributes) + 1
End If
Next
Count_Attr = Count_Attr - 1
ReDim Attr_Name(Count_Attr)
j = 0
For Each Obj_A In ThisDrawing.ModelSpace
If Obj_A.ObjectName = "AcDbBlockReference" Then
Set ObjEnt = Obj_A
varAttributes = ObjEnt.GetAttributes
For i = LBound(varAttributes) To UBound(varAttributes)
k = 1
For z = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(z) = varAttributes(i).TagString Then k = 0
Next
If k = 1 Then
Attr_Name(j) = varAttributes(i).TagString
j = j + 1
End If
Next
End If
Next

ReDim tbl(Get_Count_BlockRef() - 1, Count_Attr + 1)

k = -1
j = 0
For Each Obj_A In ThisDrawing.ModelSpace
If Obj_A.ObjectName = "AcDbBlockReference" Then
Set ObjEnt = Obj_A
varAttributes = ObjEnt.GetAttributes
k = k + 1
tbl(k, 0) = ObjEnt.name
For i = LBound(varAttributes) To UBound(varAttributes) '
For j = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(j) = varAttributes(i).TagString Then tbl(k, j + 1) = varAttributes(i).TextString
Next
Next
End If
Next

End Sub
Public Static Sub Get_Obj(name As String, kod As String)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim tbl1() As String
Dim i As Integer, j As Integer
Dim tmp_str As String

cn.ConnectionString = Connect_Str
cn.Open
tmp_str = "SELECT count(FOBRN.kodob) FROM FOBRN where sostob=2"
rs.Open tmp_str, cn
ReDim tbl1(rs.Fields(0).value, 2)
rs.Close
tmp_str = "SELECT FOBRN.ADRES, FOBRN.KODOB FROM FOBRN where sostob=2 order by FOBRN.KODOB "
rs.Open tmp_str, cn
i = 0
While Not rs.EOF
i = i + 1
tbl1(i, 0) = rs.Fields(0).value
tbl1(i, 1) = rs.Fields(1).value
rs.MoveNext
Wend
rs.Close
cn.Close
UserForm3.ListBox1.List() = tbl1
UserForm3.Show
kod = UserForm3.ListBox1.List(UserForm3.ListBox1.ListIndex, 1)
name = UserForm3.ListBox1.List(UserForm3.ListBox1.ListIndex, 0)
End Sub

Private Sub CommandButton9_Click()
Dim i As Integer, j As Integer, ii As Integer
Dim Count_Attr As Integer
Dim Count_BlckRf As Integer
Dim ObjEnt As AcadBlockReference
Dim varAttributes As Variant
Dim Attr_Name() As String
Dim z As Variant
Dim str_s As String
Dim tbl() As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim tbl1() As String
Dim tmp_str As String, msg_str As String, tmp_str1 As String
Dim Kod_Obj As String, name_obj As String
Dim i_Et As Integer, i_sostob As Integer
Dim i_Tip As Integer, i_kod As Integer
Dim Reys() As Integer 'массив рейсов

Call Get_Obj(name_obj, Kod_Obj)
Unload UserForm3
Count_BlckRf = Get_Count_BlockRef()
Call create_tbl(tbl, Attr_Name)
For i = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(i) = "Этаж" Then i_Et = i + 1 'Определяем в каком поле Этаж
Next

cn.ConnectionString = Connect_Str
cn.Open
tmp_str = "delete FROM fkmpobn where kodob=" & Kod_Obj & " and etag1=" & tbl(0, i_Et)
rs.Open tmp_str, cn 'Удаляем информацию по объекту и этажу
ii = tbl(0, i_Et)
For i = 1 To Count_BlckRf - 1
If ii <> tbl(i, i_Et) Then
tmp_str = "delete FROM fkmpobn where kodob=" & Kod_Obj & " and etag1=" & tbl(i, i_Et)
rs.Open tmp_str, cn 'Удаляем информацию по объекту и этажу
ii = tbl(i, i_Et)
End If

Next
'rs.Close

For i = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(i) = "Рейс" Then ii = i + 1 'Определяем в каком поле номер рейса
Next

For i = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(i) = "Транспорт" Then i_Tip = i + 1 'Определяем в каком поле номер рейса
Next

For i = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(i) = "Код" Then i_kod = i + 1 'Определяем в каком поле код изделия
Next


n_reys = 0
'For i = 0 To Count_BlckRf - 1
' If n_reys < tbl(i, k) Then n_reys = tbl(i, ii) 'Определяем количество рейсов
'Next
ReDim Reys(Count_BlckRf)
n_reys = 0
For i = 0 To Count_BlckRf - 1
k = 1
For j = 0 To UBound(Reys)
If Reys(j) = tbl(i, ii) Then k = 0
Next
If k = 1 Then
Reys(n_reys) = tbl(i, ii)
n_reys = n_reys + 1
End If
Next

For i = 0 To n_reys - 1 'По каждому рейсу формируем записи
k = 0
For j = 0 To Count_BlckRf - 1
If tbl(j, ii) = Reys(i) Then
k = k + 1
If k = 1 Then 'если это первый раз
If tbl(j, i_Et) = 0 Then i_sostob = 3 Else i_sostob = 2
tmp_str = "insert into fkmpobn (kodob, sostob, etag1, etag2, kodr, tipmash, zavod"
tmp_str1 = " values (" & Kod_Obj & ", " & i_sostob & ", " & tbl(j, i_Et) & ", " & tbl(j, i_Et) & ", " _
& Reys(i) & ", "
Select Case UCase(tbl(j, i_Tip))
Case "ПЛ"
tmp_str1 = tmp_str1 & "2, "
Case "Ш"
tmp_str1 = tmp_str1 & "3, "
Case "ЭР"
tmp_str1 = tmp_str1 & "4, "
Case Else
tmp_str1 = tmp_str1 & "1, "
End Select
tmp_str1 = tmp_str1 & """В" & """"
End If
tmp_str = tmp_str & ", marka" & k & ", kol" & k
tmp_str1 = tmp_str1 & ", " & tbl(j, i_kod) & "00, 1"
End If

Next
If Reys(i) = 39 Then
z = 0
End If
tmp_str = tmp_str & ") " & tmp_str1 & ") "
rs.Open tmp_str, cn ' добавляем строку

'rs.Close

Next

cn.Close
End Sub

Static Sub check_izd(WrkSht1 As Excel.Worksheet, j1 As Integer)
Dim blr_handle() As String
Dim blr_nomer() As Integer
Dim oBR As AcadBlockReference
Dim cn As New ADODB.Connection
Dim tmp_str As String, msg_str As String, tmp_str1 As String
'Dim WrkBk As Excel.Workbook
'Dim WrkSht As Excel.Worksheet
Dim i As Integer, z As Integer

'UserForm1.Hide

cn.ConnectionString = Connect_Str
cn.Open

Dim rs As New ADODB.Recordset

''rs.Open "SELECT count([slov2].[SHSL]) AS [count-record] FROM slov2", cn

''MsgBox "Количество строк в базе" & rs.Fields(0).value
z = 0
msg_str = "Все в порядке"
Call get_array(blr_handle, blr_nomer)
For i = LBound(blr_handle) + 1 To UBound(blr_handle)
Set oBR = ThisDrawing.HandleToObject(blr_handle(i))
tmp_str = get_attr("Марка", oBR)
tmp_str1 = "SELECT count(slov2.RSHSL) FROM slov2 WHERE (((slov2.RSHSL)=""" & tmp_str & """))"
rs.Open tmp_str1, cn
If rs.Fields(0).value = 0 Then
If z = 0 Then
'On Error Resume Next
'Set app = GetObject(, "Excel.Application")
'If Err Then
' Err.Clear
' Set app = CreateObject("Excel.Application")
' If Err Then
' Exit Sub
'End If
'End If
'app.Visible = False
'Set WrkBk = app.Workbooks.Add
'Set WrkSht = WrkBk.Worksheets.Add
'WrkSht.name = "Проблемы"
'j = 3
WrkSht1.Cells(j1, 1).value = ThisDrawing.Application.ActiveDocument.Path & "\" & ThisDrawing.Application.ActiveDocument.name
WrkSht1.Cells(j1 + 1, 1).value = "Номер монтажа"
WrkSht1.Cells(j1 + 1, 2).value = "Изделие"
j1 = j1 + 3
msg_str = "Найдены проблемы" & vbCr
End If
z = z + 1
With WrkSht1
.Cells(j1, 1).value = blr_nomer(i)
.Cells(j1, 2).value = tmp_str
j1 = j1 + 1
End With
'' msg_str = msg_str & "#" & blr_nomer(i) & ":" & tmp_str & vbCr
End If
rs.Close

Next
'MsgBox msg_str
'app.Visible = True
''Unload Me
cn.Close
End Sub