VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
Caption = "Меню групповой выгрузки"
ClientHeight = 2310
ClientLeft = 45
ClientTop = 375
ClientWidth = 4680
OleObjectBlob = "UserForm4.frx":0000
StartUpPosition = 1 'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim blkDef As AcadBlock, xref As AcadExternalReference
Dim fType(1) As Integer, fData(1)
Dim ss As AcadSelectionSet
Dim name As String
Dim DOC As New AcadDocument
Dim i As Integer, j As Integer
Dim WrkBk As Excel.Workbook
Dim WrkSht As Excel.Worksheet
UserForm4.Hide

For i = ThisDrawing.SelectionSets.Count To 1 Step -1
ThisDrawing.SelectionSets.Item(i - 1).Delete
Next

Set ss = ThisDrawing.SelectionSets.Add("ss5")

fType(0) = 0: fData(0) = "INSERT"
fType(1) = 2

For Each blkDef In ThisDrawing.Blocks
If blkDef.IsXRef Then
fData(1) = blkDef.name
ss.Select acSelectionSetAll, , , fType, fData

End If
Next
name = "Найдено :" & vbCr
'Open Excel

On Error Resume Next
Set app = GetObject(, "Excel.Application")
If Err Then
Err.Clear
Set app = CreateObject("Excel.Application")
If Err Then
Exit Sub
End If
End If
On Error GoTo 0
app.Visible = False
Set WrkBk = app.Workbooks.Add
Set WrkSht = WrkBk.Worksheets.Add
WrkSht.name = "Проблемы"
j = 3

For Each xref In ss
name = name & xref.Path & vbCr
Set DOC = Application.Documents.Open(xref.Path, True)
Call check_izd(WrkSht, j)
'DOC.Open xref.Path
'MsgBox xref.Path
DOC.Close (False)
Next
app.Visible = True
ss.Delete
'MsgBox name

Unload Me
End Sub

Private Sub CommandButton2_Click()
Dim app As Excel.Application
Dim Visible As Boolean
Dim WrkBk As Excel.Workbook
Dim WrkSht As Excel.Worksheet
Dim i As Integer, j As Integer
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
Dim i_nomer As Integer
Dim blkDef As AcadBlock, xref As AcadExternalReference
Dim fType(1) As Integer, fData(1)
Dim ss As AcadSelectionSet
Dim name As String
Dim DOC As New AcadDocument
Const sdvig = 3
UserForm4.Hide
For i = ThisDrawing.SelectionSets.Count To 1 Step -1
ThisDrawing.SelectionSets.Item(i - 1).Delete
Next
On Error Resume Next
Set app = GetObject(, "Excel.Application")
If Err Then
Err.Clear
Set app = CreateObject("Excel.Application")
If Err Then
Exit Sub
End If
End If
On Error GoTo 0
i_nomer = -1
app.Visible = False
Set WrkBk = app.Workbooks.Add
Set WrkSht = WrkBk.Worksheets.Add
WrkSht.name = "Выгрузка"

Set ss = ThisDrawing.SelectionSets.Add("ss8")
fType(0) = 0: fData(0) = "INSERT"
fType(1) = 2
For Each blkDef In ThisDrawing.Blocks
If blkDef.IsXRef Then
fData(1) = blkDef.name
ss.Select acSelectionSetAll, , , fType, fData
End If
Next
Call Get_Obj(name_obj, Kod_Obj)
Unload UserForm3
ReDim Attr_Name(1)
'name = "Найдено :" & vbCr
For Each xref In ss
'name = name & xref.Path & vbCr
Set DOC = Application.Documents.Open(xref.Path)
Count_BlckRf = Get_Count_BlockRef()
Call create_tbl(tbl, Attr_Name)
i_nomer = i_nomer + 1
If i_nomer = 0 Then
With WrkSht
.Cells(i_nomer + 1, 5 + sdvig).value = "Выгрузка из Autocad по объекту:"
.Cells(i_nomer + 2, 5 + sdvig).value = "Время выгрузки"
.Cells(i_nomer + 3, 5 + sdvig).value = Now
.Cells(i_nomer + 3, 5 + sdvig).Font.Bold = True
'.Cells(i_nomer + 3, 5 + sdvig).value = "**********************"
.Cells(i_nomer + 4, 1 + sdvig).value = "##"
.Cells(i_nomer + 4, 2 + sdvig).value = "Код объекта"
.Cells(i_nomer + 4, 3 + sdvig).value = "Наименование объекта"
.Cells(i_nomer + 4, 4 + sdvig).value = "Наименование Блока ACAD"
End With
For i = LBound(Attr_Name) To UBound(Attr_Name)
WrkSht.Cells(i_nomer + 4, i + 5 + sdvig).value = Attr_Name(i)
Next
End If
For i = 0 To Count_BlckRf - 1
WrkSht.Cells(i_nomer + i + 5, 1 + sdvig).value = i + 1
WrkSht.Cells(i_nomer + i + 5, 2 + sdvig).value = Kod_Obj
WrkSht.Cells(i_nomer + i + 5, 3 + sdvig).value = name_obj
For j = 0 To UBound(Attr_Name) + 1
WrkSht.Cells(i_nomer + i + 5, j + 4 + sdvig).value = tbl(i, j)
If j > 1 Then If Attr_Name(j - 1) = "Вес" Or Attr_Name(j - 1) = "Объем" Then WrkSht.Cells(i_nomer + i + 5, j + 4 + sdvig).FormulaR1C1Local = tbl(i, j)
Next
Next
i_nomer = i_nomer + i
DOC.Close
Next
app.Visible = True
Dim oRange As Range
Dim oRange1 As Range, oRange2 As Range
Set oRange = Range(WrkSht.Cells(5, 1 + sdvig), WrkSht.Cells(i_nomer + i + 4, j + 3 + sdvig))
oRange.Font.color = vbRed
'oRange.AutoFit
ii = 0
Dim ii1 As Integer, ii2 As Integer
For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Рейс" Then ii1 = k + sdvig + 5
Next
WrkSht.Cells(1, 1).value = "Рейс"
WrkSht.Cells(1, 2).value = ii1

For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Вес" Then ii2 = k + sdvig + 5
Next
WrkSht.Cells(3, 1).value = "Вес"
WrkSht.Cells(3, 2).value = ii2

For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Марка" Then ii2 = k + sdvig + 5
Next
WrkSht.Cells(1, 3).value = "Марка"
WrkSht.Cells(1, 4).value = ii2

For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Время_монтажа" Then ii2 = k + sdvig + 5
Next
WrkSht.Cells(2, 3).value = "Время_монтажа"
WrkSht.Cells(2, 4).value = ii2

For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Номер" Then ii2 = k + sdvig + 5
Next
WrkSht.Cells(3, 3).value = "Номер"
WrkSht.Cells(3, 4).value = ii2


For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Этаж" Then ii2 = k + sdvig + 5
Next
WrkSht.Cells(2, 1).value = "Этаж"
WrkSht.Cells(2, 2).value = ii2

oRange.Sort key1:=WrkSht.Cells(5, ii2), key2:=WrkSht.Cells(5, ii1)
'WrkSht.Cells(4, 1).value = ii1
'WrkSht.Cells(4, 2).value = ii2
WrkSht.Cells(5, 1).value = 1
WrkSht.Cells(6, 1).value = 2
WrkSht.Cells(5, 2).FormulaR1C1Local = "=ЕСЛИ(R[-1]C[" & ii1 - 2 & "]=RC[" & ii1 - 2 & "];R[-1]C+1;1)"
'=RC[9]&"_"&RC[7]&"_"&RC[-1]
WrkSht.Cells(6, 2).FormulaR1C1Local = "=ЕСЛИ(R[-1]C[" & ii1 - 2 & "]=RC[" & ii1 - 2 & "];R[-1]C+1;1)"
WrkSht.Cells(5, 3).FormulaR1C1Local = "=RC[" & ii2 - 3 & "]&" & """_""" & "&RC[" & ii1 - 3 & "]&" & """_""" & "&RC[-1]"
WrkSht.Cells(6, 3).FormulaR1C1Local = "=RC[" & ii2 - 3 & "]&" & """_""" & "&RC[" & ii1 - 3 & "]&" & """_""" & "&RC[-1]"
Set oRange1 = Range(WrkSht.Cells(5, 1), WrkSht.Cells(6, 3))
Set oRange2 = Range(WrkSht.Cells(5, 1), WrkSht.Cells(i_nomer + 4, 3))
oRange1.AutoFill Destination:=oRange2
Set oRange = Range(WrkSht.Cells(5, 3), WrkSht.Cells(i_nomer + 4, j + 3 + sdvig))
oRange.name = "vybor"
ss.Delete
''WrkBk.Close True, "C:\temp\Test1.xls"
''app.Quit
app.Visible = True
Unload Me

End Sub

Private Sub CommandButton3_Click()
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
Dim blkDef As AcadBlock, xref As AcadExternalReference
Dim fType(1) As Integer, fData(1)
Dim ss As AcadSelectionSet
Dim name As String
Dim DOC As New AcadDocument
str_mess = "Загружено в Эней"

If MsgBox("Произвести выгрузку в ЭНЕЙ?", 1, "Подтверждение") = vbOK Then


Call Get_Obj(name_obj, Kod_Obj)
Unload UserForm3

UserForm4.Hide
For i = ThisDrawing.SelectionSets.Count To 1 Step -1
ThisDrawing.SelectionSets.Item(i - 1).Delete
Next
Set ss = ThisDrawing.SelectionSets.Add("ss6")

fType(0) = 0: fData(0) = "INSERT"
fType(1) = 2

For Each blkDef In ThisDrawing.Blocks
If blkDef.IsXRef Then
fData(1) = blkDef.name
ss.Select acSelectionSetAll, , , fType, fData
End If
Next
'name = "Найдено :" & vbCr
cn.ConnectionString = Connect_Str
cn.Open
tmp_str = "delete FROM fkmpobn where kodob=" & Kod_Obj
rs.Open tmp_str, cn 'Удаляем информацию по объекту
cn.Close

For Each xref In ss
' name = name & xref.Path & vbCr
Set DOC = Application.Documents.Open(xref.Path)
ReDim Attr_Name(1)
Count_BlckRf = Get_Count_BlockRef()
Call create_tbl(tbl, Attr_Name)
For i = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(i) = "Этаж" Then i_Et = i + 1 'Определяем в каком поле Этаж
Next
cn.Open
'=======================================================================
' Убрана отдельное удаление по этажам, сделано удаление всего объекта
'=======================================================================
'tmp_str = "delete FROM fkmpobn where kodob=" & Kod_Obj & " and etag1=" & tbl(0, i_Et)
'rs.Open tmp_str, cn 'Удаляем информацию по объекту и этажу
'ii = tbl(0, i_Et)
'For i = 1 To Count_BlckRf - 1
' If ii <> tbl(i, i_Et) Then
' tmp_str = "delete FROM fkmpobn where kodob=" & Kod_Obj & " and etag1=" & tbl(i, i_Et)
' rs.Open tmp_str, cn 'Удаляем информацию по объекту и этажу
' ii = tbl(i, i_Et)
' End If
'Next
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
DOC.Close
Next
ss.Delete
Unload Me
MsgBox str_mess
End If
End Sub

Private Sub UserForm_Click()

End Sub