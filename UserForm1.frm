VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
Caption = "Меню нумерации"
ClientHeight = 3060
ClientLeft = 45
ClientTop = 435
ClientWidth = 9750.001
OleObjectBlob = "UserForm1.frx":0000
StartUpPosition = 1 'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
Dim i As Integer, i_max As Integer, j As Integer
Dim ObjEnt As AcadEntity
Dim coll_1 As New Collection

Dim tmp_str As String
Dim tmp_int As Integer


UserForm1.Hide
i_max = Get_Count_BlockRef()
UserForm2.TextBox1.Text = 1
UserForm2.TextBox2.Text = i_max
UserForm2.Show
Unload Me

End Sub

Private Sub CommandButton10_Click()

End Sub

Private Sub CommandButton2_Click()
Dim blr_handle() As String
Dim blr_nomer() As Integer
Dim i As Integer, j As Integer
Dim str_mess As String
Dim oBR As AcadBlockReference, oBR1 As AcadBlockReference


UserForm1.Hide
str_mess = "Все Ок"
Call get_array(blr_handle, blr_nomer)
Call sort_array(blr_handle, blr_nomer)
j = 0
For i = 1 To UBound(blr_handle) - 1
If blr_nomer(i) <> blr_nomer(i + 1) - 1 Then
If Len(str_mess) < 10 Then str_mess = "Найдены следующие проблемы:" & vbCr
j = j + 1
Set oBR = ThisDrawing.HandleToObject(blr_handle(i))
Set oBR1 = ThisDrawing.HandleToObject(blr_handle(i + 1))
str_mess = str_mess & j & "ошибка " & oBR.name & " №" & get_attr(AttrName, oBR) & " далее " & oBR1.name & " №" & get_attr(AttrName, oBR1) & vbCr
End If
Next
MsgBox str_mess
End Sub

Private Sub CommandButton3_Click()
Dim blr_handle() As String
Dim blr_nomer() As Integer
Dim i As Integer, i_c As Integer
Dim oBR As AcadBlockReference
If MsgBox("Произвести перенумерацию?", 1, "Подтверждение") = vbOK Then
UserForm1.Hide
Call get_array(blr_handle, blr_nomer)
Call sort_array(blr_handle, blr_nomer)
i_c = 0
For i = 1 To UBound(blr_handle)
Set oBR = ThisDrawing.HandleToObject(blr_handle(i))
Call set_attr(AttrName, oBR, i)
Next
End If
End Sub

Private Sub CommandButton4_Click()
Dim blr_handle() As String
Dim blr_nomer() As Integer
Dim i As Integer, i_max As Integer, i_nomer
Dim ObjEnt As AcadEntity
Dim oBR As AcadBlockReference
If MsgBox("Произвести перенумерацию?", 1, "Подтверждение") = vbOK Then
UserForm1.Hide
i_max = Get_Count_BlockRef()
Call get_array(blr_handle, blr_nomer)
Call sort_array(blr_handle, blr_nomer)
On Error Resume Next
retry:
ThisDrawing.Utility.GetEntity ObjEnt, varPick, vbCr & "Выбирите элемент c большим номером" & i
If Err <> 0 Then
Err.Clear
If MsgBox("Повторить выбор?", 1, "Ошибка") = 1 Then GoTo retry
Exit Sub
End If
On Error GoTo 0
If ObjEnt.ObjectName = "AcDbBlockReference" Then
i_nomer = get_attr(AttrName, ObjEnt)
End If
For i = 1 To i_max
If blr_nomer(i) > i_nomer Then
Set oBR = ThisDrawing.HandleToObject(blr_handle(i))
Call set_attr(AttrName, oBR, blr_nomer(i) + 1)
End If

Call set_attr(AttrName, ObjEnt, i_nomer + 1)
Next
End If
End Sub

Private Sub CommandButton5_Click()
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

UserForm1.Hide

Call Get_Obj(name_obj, Kod_Obj)

Count_BlckRf = Get_Count_BlockRef()

Call create_tbl(tbl, Attr_Name)

On Error Resume Next
Set app = GetObject(, "Excel.Application")
If Err Then
Err.Clear
Set app = CreateObject("Excel.Application")
If Err Then
Exit Sub
End If
End If
app.Visible = True

Set WrkBk = app.Workbooks.Add
Set WrkSht = WrkBk.Worksheets.Add

WrkSht.name = "Выгрузка"
Const sdvig = 3
With WrkSht
.Cells(1, 1 + sdvig).value = "Выгрузка из Autocad по объекту:"
.Cells(2, 1 + sdvig).value = "Какой то объект. Где брать ????"
.Cells(2, 1 + sdvig).Font.Bold = True
.Cells(3, 1 + sdvig).value = "**********************"
.Cells(4, 1 + sdvig).value = "##" & Count_BlckRf
.Cells(4, 2 + sdvig).value = "Код объекта"
.Cells(4, 3 + sdvig).value = "Наименование объекта"
.Cells(4, 4 + sdvig).value = "Наименование Блока ACAD"

End With
For i = LBound(Attr_Name) To UBound(Attr_Name)
WrkSht.Cells(4, i + 5 + sdvig).value = Attr_Name(i)
Next
For i = 0 To Count_BlckRf - 1
WrkSht.Cells(i + 5, 1 + sdvig).value = i + 1
WrkSht.Cells(i + 5, 2 + sdvig).value = Kod_Obj
WrkSht.Cells(i + 5, 3 + sdvig).value = name_obj
For j = 0 To UBound(Attr_Name) + 1
WrkSht.Cells(i + 5, j + 4 + sdvig).value = tbl(i, j)
Next
Next


Dim oRange As Range
Dim oRange1 As Range, oRange2 As Range

Set oRange = Range(WrkSht.Cells(5, 1 + sdvig), WrkSht.Cells(i + 4, j + 3 + sdvig))
oRange.Font.color = vbRed
oRange.AutoFit
ii = 0
For k = LBound(Attr_Name) To UBound(Attr_Name)
If Attr_Name(k) = "Рейс" Then ii = k + sdvig + 3
Next
oRange.Sort key1:=WrkSht.Cells(5, 7 + sdvig), key2:=WrkSht.Cells(5, 5 + sdvig)
WrkSht.Cells(4, 1).value = ii
WrkSht.Cells(5, 1).value = 1
WrkSht.Cells(6, 1).value = 2
WrkSht.Cells(5, 2).FormulaR1C1Local = "=ЕСЛИ(R[-1]C[" & ii & "]=RC[" & ii & "];R[-1]C+1;1)"
WrkSht.Cells(6, 2).FormulaR1C1Local = "=ЕСЛИ(R[-1]C[" & ii & "]=RC[" & ii & "];R[-1]C+1;1)"
WrkSht.Cells(5, 3).FormulaR1C1Local = "=RC[" & ii - 1 & "]&" & """_""" & "&RC[-1]"
WrkSht.Cells(6, 3).FormulaR1C1Local = "=RC[" & ii - 1 & "]&" & """_""" & "&RC[-1]"
Set oRange1 = Range(WrkSht.Cells(5, 1), WrkSht.Cells(6, 3))
Set oRange2 = Range(WrkSht.Cells(5, 1), WrkSht.Cells(i + 4, 3))

oRange1.AutoFill Destination:=oRange2

''WrkBk.Close True, "C:\temp\Test1.xls"
''app.Quit
Unload Me

End Sub

Private Sub CommandButton7_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim tmp_str As String, msg_str As String, tmp_str1 As String
Dim oBR As AcadBlockReference

UserForm1.Hide
cn.ConnectionString = Connect_Str
cn.Open
Dim blr_handle() As String
Dim blr_nomer() As Integer
msg_str = "Атрибуты не загружены для изделий:" & vbCr
Call get_array(blr_handle, blr_nomer)
For i = LBound(blr_handle) + 1 To UBound(blr_handle)
Set oBR = ThisDrawing.HandleToObject(blr_handle(i))
tmp_str = get_attr("Марка", oBR)
tmp_str1 = "SELECT VESIZD, VYSIZD, TLSIZD, SHRIZD, SHSL, OBIZD FROM slov2 WHERE RSHSL=""" & tmp_str & """"
rs.Open tmp_str1, cn
If rs.EOF Then
msg_str = msg_str & "#" & blr_nomer(i) & ":" & tmp_str & vbCr
Else
Call set_attr("Вес", oBR, rs.Fields(0).value)
Call set_attr("Высота", oBR, rs.Fields(1).value)
Call set_attr("Длина", oBR, rs.Fields(2).value)
Call set_attr("Ширина", oBR, rs.Fields(3).value)
Call set_attr("Код", oBR, rs.Fields(4).value)
Call set_attr("Объем", oBR, rs.Fields(5).value)
End If
rs.Close
Next
MsgBox msg_str
Unload Me
End Sub

Private Sub CommandButton8_Click()
Dim blr_handle() As String
Dim blr_nomer() As Integer
Dim oBR As AcadBlockReference
Dim cn As New ADODB.Connection
Dim tmp_str As String, msg_str As String, tmp_str1 As String
Dim WrkBk As Excel.Workbook
Dim WrkSht As Excel.Worksheet
Dim i As Integer, j As Integer, z As Integer

UserForm1.Hide

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
On Error Resume Next
Set app = GetObject(, "Excel.Application")
If Err Then
Err.Clear
Set app = CreateObject("Excel.Application")
If Err Then
Exit Sub
End If
End If
app.Visible = True
Set WrkBk = app.Workbooks.Add
Set WrkSht = WrkBk.Worksheets.Add
WrkSht.name = "Проблемы"
j = 2
WrkSht.Cells(1, 1).value = "Номер монтажа"
WrkSht.Cells(1, 2).value = "Изделие"
msg_str = "Найдены проблемы" & vbCr
End If
z = z + 1
With WrkSht
.Cells(j, 1).value = blr_nomer(i)
.Cells(j, 2).value = tmp_str
j = j + 1
End With
'' msg_str = msg_str & "#" & blr_nomer(i) & ":" & tmp_str & vbCr
End If
rs.Close

Next
MsgBox msg_str
Unload Me
End Sub

Private Sub CommandButton9_Click()
UserForm4.Show
Unload Me
End Sub