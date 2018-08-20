VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
Caption = "Ручная нумерация"
ClientHeight = 2280
ClientLeft = 45
ClientTop = 435
ClientWidth = 4395
OleObjectBlob = "UserForm2.frx":0000
StartUpPosition = 1 'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim i As Integer, i_max As Integer, j As Integer
Dim ObjEnt As AcadEntity
Dim coll_1 As New Collection
Dim blr_handle() As String
Dim blr_nomer() As Integer
Dim tmp_str As String
Dim tmp_int As Integer

i_max = UserForm2.TextBox2.Text
UserForm2.Hide

For i = 1 To i_max
On Error Resume Next
retry:
ThisDrawing.Utility.GetEntity ObjEnt, varPick, vbCr & "Выбирите элемент №" & i + UserForm2.TextBox1.Text - 1
If Err <> 0 Then
Err.Clear
If MsgBox("Повторить выбор?", 1, "Ошибка") = 1 Then GoTo retry
Exit Sub
End If
On Error GoTo 0
If ObjEnt.ObjectName = "AcDbBlockReference" Then
Call set_attr(AttrName, ObjEnt, i + UserForm2.TextBox1.Text - 1)
End If
Next

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub TextBox1_Change()

End Sub