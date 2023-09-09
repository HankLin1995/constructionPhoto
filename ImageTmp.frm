VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImageTmp 
   Caption         =   "施工照片大圖資訊"
   ClientHeight    =   9105.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16035
   OleObjectBlob   =   "ImageTmp.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ImageTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox1_Change()
Cells(TextBox1.Value + 1, 10) = Me.ComboBox1.Value
End Sub

Private Sub CommandButton1_Click()
TextBox1.Value = CInt(TextBox1.Value) + 1
End Sub

Private Sub TextBox1_Change()

Dim clsI As New clsInformation

If clsI.IsShowEditForm = False Then Exit Sub

On Error Resume Next

If TextBox1.Value <> "0" Then

s = Cells(TextBox1.Value + 1, 3)

'On Error Resume Next

Image1.Picture = LoadPicture(s)

ImageTmp.TextBox2 = Cells(TextBox1.Value + 1, 7)
ImageTmp.TextBox3 = Cells(TextBox1.Value + 1, 8)
ImageTmp.TextBox4 = Cells(TextBox1.Value + 1, 9)
ImageTmp.ComboBox1 = Cells(TextBox1.Value + 1, 10)
'ImageTmp.TextBox5 = Cells(TextBox1.Value + 1, 10)
ImageTmp.TextBox6 = Cells(TextBox1.Value + 1, 11)

End If

End Sub

Private Sub TextBox2_AfterUpdate()

Dim myFunc As New clsMyfunction

If myFunc.tranStrToDate(TextBox2.Value) = "" Then
    MsgBox "請確認日期格式為20230904!", vbCritical
    TextBox2.Value = ""
    Exit Sub
End If

r = TextBox1.Value + 1

Cells(r, 7) = TextBox2.Value

End Sub

Private Sub TextBox3_AfterUpdate()

r = TextBox1.Value + 1

Cells(r, 8) = TextBox3.Value

End Sub

Private Sub TextBox4_AfterUpdate()

r = TextBox1.Value + 1

Cells(r, 9) = TextBox4.Value

End Sub

Private Sub TextBox5_AfterUpdate()

r = TextBox1.Value + 1

Cells(r, 10) = TextBox5.Value

End Sub

Private Sub TextBox6_AfterUpdate()

r = TextBox1.Value + 1

Cells(r, 11) = TextBox6.Value

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Initialize()

With Sheets("Result")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    Me.Label2.Caption = "共" & lr - 1 & "張"
    
    Me.Label3.Caption = .Cells(1, 7)
    Me.Label4.Caption = .Cells(1, 8)
    Me.Label5.Caption = .Cells(1, 9)
    Me.Label6.Caption = .Cells(1, 10)
    Me.Label7.Caption = .Cells(1, 11)
    
    Me.ComboBox1.AddItem "查驗"
    Me.ComboBox1.AddItem "施工中"
    Me.ComboBox1.AddItem "缺失"

End With

End Sub


