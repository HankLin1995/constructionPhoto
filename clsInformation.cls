VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function IsShowEditForm()
IsShowEditForm = Sheets("Main").CheckBox1.Value
End Function

Function IsPrintDate()
IsPrintDate = Sheets("Main").CheckBox2.Value
End Function

Function IsPrintDateBack()
IsPrintDateBack = Sheets("Main").CheckBox3.Value
End Function

Function IsPasteIMG()
IsPasteIMG = Sheets("Main").CheckBox4.Value
End Function

Function getIMGwidth()

If Sheets("Main").Range("C3") = "" Then Sheets("Main").Range("C3") = InputBox("請輸入縮圖寬度:")

getIMGwidth = Sheets("Main").Range("C3")

End Function

Function getMainPath()
getMainPath = Sheets("Main").Range("B2")
End Function

Function getFontColor()

getFontColor = Sheets("Main").Range("C4").Font.color

End Function

Function getInteriorColor()

getInteriorColor = Sheets("Main").Range("C4").Interior.color

End Function

Function getReNameStruc()

getReNameStruc = Sheets("Main").Range("B5")

End Function

Function getExNameStruc()

getExNameStruc = Sheets("Main").Range("B6")

End Function


Sub VBLongToRGB(ByVal TargetColour As Long, Red As Long, Green As Long, Blue As Long)
    Red = (TargetColour Mod 256)
    Green = (TargetColour \ 256) Mod 256
    Blue = (TargetColour \ 65536) Mod 256
End Sub

Function getCollStructure(ByVal myStruc As String)

Dim coll As New Collection

'myStruc = Sheets("Main").Range("B5")

tmp = split(myStruc, "_")

For Each ch In tmp

    Set rng = Sheets("Result").Rows(1).Find(ch)
    
    If Not rng Is Nothing Then
        'Debug.Print ch & ":" & rng.Column
        coll.Add rng.Column
    Else
        MsgBox "關鍵字【" & ch & "】在工作表[Result]標頭中找不到!", vbCritical
        End
    End If

Next

Set getCollStructure = coll

End Function


