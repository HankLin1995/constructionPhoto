Attribute VB_Name = "UnitTest"


Sub test_Main()

Call test_tranDateToStr
Call test_tranStrToDate

MsgBox "UnitTest PASS!", vbInformation

End Sub

Sub test_renameStrcu()

Dim o As New clsInformation
Dim f As New clsMyfunction

Set coll = o.getCollStructure

f.showList (coll)

End Sub

Sub test_RGB()

Dim o As New clsInformation
Dim r As Long
Dim g As Long
Dim b As Long
Dim r2 As Long
Dim g2 As Long
Dim b2 As Long

RGB_Interior = o.getInteriorColor
Call o.VBLongToRGB(RGB_Interior, r, g, b)
RGB_Font = o.getFontColor
Call o.VBLongToRGB(RGB_Font, r2, g2, b2)

Range("D4").Interior.color = RGB(r, g, b)
Range("D4").Font.color = RGB(r2, g2, b2)

End Sub

Sub Colors()
    Dim color As Long, r As Long, g As Long, b As Long
    color = Range("C4").Interior.color
    r = color Mod 256
    g = color \ 256 Mod 256
    b = color \ 65536 Mod 256
    MsgBox "RGB (" & r & ", " & g & ", " & b & ")"
End Sub

Sub test_clsFSO_kill()

folder_path = ThisWorkbook.path & "\施工照片Output\"

Dim o As New clsFSO

o.killFilesInFolder (folder_path)

End Sub


Sub test_clsFSO_folderExist()

Dim clsI As New clsInformation

folder_name = clsI.getMainPath & "\" & "1120130-鋼板樁打設抽查" & "\"

Dim o As New clsFSO

Debug.Assert o.IsFolderExists(folder_name) = True

End Sub

Sub test_tranDateToStr()

Dim f As New clsMyfunction

s = CDate("2023/8/19")

Debug.Assert f.tranDateToStr(s) = "20230819"

s = CDate("112/8/19")

Debug.Assert f.tranDateToStr(s) = "20230819"

End Sub

Sub test_tranStrToDate()

Dim f As New clsMyfunction

s = "20230819"

Debug.Assert f.tranStrToDate(s) = CDate("2023/08/19")

s = "1120819"

Debug.Print f.tranStrToDate(s) = CDate("112/08/19")

End Sub


