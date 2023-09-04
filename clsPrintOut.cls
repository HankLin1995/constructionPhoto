VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrintOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private filename As String
Private sht As Object

Sub BeforePrintCheck()

'For Each rng In Sheets("Report").UsedRange

    'If rng.Interior.ColorIndex = 44 Then MsgBox "�֭p�ƶq�W�L�����ƶq�o~": End

'Next

End Sub

Sub ToPDF() '���]�ӰѦҥΪ�

sr = Val(InputBox("�}�l����"))

er = Val(InputBox("��������"))

wbpath = getThisWorkbookPath

For i = sr To er

    Sheets("Report").Activate
    
    Range("K2") = i
    
    Call ReportRun
    
        On Error Resume Next
        MkDir wbpath & "\" & "�ʳy�����Output"
        On Error GoTo 0
        'ChDir wbpath
        
        On Error Resume Next
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        wbpath & "\" & "�ʳy�����Output\" & Range("K2"), Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
        If Err <> 0 Then
            MsgBox "�A��excel�����L�¹q���L�k�t�s��PDF!" & vbNewLine & "���I��C�L�ȱi!": End
            Exit Sub
            
        End If
            
        On Error GoTo 0
        
Next

Shell "explorer.exe " & wbpath & "\" & "�ʳy�����Output\", vbNormalFocus

End Sub

Sub ShtToPDF(ByVal sht, ByVal file_name As String)

On Error Resume Next

sht.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
ThisWorkbook.path & "\�I�u�Ӥ�Output_PDF\" & file_name, Quality:= _
xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
OpenAfterPublish:=False

If Err <> 0 Then
    MsgBox "�A��excel�����L�¹q���L�k�t�s��PDF!" & vbNewLine & "���I��C�L�ȱi!": End
    Exit Sub
    
End If
    
On Error GoTo 0

End Sub

Sub ToPaper()

msg = MsgBox("�ثe���L������:" & Application.ActivePrinter & vbNewLine & "�O�_�n�~��C�L?", vbYesNo + vbInformation)

If msg = vbNo Then End

MsgBox "�o�ӫ��s�ӦM�I~���L���q�l�ɧa!"

Exit Sub

sr = Val(InputBox("�}�l����"))
er = Val(InputBox("��������"))

wbpath = getThisWorkbookPath

For i = sr To er

    Sheets("Report").Activate
    
    Range("K2") = i
        
Next

End Sub

Sub ToPDF_Check(ByVal sht As Worksheet, ByVal fname As String, ByVal mode) '�d�����X��PDF

    wbpath = getThisWorkbookPath
    
    If mode = 1 Then
    
    On Error Resume Next
    MkDir wbpath & "\" & "�d���Output"
    On Error GoTo 0
    
    On Error Resume Next
    
    sht.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
    wbpath & "\" & "�d���Output\" & fname, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False

    If Err <> 0 Then MsgBox "�A��excel�����L�¹q���L�k�t�s��PDF!" & vbNewLine & "���I��C�L�ȱi!": End
    
    Else
    
    sht.PrintOut
    Exit Sub
    End If
    
    'Shell "explorer.exe " & wbpath & "\" & "�d���Output\", vbNormalFocus
    
End Sub

'ToWorksheets

Sub ToXLS_test()

Application.DisplayAlerts = False

sr = Val(InputBox("�}�l����"))
er = Val(InputBox("��������"))

Set wb = Workbooks.Add

For r = er To sr Step -1

    Debug.Print "�C�L����=" & r

    ThisWorkbook.Activate

    ThisWorkbook.Sheets("Report").Range("K2") = r
    
    Dim obj As New clsReport
    
    obj.WriteReport_Test
    obj.hideRow
    
    ThisWorkbook.Sheets("Report").Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    For Each shp In sht.Shapes
    
        shp.Delete
        
    Next
    
    With sht
        
        .Name = r
        .Columns("A:K").Select
        .Range("A7").Activate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    End With
    
Next

wb.Sheets("�u�@��1").Delete

Application.DisplayAlerts = True

End Sub

Sub ToXLS()

Application.DisplayAlerts = False

sr = Val(InputBox("�}�l����"))
er = Val(InputBox("��������"))

Set wb = Workbooks.Add

For r = er To sr Step -1

    Debug.Print "�C�L����=" & r

    ThisWorkbook.Activate

    ThisWorkbook.Sheets("Report").Range("K2") = r
    
    Call ReportRun
    
    ThisWorkbook.Sheets("Report").Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    For Each shp In sht.Shapes
    
        shp.Delete
        
    Next
    
    With sht
        
        .Name = r
        .Columns("A:K").Select
        .Range("A7").Activate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    End With
    
Next

wb.Sheets("�u�@��1").Delete

Application.DisplayAlerts = True

End Sub

Sub SpecificShtToXLS(ByVal shtName As String, Optional file_path As String = "", Optional shtName_replace As String = "")

Application.DisplayAlerts = False

Set wb = Workbooks.Add
    
    ThisWorkbook.Sheets(shtName).Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    With sht
        
        .Name = shtName
        
        Call clearMark(sht)
        
        If shtName_replace <> "" Then .Name = shtName_replace
    
    End With

wb.Sheets("�u�@��1").Delete

If file_path <> "False" And file_path <> "" Then

    wb.SaveAs filename:=file_path, FileFormat:=xlExcel8
    wb.Close
    
End If

Application.DisplayAlerts = True

End Sub

Sub clearMark(ByVal sht)

For Each rng In sht.UsedRange

    If rng Like "<<*" Then rng.Value = ""

Next

End Sub

Sub clearMark_color(ByVal sht)

For Each rng In sht.UsedRange

    If rng Like "<<*" Then rng.Font.ColorIndex = 2
    
Next

End Sub

Sub clearMark_color_Recover(ByVal sht)

For Each rng In sht.UsedRange

    If rng Like "<<*" Then rng.Font.ColorIndex = 1
    
Next

End Sub


Sub SpecificShtObjToXLS(ByVal shtObj As Object, ByVal shtName As String, Optional file_path As String = "")

Application.DisplayAlerts = False

Set wb = Workbooks.Add
    
    shtObj.Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    With sht
        
        .Name = shtName
    
    End With

wb.Sheets("�u�@��1").Delete

If file_path <> "False" And file_path <> "" Then

    wb.SaveAs filename:=file_path, FileFormat:=xlExcel8
    wb.Close
    
End If

Application.DisplayAlerts = True

End Sub

Sub combineFiles(ByVal coll_path_sort)

Set wb_combine = Workbooks.Add

For Each filePath In coll_path_sort

    filename = mid(filePath, InStrRev(filePath, "\") + 1)
    fileExtension = mid(filename, InStrRev(filename, ".") + 1)
    
    filename_short = mid(filename, 1, InStrRev(filename, ".") - 1)
    
    Set wb_check = Workbooks.Open(filePath)
    
    For Each sht In wb_check.Sheets
    
    Debug.Print sht.Name
    
    'Set sht = wb_check.Sheets(1)
    'sht.Name = filename_short
    Application.DisplayAlerts = False
    sht.Copy before:=wb_combine.Sheets(wb_combine.Sheets.Count)
    Application.DisplayAlerts = True
    
    Next
    
    wb_check.Close False
    
    Set wb_check = Nothing
    
Next

Application.DisplayAlerts = False
wb_combine.Sheets("�u�@��1").Delete
wb_combine.Activate
Application.DisplayAlerts = True

End Sub

