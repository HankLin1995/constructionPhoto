VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtReport As Object
Private shtResult As Object
Private wb As Object

Private coll_photo_path As New Collection
Private coll_rng_address As New Collection
Private collitem As New Collection
Private collday As New Collection
Public IsXLS As Boolean

Private Sub Class_Initialize()

Application.ScreenUpdating = False

Set shtReport = ThisWorkbook.Sheets("Report")
Set shtResult = ThisWorkbook.Sheets("Result")

coll_rng_address.Add "C3"
coll_rng_address.Add "C9"
coll_rng_address.Add "C15"

End Sub

Sub GetReportByItem(addr)

If IsXLS Then Set wb = Workbooks.Add

With shtResult

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For Each it In collitem
    
        Page = 0
        Count = 0
        
        For r = 2 To lr
        
            If .Cells(r, addr) = it Then
            
                photo_path = .Cells(r, 3)
                checkitem = .Cells(r, "J")
                checkch = .Cells(r, "G") '��D�W��
                checkloc = .Cells(r, "H") '���I
                checkday = .Cells(r, "I") '���
                check_note = .Cells(r, "K")
                
                If .Cells(r, "K").Interior.ColorIndex <> -4142 Then check_note = ""
                    
                Count = Count + 1
                
                If Count = 4 Then
                
                    Count = 1
                    Page = Page + 1
                    
                    If IsXLS Then
                        Call PrintXLS(it & "-" & Page)
                    Else
                        Call printReport(it & "-" & Page)
                    End If
                    
                    Call ClearAllPhoto
                    Call ClearAllDetail

                End If
                
                Call PastePhoto(Count, photo_path, checkday)
                Call PasteDetail(Count, checkitem, checkch, checkloc, checkday, check_note)
                         
            End If
        
        Next
    
        If Count < 4 Then
            
            Page = Page + 1
            
            If IsXLS Then
                Call PrintXLS(it & "-" & Page)
            Else
                Call printReport(it & "-" & Page)
            End If
            Call ClearAllPhoto
            Call ClearAllDetail

        End If
    
    Next

End With

End Sub

Sub PrintXLS(ByVal filename As String)

shtReport.Copy After:=wb.Sheets(wb.Sheets.Count)

filename = Replace(filename, "\", "1")

wb.Sheets(wb.Sheets.Count).Name = filename

End Sub

Sub printReport(ByVal filename As String)

wbpath = ThisWorkbook.path

shtReport.Activate

On Error Resume Next
MkDir wbpath & "\" & "�d��Ӥ�Output"
On Error GoTo 0

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
wbpath & "\" & "�d��Ӥ�Output\" & filename, Quality:= _
xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
OpenAfterPublish:=False

If Err <> 0 Then
    MsgBox "�A��excel�����L�¹q���L�k�t�s��PDF!" & vbNewLine & "���I��C�L�ȱi!": End
    Exit Sub
    
End If


End Sub

Sub CollectItem(ByVal col As String)

'"G"=�d�綵��
'"J"=�d����

With shtResult

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        If .Cells(r, col) <> "" Then
            
            On Error Resume Next
            
            collitem.Add .Cells(r, col), CStr(.Cells(r, col))
            
        End If
    
    Next

End With

End Sub

Sub PasteDetail(ByVal Count As Integer, ByVal checkitem As String, ByVal checkch As String, _
                ByVal checkloc As String, ByVal checkday As String, ByVal check_note As String)

With shtReport

    Set rng = .Range(coll_rng_address(Count)).Offset(0, -1)
    rng.Offset(0, 0) = checkch
    rng.Offset(1, 0) = checkloc
    rng.Offset(2, 0) = tranDate(checkday) ' "'" & checkday
    rng.Offset(3, 0) = checkitem
    rng.Offset(4, -1) = check_note
    
End With

End Sub

Sub PastePhoto(ByVal Count As Integer, ByVal photo_path As String, ByVal checkday As String)

'�B�~���Jcheckday���@text�򩳤�r

With shtReport

    Set rng = .Range(coll_rng_address(Count))

    If rng.MergeCells Then
    
        myMergeArea = rng.MergeArea
        dy = UBound(myMergeArea, 1)
        dx = UBound(myMergeArea, 2)
    
    End If
    
    Set LastRng = .Cells(rng.Row + dy - 1, rng.Column + dx - 1).Offset(1, 1)
    
    X0 = rng.Left
    Y0 = rng.Top
    
    X1 = LastRng.Left
    Y1 = LastRng.Top
    
    Set pic = .Shapes.AddPicture(photo_path, True, True, X0 + 2, Y0 + 2, X1 - X0 - 4, Y1 - Y0 - 4)
    pic.LockAspectRatio = msoFalse
    
    Call AddText(X1 - 100, Y1 - 30, 25, 40, tranDate(checkday), 3) '�B�~���J��checkday�n��ܩ�Ӥ��k�U��
    
End With

End Sub

Sub PastePhoto_giveRng(ByVal rng As Object, ByVal photo_path As String, ByVal checkday As String)

'�B�~���Jcheckday���@text�򩳤�r

With shtReport

    'Set rng = .Range(coll_rng_address(Count))

    If rng.MergeCells Then
    
        myMergeArea = rng.MergeArea
        dy = UBound(myMergeArea, 1)
        dx = UBound(myMergeArea, 2)
    
    End If
    
    Set LastRng = .Cells(rng.Row + dy - 1, rng.Column + dx - 1).Offset(1, 1)
    
    X0 = rng.Left
    Y0 = rng.Top
    
    X1 = LastRng.Left
    Y1 = LastRng.Top
    
    Set pic = .Shapes.AddPicture(photo_path, True, True, X0 + 2, Y0 + 2, X1 - X0 - 4, Y1 - Y0 - 4)
    pic.LockAspectRatio = msoFalse
    
    Call AddText(X1 - 100, Y1 - 30, 25, 40, tranDate(checkday), 3) '�B�~���J��checkday�n��ܩ�Ӥ��k�U��
    
End With

End Sub

Function tranDate(ByVal dateStr As String)

On Error GoTo ERRORHANDLE
If IsDate(dateStr) = True Then tranDate = dateStr: Exit Function

If Len(dateStr) = 7 Then

a = mid(dateStr, 1, 3)
b = mid(dateStr, 4, 2)
c = mid(dateStr, 6, 2)

ElseIf Len(dateStr) = 8 Then

a = mid(dateStr, 1, 4)
b = mid(dateStr, 5, 2)
c = mid(dateStr, 7, 2)

Else
ERRORHANDLE:
Exit Function

End If

myDate = a & "/" & b & "/" & c

tranDate = myDate

End Function

Sub AddText(ByVal txtX, ByVal txtY, ByVal txtHeight, ByVal txtWidth, ByVal txtStr, Optional ByVal RGB_Selector As Integer = 0)

Dim o As New clsInformation
Dim r As Long
Dim g As Long
Dim b As Long
Dim r2 As Long
Dim g2 As Long
Dim b2 As Long

If o.IsPrintDate = False Then Exit Sub

'If ThisWorkbook.Sheets("Main").Range("B4") = "N" Then Exit Sub

'If txtStr = specific_date Then RGB_Selector = 3
'If txtStr = specific_date Then
    RGB_Interior = o.getInteriorColor
    Call o.VBLongToRGB(RGB_Interior, r, g, b)
    RGB_Font = o.getFontColor
    Call o.VBLongToRGB(RGB_Font, r2, g2, b2)
'End If

With ThisWorkbook.Sheets("Report")

    .Activate

    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtX, txtY, txtHeight * Len(txtStr) * 0.33, txtHeight).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = txtStr
        
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.ZOrder msoBringToFront
        ' Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font
'        .NameComplexScript = "+mn-cs"
'        .NameFarEast = "+mn-ea"
'        .Fill.Visible = msoTrue
'        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
'        .Fill.ForeColor.TintAndShade = 0
'        .Fill.ForeColor.Brightness = 0
        '.Fill.Transparency = 0
'        .Fill.Solid
        .Bold = msoTrue
        'Select Case RGB_Selector
        
        'Case 0: .Fill.ForeColor.RGB = RGB(0, 0, 0)
        'Case 1: .Fill.ForeColor.RGB = RGB(0, 204, 255)
        'Case 2: .Fill.ForeColor.RGB = RGB(255, 0, 0)
        'Case 3:
        .Fill.ForeColor.RGB = RGB(r2, g2, b2)
        
        
        'End Select
        
        '.Fill.Visible = msoFalse
        .Size = txtHeight * 0.5
'        .Name = "+mn-
    End With
    
    If o.IsPrintDateBack = True Then
    
        With Selection.ShapeRange(1).Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(r, g, b)
            '.Transparency = 0
            '.Solid
        End With

    End If

End With

End Sub

Sub ClearAllPhoto() '�M��Report�·Ӥ�

With shtReport

For Each shp In .Shapes
    
    shp.Delete
    
Next

End With

End Sub

Sub ClearAllDetail()

With shtReport
    
    arr = Array("B3", "B9", "B15")
    
    For Each ad In arr
    
        .Range(ad).Resize(4, 1).ClearContents
    
    Next

End With

End Sub

Private Sub Class_Terminate()

If IsXLS = True Then
    
    Application.DisplayAlerts = False
    wb.Sheets(1).Delete
    Application.DisplayAlerts = True
    
End If

Application.ScreenUpdating = True

End Sub