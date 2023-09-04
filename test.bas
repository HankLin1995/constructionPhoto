Attribute VB_Name = "test"


Sub extractNames()

With Sheets("Result")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    s = .Cells(r, "E")
    Call extractName(r, s)

Next

.Activate

End With

End Sub

Sub extractName(ByVal r As Integer, ByVal s As String)

's = "6_濁幹線_0+000_20230129_鋼軌樁_A.jpg"

Dim o As New clsFSO
Dim clsInf As New clsInformation

s = o.pathToFileName(s)
tmp = split(s, "_")

Set coll = clsInf.getCollStructure(clsInf.getExNameStruc)

If UBound(tmp) = coll.Count - 1 Then

    For j = LBound(tmp) To UBound(tmp)
    
    'Debug.Print tmp(j) & ":" & coll(j + 1)
    
        If tmp(j) > 5 Then
        
            Sheets("Result").Cells(r, coll(j + 1)) = tmp(j)
        
        End If
    
    Next

End If

End Sub

Sub listReports()

targetItem = Sheets("Main").ComboBox1.Value

Select Case targetItem

Case "資料夾": targetCol = 4

Case "日期", "檢查項目"

Set rng = Sheets("Result").Rows(1).Find(targetItem)

targetCol = rng.Column

End Select

If IsEmpty(targetCol) = True Then Exit Sub

With Sheets("Main")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

If lr > 8 Then .Range("A9:B" & lr).Clear

End With

Dim o As New clsFSO

Set coll_paths = o.getFilePathsInFolder(ThisWorkbook.path & "\施工照片Output\")

For Each file_path In coll_paths

    file_name = o.pathToFileName(file_path)
    
    If targetCol = 4 Then
        file_name_find = "\" & file_name & "\"
    Else
        file_name_find = file_name
    End If
    
    Set rng = Sheets("Result").Columns(targetCol).Find(file_name_find)

    If Not rng Is Nothing Then
    
        Call keyInOutputLists(file_path)
    
    End If

Next

End Sub

Sub keyInOutputLists(ByVal file_path As String)

With Sheets("Main")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

r = lr + 1

.Cells(r, 1) = r - 8
.Cells(r, 2) = file_path

.Range("A" & r & ":A" & r).HorizontalAlignment = xlCenter
.Range("B" & r & ":B" & r).HorizontalAlignment = xlRight
.Range("A" & r & ":B" & r).Borders.LineStyle = 1

End With

End Sub

Sub getPrintGroups()

Call checkDateFormat

Dim f As New clsMyfunction
Dim coll_rows_final As New Collection

targetMode = InputBox("請選擇排序方法:" & vbNewLine & "1.資料夾" & vbNewLine & "2.日期" & vbNewLine & "3.檢查項目", , 1)

Select Case targetMode

Case "1": targetCol = "D"
Case "2": targetCol = "G"
Case "3": targetCol = "J"

End Select

Set coll_folders = f.getUniqueItems("Result", 2, targetCol)

f.showList (coll_folders)

For Each folder_name In coll_folders

    Set coll_rows = f.getRowsByUser("Result", targetCol, folder_name)
    
    folder_name = Replace(folder_name, "\", "")
    
    For Each r In coll_rows
    
        With Sheets("Result")
    
            If .Cells(r, .Columns.Count).End(xlToLeft).Column > 5 Then
        
            coll_rows_final.Add r
            
            End If
    
        End With
    
    Next
    
    If coll_rows_final.Count > 0 Then Call printFilesByRows(coll_rows_final, True, folder_name)
    
    Set coll_rows_final = Nothing
    
Next

Sheets("Main").Activate

End Sub

Sub checkDateFormat()

Dim myFunc As New clsMyfunction
Dim coll As New Collection

With Sheets("Result")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        If .Cells(r, "G") <> "" Then
        
            .Cells(r, "G").Interior.ColorIndex = xlNone
        
            If myFunc.tranStrToDate(.Cells(r, "G")) = "" Then
            
                p = p & "第【" & r & "】列=【" & .Cells(r, "G") & "】...NG" & vbNewLine
                .Cells(r, "G").Interior.ColorIndex = 22
                coll.Add r
            
            End If
        
        End If
    
    Next
    
    If Len(p) > 10 Then
        MsgBox "請將工作表[Result]中日期改為【YYYYMMDD】的格式" & vbNewLine & vbNewLine & p, vbCritical
        Sheets("Result").Activate
        Sheets("Result").Cells(coll(1), "G").Select
        End
    End If

End With

End Sub

Sub printFilesByRows(ByVal coll_rows, ByVal IsXLS As Boolean, ByVal folder_name As String)

Dim o As New clsReport
Dim f As New clsMyfunction
Dim coll_recover As New Collection

Set coll_key = getKeyWords '第G欄開始

Set wb = Workbooks.Add

ThisWorkbook.Activate

For Each r In coll_rows

KeepPrint:
    i = i + 1
    
    For Each s In coll_key
        
        myKey = split(s, ",")(0)
        col = split(s, ",")(1)
    
        myAddress = getAddressByKeyWord(myKey, CStr(i))
    
        If myAddress <> "" Then
            
            If myKey = "照片" Then
            
                file_name = file_name & r & "-"
                'paste photo
                Call o.PastePhoto_giveRng(Sheets("Report").Range(myAddress), Sheets("Result").Cells(r, "C"), Sheets("Result").Cells(r, CInt(col)))
            
            ElseIf myKey = "日期" Then
                
                Sheets("Report").Range(myAddress) = f.tranStrToDate(Sheets("Result").Cells(r, CInt(col)))
            
            Else
            
                Sheets("Report").Range(myAddress) = Sheets("Result").Cells(r, CInt(col))
            
            End If
            
            coll_recover.Add "<<" & myKey & "-" & CStr(i) & ">>;" & myAddress
            IsPrinted = False
            
        Else
            
            IsPrinted = True
            If IsXLS = False Then
                Call printReportPDF(file_name)
            Else
                Call printReportToWb(wb, file_name)
            End If
            
            Call clearReport(coll_recover)
            file_name = ""
            i = 0
            GoTo KeepPrint
            
        End If
        
    Next

Next

If IsPrinted = False Then

    If IsXLS = False Then
        Call printReportPDF(file_name)
    Else
        Call printReportToWb(wb, file_name)
    End If
    
    Call clearReport(coll_recover)
    
End If

If IsXLS = True Then Call printReportToWb_Save(wb, folder_name)

End Sub

Sub printReportToWb_Save(ByVal wb, ByVal folder_name As String)

    On Error Resume Next
    
    MkDir ThisWorkbook.path & "\施工照片Output\"
    
    On Error GoTo 0

    Application.DisplayAlerts = False
    
    wb.Sheets("工作表1").Delete
    file_path = ThisWorkbook.path & "\施工照片Output\" & folder_name
    wb.SaveAs filename:=file_path, FileFormat:=xlExcel8
    wb.Close
    
    Application.DisplayAlerts = True

End Sub

Sub printReportToWb(ByVal wb, ByVal file_name As String)

Sheets("Report").Copy After:=wb.Sheets(wb.Sheets.Count)
Set sht = wb.Sheets(wb.Sheets.Count)
sht.Name = file_name

For Each rng In sht.UsedRange
    If rng Like "<<*" Then rng.Font.ColorIndex = 2
Next

ThisWorkbook.Activate

End Sub

Sub printReportPDF(ByVal file_name As String)

folder_path = ThisWorkbook.path & "\施工照片Output_PDF\"

On Error Resume Next
MkDir (folder_path)
On Error GoTo 0

Set sht = Sheets("Report")

Dim o As New clsPrintOut

o.clearMark (sht)
Call o.ShtToPDF(sht, mid(file_name, 1, Len(file_name) - 1))
o.clearMark_color_Recover (sht)

End Sub

Sub clearReport(ByVal coll_recover)

For Each shp In Sheets("Report").Shapes

    shp.Delete

Next

For Each it In coll_recover

    tmp = split(it, ";")
    rngValue = tmp(0)
    rngAddress = tmp(1)

    Sheets("Report").Range(rngAddress) = rngValue

Next

End Sub

Function getKeyWords()

Dim coll As New Collection

With ThisWorkbook.Sheets("Result")

    lc = .Cells(1, .Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lc
    
        If .Cells(1, c).Interior.ColorIndex <> -4142 Then
    
            If .Cells(1, c).Value = "日期" Then checkDayCol = c
        
            coll.Add .Cells(1, c).Value & "," & CStr(c)
    
        End If
    
    Next
    
    coll.Add "照片," & checkDayCol

End With

Set getKeyWords = coll

End Function

Function getAddressByKeyWord(ByVal myKeyWord As String, ByVal cnt As String) As String

find_text = "<<" & myKeyWord & "-" & cnt & ">>"

With ThisWorkbook.Sheets("Report")

Set rng = .Cells.Find(find_text)

If Not rng Is Nothing Then
getAddressByKeyWord = rng.Address
End If

End With

End Function
