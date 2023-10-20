Attribute VB_Name = "cmdMain"
Sub PrintOut() '列印PDF

addr = InputBox("查驗項目排序請輸入J" & vbCrLf & _
                "查驗時間排序請輸入I", , "I")

printmode = InputBox("列印成" & vbCrLf & _
                     "1.XLS" & vbCrLf & _
                     "2.PDF", , "1")

Dim objReport As New clsReport

If printmode = "1" Then
    objReport.IsXLS = True
Else
    objReport.IsXLS = False
End If

objReport.CollectItem (addr)
objReport.GetReportByItem (addr)

MsgBox "Complete!!"

End Sub

Sub getDataFromFolder()

Dim clsInf As New clsInformation
Dim objFile As New clsmyFile

With Sheets("Main")

'If MsgBox("是否要貼上縮圖?", vbYesNo) = vbYes Then
If clsInf.IsPasteIMG = True Then

    objFile.IsPaste = True
    mywidth = CInt(clsInf.getIMGwidth) '  CInt(InputBox("請輸入縮圖寬度:"))
    objFile.photo_width = mywidth / 4
    objFile.photo_height = mywidth
    
End If

objFile.main_path = .Range("B2")

objFile.delRng
objFile.getAllFolder
objFile.PastePictures

End With

Call extractNames
Call ApplyFilterToAllUsedCells

End Sub

Sub SelectFolder()

Dim objDialog As FileDialog

Set objDialog = Application.FileDialog(msoFileDialogFolderPicker)

With objDialog

    If .Show = True Then
    
        Sheets("Main").Range("B2") = objDialog.SelectedItems(1)
    
    End If

End With

End Sub

Sub ChangeAllFileName()

If Sheets("Result").AutoFilterMode Then
    Sheets("Result").AutoFilterMode = False
End If

Call getCombineNames

Dim objFile As New clsmyFile

With Sheets("Main")

    objFile.main_path = .Range("B2")
    'objFile.getRenameFile
    objFile.changeFileName

End With

Call getDataFromFolder

End Sub

Sub getCombineNames()

Dim clsInf As New clsInformation

Set coll = clsInf.getCollStructure(clsInf.getReNameStruc)

With Sheets("Result")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        If .Cells(r, .Columns.Count).End(xlToLeft).Column > 5 Then
        
            p = ""
        
            For Each col In coll
            
                p = p & .Cells(r, col) & "_"
            
            Next
            
            .Cells(r, "F") = mid(p, 1, Len(p) - 1)
        
        End If
    
    Next

End With

End Sub

Sub combineWorkbooks()

Dim f As New clsMyfunction
Dim coll_paths As New Collection
Dim coll As New Collection

For Each rng In Selection

    coll.Add rng.Row, CStr(rng.Row)

Next

For Each r In coll

    If Sheets("Main").Cells(r, 2) <> "" And r > 8 Then

        coll_paths.Add Sheets("Main").Cells(r, 2).Value

    End If

Next

Call f.showList(coll_paths)

Dim print_obj As New clsPrintOut

Call print_obj.combineFiles(coll_paths)

End Sub


Sub showFolder()

main_path = Sheets("Main").Range("B2")

Shell "explorer.exe " & main_path, vbNormalFocus

End Sub

