VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmyFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private coll_photo_path As New Collection

Private main_folder As String
Public main_path As String
Public photo_width As Double
Public photo_height As Double
Public IsPaste As Boolean

Private shtMain As Object
Private shtResult As Object
Private objFso As Object

Private Sub Class_Initialize()

Set shtMain = Sheets("Main")
Set shtResult = Sheets("Result")

Set objFso = CreateObject("Scripting.FileSystemObject")

End Sub

Sub getAllFolder()

Dim objFolder As Object

Set objFolder = objFso.GetFolder(main_path)

Call getAllFile(objFolder)

End Sub

Private Sub getAllFile(ByVal objFolder As Object)

If objFolder.Files.Count <> 0 Then

    For Each objFile In objFolder.Files
    
    s = objFso.GetExtensionName(objFile)
    
        If IsPhoto(s) Then coll_photo_path.Add objFile.path
        
    Next

End If

If objFolder.subFolders.Count <> 0 Then

    For Each objSubFolder In objFolder.subFolders
    
        Call getAllFile(objSubFolder)
    
    Next

End If

End Sub

Private Function IsPhoto(ByVal s As String)

Select Case UCase(s)

Case "JPG"
    IsPhoto = True
Case Else
    IsPhoto = False
End Select

End Function

Sub delRng()

tmp = split(main_path, "\")
main_folder = tmp(UBound(tmp))

With shtResult

'.Range("A1:F1") = Array("ID", "縮圖", "全路徑", "資料夾(" & main_folder & ")", "檔名", "更改檔名")

.Range("D1") = "資料夾(" & main_folder & ")"

lr = .Cells(1, 1).End(xlDown).Row
lc = .Cells(1, 1).End(xlToRight).Column - 1

Set rng_last = .Cells(lr, lc)

'Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)

'lr = .Cells(.Rows.Count, 1).End(xlUp).Row
'lc = .Cells(1, .Columns.Count).End(xlToLeft).Column

If rng_last.Row > 1 Then .Range("A2").Resize(rng_last.Row, rng_last.Column).Clear 'Contents

For Each Shape In .Shapes
    
    'Debug.Print Shape.Address
    
    If Shape.TopLeftCell.Row <> 1 Then Shape.Delete

Next

End With

End Sub

Sub PastePictures()

With shtResult
    
    lr = .Cells(1, 1).End(xlDown).Row
    lc = .Cells(1, 1).End(xlToRight).Column - 1
    
    Set rng_last = .Cells(lr, lc)
    
    'Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)
    
    r = 2
    
    For Each photo_path In coll_photo_path
    
        If IsPaste = True Then
    
            Set objTargetCell = .Cells(r, 2)
    
            objTargetCell.ColumnWidth = photo_width
            objTargetCell.RowHeight = photo_height
        
            Set pic = .Pictures.Insert(photo_path)
            'Set pic = .Pictures.Insert(imagePath).ShapeRange.PictureFormat.Compression = 75
            'Set pic = .Shapes.AddPicture(photo_path) ', True, True, objTargetCell.Left + 2, objTargetCell.Top + 2, objTargetCell.Width - 4, objTargetCell.Height - 4)
            'pic.LockAspectRatio = msoFalse
        
            With pic
            
                dblGap = 2#
            
                .ShapeRange.LockAspectRatio = msoTrue '鎖定照片長寬比
                
                Set ranOri = objTargetCell

                dblRatioPic = .Width / .Height
                dblRatioOri = ranOri.Width / ranOri.Height
                
                If dblRatioPic > dblRatioOri Then '寬度控制
                    .Width = ranOri.Width - 2 * dblGap
                    .Top = ranOri.Top + 0.5 * ranOri.Height - 0.5 * .Height
                    .Left = ranOri.Left + dblGap
                Else                                '高度控制
                    .Height = ranOri.Height - 2 * dblGap
                    .Top = ranOri.Top + dblGap
                    .Left = ranOri.Left + 0.5 * ranOri.Width - 0.5 * .Width
                    
                End If
            
            End With
        
        End If
        
        
        .Cells(r, 1) = r - 1
        .Cells(r, 3) = photo_path
        .Cells(r, 4) = getParentFolder(photo_path)
        .Cells(r, 5) = getFileName(photo_path)
        
        .Cells(r, 1).Resize(1, rng_last.Column).Borders.LineStyle = 1
        .Cells(r, 6).Resize(1, rng_last.Column).HorizontalAlignment = xlLeft
        
        'Call getMarkDetail(.Cells(r, 5), r)
        
        
        
        r = r + 1
    
    Next

End With

End Sub
'
'Private Sub getMarkDetail(ByVal filename As String, ByVal r As Integer)
'
'tmp = split(filename, "_")
'
'If UBound(tmp) >= 3 Then
'
'    checkch = tmp(0 + 1)
'    checkloc = tmp(1 + 1)
'    checkday = tmp(2 + 1)
'    checkitem = tmp(3 + 1)
'    checknote = tmp(4 + 1)
'
'    With shtResult
'
'        .Cells(r, "G") = checkch
'        .Cells(r, "H") = checkloc
'        .Cells(r, "I") = checkday
'        .Cells(r, "J") = checkitem
'        .Cells(r, "K") = split(checknote, ".")(0)
'
'    End With
'
'End If
'
'End Sub

Function getFileName(ByVal path As String)

tmp = split(path, "\")

getFileName = tmp(UBound(tmp))

End Function

Function getParentFolder(ByVal path As String)

tmp_folder = Replace(path, main_path, "")

tmp = split(tmp_folder, "\")

tmp_filename = tmp(UBound(tmp))

getParentFolder = Replace(tmp_folder, tmp_filename, "")

End Function

Sub changeFileName()

With shtResult

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        oldname = .Cells(r, 3)
        
        If .Cells(r, 6) Like "*.jp*" Then
            newname = main_path & .Cells(r, 4) & .Cells(r, 6)
        Else
            newname = main_path & .Cells(r, 4) & .Cells(r, 6) & ".jpg"
        End If
    
        If .Cells(r, 6) <> "" Then
            
            Debug.Print "old:" & oldname
            Debug.Print "new:" & newname
            
            Name oldname As newname
            
        End If
    Next

End With

End Sub

Sub getRenameFile()

With shtResult

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        checkitem = .Cells(r, "J")
    
        If checkitem <> "" Then
        
            checkch = .Cells(r, "G")
            checkloc = .Cells(r, "H")
            checkdate = .Cells(r, "I")
            'checkitem = .Cells(r, "J")
            checknote = .Cells(r, "K")
        
            '.Cells(r, "F") = checkch & "_" & checkloc & ".JPG"
            arr = Array(checkch, checkloc, checkdate, checkitem, checknote)
            .Cells(r, "F") = Join(arr, "_")
        
        End If
    
    Next

End With

End Sub

Private Sub Class_Terminate()

'msg = MsgBox("是否要顯示目標資料夾", vbYesNo)
'
'If msg = vbYes Then Shell "explorer.exe " & main_path, vbNormalFocus
'
'With shtResult
'
'    .Activate
'    .Columns("C").Hidden = True
'    .Columns("D:E").AutoFit
'
'End With

End Sub
