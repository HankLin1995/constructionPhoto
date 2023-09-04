Attribute VB_Name = "PastePic"
Option Explicit
Public Sub PastePic()
Dim fleTmp          As file
Dim fldMain         As folder
Dim fsoMain         As New FileSystemObject
Dim shtAct          As Worksheet
Dim picAct          As Picture
Dim strFldPath      As String
Dim strTmp          As String
Dim Count           As Integer
Dim i               As Integer
Dim j               As Integer
Dim ranOri          As Range
Dim ranNum          As Range
Dim dblRatioPic     As Double
Dim dblRatioOri     As Double
Dim dblGap          As Double

dblGap = 2#
strFldPath = GetFolder
If Not fsoMain.FolderExists(strFldPath) Then
    MsgBox "��Ƨ����s�b!", vbExclamation + vbOKOnly
    Set fsoMain = Nothing
    Exit Sub
Else
    Set fldMain = fsoMain.GetFolder(strFldPath)
End If
'====�p��ۤ��i��===========
For Each fleTmp In fldMain.Files
    strTmp = UCase(fsoMain.GetExtensionName(fleTmp.path))
    If strTmp = "JPG" Or strTmp = "JPEG" Then
        Count = Count + 1
    End If
Next fleTmp
If Count < 1 Then
    MsgBox "��Ƨ��U�L*.JPG�B*.JPEG���Ӥ�!", vbExclamation + vbOKOnly
    Set fsoMain = Nothing
    Exit Sub
End If

'====�ƻs�u�@��===========
ThisWorkbook.Worksheets("���d��").Copy ' After:=Sheets(Sheets.Count)
Set shtAct = ActiveSheet
shtAct.Name = "�Ӥ����"
With shtAct
    .Shapes.Range("Button 1").Delete
    Set ranOri = .Range(.Cells(1, 1), .Cells(28, 5))
    .Range("E26") = 1
    j = Count \ 2
    If Count Mod 2 = 0 Then j = j - 1
    For i = 1 To j
        ranOri.Copy .Range(.Cells(1 + 28 * i, 1), Cells(1 + 28 * i, 1))
        .Rows(1 + 28 * i).RowHeight = 50.75
        .Rows(1 + 28 * i + 14).RowHeight = 8
        .Range(.Cells(1 + 28 * i + 25, 5), Cells(1 + 28 * i + 25, 5)) = i + 1
    Next i

End With
'====�K�ۤ�===============
Count = 0
For Each fleTmp In fldMain.Files
    strTmp = UCase(fsoMain.GetExtensionName(fleTmp.path))
    If strTmp = "JPG" Or strTmp = "JPEG" Then
        Count = Count + 1
        With shtAct
            Set ranOri = .Range(.Cells(2 + (Count - 1) * 14, 4), .Cells(2 + (Count - 1) * 14, 4))
            Set ranNum = ranOri.Offset(2, -2)
            ranNum = Count
        End With
        Set picAct = shtAct.Pictures.Insert(fleTmp.path)
        With picAct
            .ShapeRange.LockAspectRatio = msoTrue '��w�Ӥ����e��
            If .Height > .Width Then '����
                Set ranOri = ranOri.Resize(13, 1) '.Merge
                ranOri.Merge
            Else '�
                Set ranOri = ranOri.Offset(4, -2).Resize(9, 3) '.Merge
                ranOri.Merge
            End If
            dblRatioPic = .Width / .Height
            dblRatioOri = ranOri.Width / ranOri.Height
            If dblRatioPic > dblRatioOri Then '�e�ױ���
                .Width = ranOri.Width - 2 * dblGap
                .Top = ranOri.Top + 0.5 * ranOri.Height - 0.5 * .Height
                .Left = ranOri.Left + dblGap
            Else                                '���ױ���
                .Height = ranOri.Height - 2 * dblGap
                .Top = ranOri.Top + dblGap
                .Left = ranOri.Left + 0.5 * ranOri.Width - 0.5 * .Width
            End If
        End With
    End If
Next fleTmp
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Done!", vbInformation + vbOKOnly

Set picAct = Nothing
Set ranOri = Nothing
Set ranNum = Nothing
Set shtAct = Nothing
Set fleTmp = Nothing
Set fldMain = Nothing
Set fsoMain = Nothing
End Sub
Private Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "��ܷӤ���Ƨ�"
        .AllowMultiSelect = False
        .initialFilename = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function


