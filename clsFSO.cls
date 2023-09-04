VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fso As Object
Private folder As Object
Private file As Object

Private Sub Class_Initialize()

Set fso = CreateObject("Scripting.FileSystemObject")

End Sub

Function getFilePathsInFolder(ByVal folderPath As String)

    Dim coll As New Collection
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files

        coll.Add file.path
        
    Next file
    
    Set getFilePathsInFolder = coll
    
End Function

Sub killFilesInFolder(ByVal folder_path)

Dim clsInf As New clsInformation

Set coll_path = getFilePathsInFolder(folder_path)

For Each file_path In coll_path

    'filename = mid(file_path, InStrRev(file_path, "\") + 1)
    'filename = mid(filename, 1, InStrRev(filename, ".") - 1)
    
    filename = pathToFileName(file_path)
    
    If IsFolderExists(clsInf.getMainPath & "\" & filename & "\") = False Then
    
        Kill file_path
    
    End If

Next

End Sub

Function pathToFileName(ByVal file_path As String)

    filename = mid(file_path, InStrRev(file_path, "\") + 1)
    pathToFileName = mid(filename, 1, InStrRev(filename, ".") - 1)

End Function

Function IsFolderExists(ByVal folder_name As String)

IsFolderExists = fso.FolderExists(folder_name)

End Function



