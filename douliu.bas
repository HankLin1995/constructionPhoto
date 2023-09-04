Attribute VB_Name = "douliu"
Sub main()

Set coll = getEachFolder

For Each i In coll

    Debug.Print i

Next

End Sub

Function getEachFolder()

Dim coll As New Collection

With Sheets("Result")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        fname = .Cells(r, 4)
    
        On Error Resume Next
        coll.Add fname, fname
        On Error GoTo 0
        
    Next
    
    Set getEachFolder = coll

End With

End Function

