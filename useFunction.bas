Attribute VB_Name = "useFunction"
Function getDate(ByVal s As String)

'get number from a string.
'In this case,
'we can get checkdate from folder name.
    
For i = 1 To Len(s)

    ch = mid(s, i, 1)

    If IsNumeric(ch) Then
    
        getDate = getDate & ch
    
    End If

Next

End Function

