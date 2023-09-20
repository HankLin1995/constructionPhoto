Attribute VB_Name = "FetchURL"
Sub FetchURL_Main()

MAC_ADDRESS = getMacAddress

Debug.Print MAC_ADDRESS

If MAC_ADDRESS = "" Then
    Debug.Print "�нT�{�����O�_���s�W����~!"
    ThisWorkbook.Close False
Else
    
    Status = AccessStatus(MAC_ADDRESS)
    
    If Status = "" Then
        Call stopAccess
    End If
    
    myChoose = split(Status, ",")

    'If Status <> "PASS" Then
    If myChoose(0) <> "PASS" Then
    
        MsgBox "�{�����g���v�A�Y�N����!", vbCritical
        'MsgBox "�z���ϥΦ��Ƥw��," & vbNewLine & "�p���~��ϥνХ[HankLin��LINE", vbInformation
        
        'UserForm2.Show
        
        Call stopAccess
        
    Else
    
        'MsgBox "���v���\!" & vbNewLine & "�եΪ��Ѿl�Ѽơi" & myChoose(1) & "�j��" & vbNewLine & "�p���ϥΤW�����D�Х[HankLin��LINE", vbInformation
        
        'Call ShowDialoge(myChoose(3)) ', myChoose(3))
        
        
        On Error GoTo ERRORHANDLE
        
        If myChoose(2) <> "" Then MsgBox "***�ӤH���i***" & vbNewLine & vbNewLine & "�y" & myChoose(2) & "�z"
        If myChoose(3) <> "" Then MsgBox "***�t�Τ��i***" & vbNewLine & vbNewLine & "�y" & myChoose(3) & "�z"
        
ERRORHANDLE:
        
    End If

End If

End Sub

Function AccessStatus(ByVal mac_add As String) As String

KEEPACCESS:

Dim o As New clsFetchURL
Dim bIsClientSigned As Boolean

myURL = o.CreateURL("Access", mac_add)
Status = o.ExecHTTP(myURL)
On Error GoTo ERRORHANDLE
myChoose = split(Status, ",")

Select Case myChoose(0) 'Status

Case "PASS"

    Debug.Print "���ҳq�L!"
    
Case "NOT_FOUND"

    Debug.Print "�䤣���Ʈw���A�������Ǹ�"

    bIsClientSigned = IsClientSigned(mac_add) '�i����U���ըæ^�ǵ��G
    
    If bIsCliendSigned = False Then GoTo KEEPACCESS

Case "ARRIVED":

'        myStd = DesktopZoom()
'
'        Call System.Workbook_Open2(X, Y)
'
'        myHeight = Y * 0.2 * 100 / myStd
'        myWidth = X * 0.2 * 100 / myStd
'
'        Debug.Print myHeight & ":" & myWidth
'
'        UserForm1.Height = myHeight
'        UserForm1.Width = myWidth
'        UserForm1.Image1.Height = myHeight
'        UserForm1.Image1.Width = myWidth
         'UserForm1.Show

    'MsgBox "������ϥΤѼƬ�0��A�p�G�n�ϥν��ʶR������!", vbInformation

Case Else

    Debug.Print Status

End Select

AccessStatus = Status

Exit Function

ERRORHANDLE:

End Function

Function IsClientSigned(ByVal mac_add As String) As Boolean

'�i����U�æ^�ǵ��U���A
'1.�w���U:True
'2.���U�q�L:False

IsClientSigned = False

Dim o As New clsFetchURL

myURL = o.CreateURL("Sign", mac_add)

If o.ExecHTTP(myURL) = "signed" Then
    MsgBox "�ӹq���w�g�Q���U�L�F!", vbCritical
    IsClientSigned = True
Else

MsgBox "������z���Ĥ@���ϥΪ��ϥΪ̡A�д��ѥH�U��ƨѰ�", vbInformation

myName = InputBox("�п�J�z���m�W!")
myJob = InputBox("�п�J�z�����q�W��!")
myMail = InputBox("�п�J�z��Email�b��!")

'Do Until myMail <> ""
'
'    myMail = InputBox("�п�J�i���`�ϥΤ�Email�b��!")
'
'Loop

myURL = o.CreateURL("SignDetail", mac_add, myName, myJob, myMail)
Call o.ExecHTTP(myURL)

MsgBox "���±z���t�X!!��������D�A��x���^��~", vbInformation

'Sheets("�a�_��ø��").Range("E1") = myMail

End If

End Function

Sub t()

Call ShowDialoge("ee")

End Sub


Sub ShowDialoge(ByVal s1 As String) ', ByVal s2 As String)

If s1 = "" And s2 = "" Then Exit Sub

UserForm3.TextBox1.Value = s1
'UserForm3.TextBox2.Value = s2
UserForm3.Show


End Sub

Sub stopAccess()

MsgBox "STOP"
End

Exit Sub

ThisWorkbook.Close False
Application.Quit

End Sub

'======�o�̽T�w���|��==========

Function getMacAddress()

Dim objVMI As Object
Dim vAdptr As Variant
Dim objAdptr As Object
'Dim adptrCnt As Long


Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objAdptr In vAdptr
    If Not IsNull(objAdptr.MACAddress) And IsArray(objAdptr.IPAddress) Then
        For adptrCnt = 0 To UBound(objAdptr.IPAddress)
        If Not objAdptr.IPAddress(adptrCnt) = "0.0.0.0" Then
            GetNetworkConnectionMACAddress = objAdptr.MACAddress
            Exit For
        End If
        Next
    End If
Next

getMacAddress = GetNetworkConnectionMACAddress

End Function




