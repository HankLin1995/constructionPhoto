VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFetchURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'HankLin edited @ 20210823
'�ϥζ��ݱ��v���XGoogle app script���@��O����
'����xlpadlock�sĶxls�ܦ�exe�ɮקY�i�F��p�յL�k�I�J������
'GAS�{���X@Ume/���ݱ��v_FINAL_ø�ϸ�Ʈw


Dim XMLHTTP As Object
Dim DOM As Object

Const URL As String = "https://script.google.com/macros/s/AKfycbwHfZ5f3L0gRwDz8d9ZvvU6kUPjkCqLETfrhH8Si4Vj9ZiZtVLIpy6_pnvRiRI6ve80/exec"

Sub Class_Initialize()

Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
Set DOM = CreateObject("Htmlfile")

End Sub

Function CreateURL(ByVal mode As String, Optional ByVal user_mac As String, Optional ByVal user_name As String, Optional ByVal user_company As String, Optional ByVal user_mail As String, Optional ByVal msg As String) ', ByVal myName As String)

Debug.Assert user_mac <> ""

'user_mac = getUnicode(user_mac)
user_name = getUnicode(user_name)
user_company = getUnicode(user_company)
msg = getUnicode(msg)

  'let mode=e.parameter.mode; //��k���
  'let mac=deUnicode(e.parameter.mac); //�b��
  'let myname=deUnicode(e.parameter.myname); //�K�X
  'let workgroup=deUnicode(e.parameter.workgroup); //�����W
  'let mymail=deUnicode(e.parameter.mymail);//�q�l�l��

CreateURL = URL & "?func=" & mode & _
                    "&mac=" & user_mac & _
                    "&myname=" & user_name & _
                    "&workgroup=" & user_company & _
                    "&mymail=" & user_mail & _
                    "&msg=" & msg
                    
End Function

Function ExecHTTP(ByVal myURL As String)

On Error GoTo ERRORHANDLE

With XMLHTTP

    .Open "GET", myURL, False
    .send
    
    'Debug.Print myURL
    
    If .Status = 200 Then
    
        tmp = split(.responsetext, "***")
        ExecHTTP = tmp(1)
        
    Else
ERRORHANDLE:
        MsgBox "�A�S���s�u!", vbCritical
        
        'clientObj.hideCmd
        
        'ThisWorkbook.Close Savechanges:=False

    End If
    
End With

End Function

Sub signup(ByVal id As String, ByVal password As String, ByVal wg As String, ByVal myName As String)

'myURL = CreateURL("1", id, password, wg, "123", myName)


If ExecHTTP(myURL) = "signed" Then
    MsgBox "�b���Ϊ̸ӹq���w�g�Q���U�L�F!", vbCritical
    
Else
    MsgBox "���U���\�A�аȥ��c�O�b��!!", vbInformation
End If

End Sub

Function checkAccessByIP()

myURL = CreateURL("2", "test", "123", "wg", "123", "test")

tmp = split(ExecHTTP(myURL), ":")

If tmp(0) <> "Pass" Then
    MsgBox "IP���ҥ���!!�еn�J�b���K�X", vbCritical
    checkAccessByIP = False
    Exit Function
Else
    'Application.StatusBar = "�]�p�̡G�L�v�~(Hank Lin) @2021," & tmp(3) & "�ϥΪ̡G" & tmp(2)
    checkAccessByIP = True
    frm_login.Label6 = "Pass"
End If

If tmp(1) <> "" Then
    MsgBox "�t�Τ��i>>" & vbNewLine & tmp(1), vbInformation
    checkAccessByIP = True
End If

End Function

Sub checkAccesByID(ByVal id As String, ByVal password As String)

myURL = CreateURL("3", id, password, "wg", "123", "123")

tmp = split(ExecHTTP(myURL), ":")

If tmp(0) <> "Pass" Then
    MsgBox "���ҥ���!!", vbCritical: ThisWorkbook.Close SaveChanges:=False
Else
    'Application.StatusBar = "�]�p�̡G�L�v�~(Hank Lin) @2021," & tmp(3) & "�ϥΪ̡G" & tmp(2)
    frm_login.Label6 = "Pass"
End If

If tmp(1) <> "" Then
    MsgBox "�t�Τ��i>>" & vbNewLine & tmp(1), vbInformation
End If

End Sub

Sub getReport(ByVal prompt)

myURL = CreateURL("4", "test", "1234", "wg", prompt, "123")

If ExecHTTP(myURL) = "ok" Then
    MsgBox "�^�X�N���w�o�e���@��!", vbInformation
Else
    MsgBox "���q�����i����U!!�гz�LGmail�o�e�N��", vbCritical
End If

End Sub

Sub getPassword(ByVal id As String)

myURL = CreateURL("5", id, "1234", "wg", "123", "123")

s = ExecHTTP(myURL)

If s <> "undefined" Then
    MsgBox "�z���K�X��" & vbNewLine & s, vbInformation
Else
    MsgBox "�䤣��K�X�A���p���@��!", vbCritical
End If

End Sub

'============method===============

Private Function GetIPAddress()
        Const strComputer As String = "."   ' Computer name. Dot means local computer
        Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
        Dim strIPAddress As String

        ' Connect to the WMI service
        Set objWMIService = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

        ' Get all TCP/IP-enabled network adapters
        Set IPConfigSet = objWMIService.ExecQuery _
            ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

        ' Get all IP addresses associated with these adapters
        For Each IPConfig In IPConfigSet
            IPAddress = IPConfig.IPAddress
            If Not IsNull(IPAddress) Then
                If InStr(1, IPConfig.Description, "WAN (", vbTextCompare) Then
                   MsgBox "���� IP = " + IPAddress(0)
                End If
                strIPAddress = strIPAddress & Join(IPAddress, "/") + vbCrLf
            End If
        Next

        GetIPAddress = strIPAddress

        'MsgBox strIPAddress
    End Function
    
Private Function getUnicode(s)

For i = 1 To Len(s)

    myUnicode = myUnicode & "," & CharToUnicode(mid(s, i, 1))

Next

getUnicode = mid(myUnicode, 2)

End Function

    
Private Function CharToUnicode(strChar As String)

    'https://stackoverflow.com/questions/4006246/getting-the-unicode-value-of-a-char-in-vb/34780532

    Dim lngUnicode As Long

    lngUnicode = AscW(strChar)

    If lngUnicode < 0 Then
        lngUnicode = 65536 + lngUnicode
    End If

    CharToUnicode = lngUnicode
    
End Function





