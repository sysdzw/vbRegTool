Attribute VB_Name = "modPub"
Option Explicit

Public reg As New RegExp
Public matchs, match

Public isShowFa  As Boolean
Public isShowSubs As Boolean
Public isShowNumber As Boolean
Public isShowOfLine As Boolean
Public strDownMode As String
Public intLanMode As Integer
Public isUseServerXMLHTTP As Boolean
Public isInsertReg As Boolean
Public isStayRegForm As Boolean

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Enum E_LAN
    eCH
    eEN
End Enum
'������Թ���
Public strAppPath As String 'Ӧ�ó���Ŀ¼

Sub Main()
    strAppPath = App.Path
    If Right(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    iniFileName = strAppPath & "REGTEST.INI"
    
    If Dir(iniFileName) <> "" Then
        Call initFromIniFile
    Else
        Call initFromApp
        Call saveToIniFile
    End If
    
    frmMain.Show
End Sub
'�������ļ���ʼ��
Private Sub initFromIniFile()
    On Error GoTo err1
    reg.IgnoreCase = CBool(GetIniS("Reg", "IgnoreCase"))
    reg.Global = CBool(GetIniS("Reg", "Global"))
    reg.MultiLine = CBool(GetIniS("Reg", "MultiLine"))
    isShowFa = CBool(GetIniS("UserSet", "ShowFa"))
    isShowSubs = CBool(GetIniS("UserSet", "ShowSubs"))
    isShowNumber = CBool(GetIniS("UserSet", "ShowNumber"))
    isShowOfLine = CBool(GetIniS("UserSet", "ShowOfLine"))
    strDownMode = Trim(GetIniS("UserSet", "DownMode"))
    intLanMode = Val(GetIniS("UserSet", "Language"))
    isInsertReg = CBool(GetIniS("UserSet", "InsertReg"))
    isStayRegForm = CBool(GetIniS("UserSet", "StayRegForm"))
    
    Exit Sub
err1:
    Call initFromApp
    Call saveToIniFile
End Sub
'���浽�����ļ�
Public Sub saveToIniFile()
    SetIniS "Reg", "IgnoreCase", CStr(reg.IgnoreCase)
    SetIniS "Reg", "Global", CStr(reg.Global)
    SetIniS "Reg", "MultiLine", CStr(reg.MultiLine)
    SetIniS "UserSet", "ShowFa", CStr(isShowFa)
    SetIniS "UserSet", "ShowSubs", CStr(isShowSubs)
    SetIniS "UserSet", "ShowNumber", CStr(isShowNumber)
    SetIniS "UserSet", "ShowOfLine", CStr(isShowOfLine)
    SetIniS "UserSet", "DownMode", CStr(strDownMode)
    SetIniS "UserSet", "Language", CStr(intLanMode)
    SetIniS "UserSet", "InsertReg", CStr(isInsertReg)
    SetIniS "UserSet", "StayRegForm", CStr(isStayRegForm)
    
End Sub
'Ӧ�ó��������ʼ��
Public Sub initFromApp()
    On Error GoTo err1
    reg.IgnoreCase = True
    reg.Global = True
    reg.MultiLine = True
    
    isShowFa = True
    isShowSubs = True
    isShowNumber = True
    isShowOfLine = False
    isInsertReg = False
    isStayRegForm = False
    strDownMode = "Normal"
    intLanMode = eCH
err1:
End Sub
Public Sub setComboHeight(oComboBox As ComboBox, lNewHeight As Long)
    Dim oldscalemode As Integer
    Dim lngLeft&, lngTop&, lngWidth&
    lngLeft = oComboBox.Left
    lngTop = oComboBox.Top
    lngWidth = oComboBox.Width
    If TypeOf oComboBox.Parent Is Frame Then Exit Sub
    oldscalemode = oComboBox.Parent.ScaleMode
    oComboBox.Parent.ScaleMode = vbPixels
    MoveWindow oComboBox.hwnd, lngLeft \ 15, lngTop \ 15, lngWidth \ 15, lNewHeight, 1
    oComboBox.Parent.ScaleMode = oldscalemode
End Sub
Public Sub setTextBoxHeight(oTextBox As TextBox, lNewHeight As Long)
    Dim oldscalemode As Integer
    Dim lngLeft&, lngTop&, lngWidth&
    lngLeft = oTextBox.Left
    lngTop = oTextBox.Top
    lngWidth = oTextBox.Width
    If TypeOf oTextBox.Parent Is Frame Then Exit Sub
    oldscalemode = oTextBox.Parent.ScaleMode
    oTextBox.Parent.ScaleMode = vbPixels
    MoveWindow oTextBox.hwnd, lngLeft \ 15, lngTop \ 15, lngWidth \ 15, lNewHeight, 1
    oTextBox.Parent.ScaleMode = oldscalemode
End Sub
'�õ���ҳԴ����(�°汾�����ֻ��������⣬������ǰ�ĺ�����������)
Public Function GetHtmlByMicrosoftXMLHTTP(strUrl$, Optional ByVal strPageType As String = "Normal") As String
    Dim XmlHttp As Object
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    
    isUseServerXMLHTTP = False
    On Error Resume Next
    XmlHttp.Open "GET", strUrl, False
    XmlHttp.SetRequestHeader "If-Modified-Since", "0"
    XmlHttp.send
    If Err.Number = "-2147024891" Then 'Microsoft.XMLHTTP�����ʾ�ܾ�����
        isUseServerXMLHTTP = True
        Set XmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
        XmlHttp.Open "GET", strUrl, False
        XmlHttp.send
    End If
    
    If strPageType = "Normal" Then
        GetHtmlByMicrosoftXMLHTTP = StrConv(XmlHttp.ResponseBody, vbUnicode)
    Else 'UTF8,big5��
        GetHtmlByMicrosoftXMLHTTP = BytesToBstr(XmlHttp.ResponseBody, strPageType)
    End If
    
    Set XmlHttp = Nothing
End Function
'תutf8,big5�ȸ�ʽ
Private Function BytesToBstr(strBody, codeBase) As String
    Dim objStream As Object
    Set objStream = CreateObject("Adodb.Stream")
    objStream.Type = 1
    objStream.Mode = 3
    objStream.Open
    objStream.Write strBody
    objStream.position = 0
    objStream.Type = 2
    objStream.Charset = codeBase
    BytesToBstr = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ������������ļ��������ļ�������
'��������fileStr
'��ڲ���(����)��
'  strFileName �������ļ�����
'����ֵ���ļ�������
'��ע��sysdzw �� 2007-5-3 �ṩ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function fileStr(ByVal strFileName As String) As String
'    On Error GoTo err1
    Dim fileHandl%
    fileHandl = FreeFile
    Open strFileName For Input As #fileHandl
    fileStr = StrConv(InputB$(LOF(fileHandl), #fileHandl), vbUnicode)
    Close #fileHandl
    Exit Function
err1:
'    MsgBox "�����ڸ��ļ�����ļ����ܷ��ʣ�", vbExclamation
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ����������ļ���������ֱ��д�ļ�
'��������writeToFile
'��ڲ���(����)��
'  strFileName �������ļ�����
'  strContent Ҫ���뵽�����ļ����ַ���
'  isCover �Ƿ񸲸Ǹ��ļ���Ĭ��Ϊ����
'����ֵ��True��False���ɹ��򷵻�ǰ�ߣ����򷵻غ���
'��ע��sysdzw �� 2007-5-2 �ṩ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo err1
    Dim fileHandl%
    fileHandl = FreeFile
    If isCover Then
        Open strFileName For Output As #fileHandl
    Else
        Open strFileName For Append As #fileHandl
    End If
    Print #fileHandl, strContent
    Close #fileHandl
    writeToFile = True
    Exit Function
err1:
    writeToFile = False
End Function

Public Function getRegMatch1(ByVal strData$, ByVal strPattern$) As String
    Dim reg As Object
    Dim matchs As Object, match As Object

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern ' "//��ַ[\s\S]*?\r\n#"
    Set matchs = reg.Execute(strData)
    For Each match In matchs
        getRegMatch1 = match.Value
    Next
End Function

Public Function getRegMatchSub1(ByVal strData$, ByVal strPattern$) As String
    Dim reg As Object
    Dim matchs As Object, match As Object

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern ' "//��ַ([\s\S]*?)\r\n#"
    Set matchs = reg.Execute(strData)
    For Each match In matchs
        getRegMatchSub1 = match.SubMatches(0)
    Next
End Function
