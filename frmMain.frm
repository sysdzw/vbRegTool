VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9915
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   2640
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2160
      Top             =   720
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   22
      Top             =   1800
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtKeyword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ËÑ Ë÷"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   25
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Height          =   495
         Left            =   1920
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   840
      MousePointer    =   7  'Size N S
      Picture         =   "frmMain.frx":1272
      ScaleHeight     =   90
      ScaleWidth      =   9015
      TabIndex        =   21
      ToolTipText     =   "°´×¡ÍÏ¶¯¿Éµ÷Õû´°Ìå´óÐ¡"
      Top             =   3480
      Width           =   9015
   End
   Begin VB.TextBox txtReplace 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5640
      TabIndex        =   7
      Top             =   4012
      Width           =   1575
   End
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Flat
      Height          =   165
      Left            =   840
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "http://www.baidu.com"
      Top             =   75
      Width           =   3855
   End
   Begin VB.ComboBox cboUrl 
      BackColor       =   &H8000000E&
      Height          =   300
      Left            =   840
      TabIndex        =   20
      Text            =   "http://www.baidu.com"
      Top             =   60
      Width           =   4215
   End
   Begin VB.ComboBox cboReplace 
      Height          =   300
      Left            =   5640
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   3997
      Width           =   1935
   End
   Begin VB.ComboBox cboPattern 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Text            =   "[a-zA-z]+://[^""\s>\)]*"
      ToolTipText     =   "°´Ctrl+R¿ÉÒÔµ÷³öÕýÔò±í´ïÊ½Å¶"
      Top             =   3600
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "²âÊÔ"
      Height          =   1095
      Left            =   1320
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox chkDelTab 
         BackColor       =   &H008080FF&
         Caption         =   "É¾³ýÖÆ±í·û"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox chkDelHuiche 
         BackColor       =   &H008080FF&
         Caption         =   "É¾³ý»Ø³µ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   720
   End
   Begin RichTextLib.RichTextBox txtResult 
      Height          =   1695
      Left            =   840
      TabIndex        =   9
      Top             =   4440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":2E46
   End
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   3015
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Ë«»÷¿ìËÙµ÷³öÉ¾³ý»Ø³µµÄ´°¿Ú£»°´Ctrl+F´ò¿ª²éÕÒ×Ö·û´®µÄ´°¿Ú"
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   1
      TextRTF         =   $"frmMain.frx":2EE3
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "È«²¿Ìæ»»Îª"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "ÉèÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "¼ìË÷½á¹û"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "ÏÂÔØ´úÂë¡ý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6120
      TabIndex        =   1
      Top             =   45
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "ÍêÈ«Æ¥Åä"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "µã»÷´ÎÊý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "½á¡¡¹û:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Íø¡¡Ö·:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   12
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "±í´ïÊ½:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   11
      Top             =   3600
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "×Ö·û´®:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   480
      Width           =   585
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intClickTimes%
Private Const LAN_CH = "ÕýÔò²âÊÔ¹¤¾ß    ,Íø¡¡Ö·:,ÔØÈëÊý¾Ý¡ý ,×Ö·û´®:,±í´ïÊ½: ,½á¡¡¹û:,ÍêÈ«Æ¥Åä,¼ìË÷½á¹û    ,È«²¿Ìæ»»Îª    ,ÏµÍ³ÉèÖÃ"
Private Const LAN_EN = "RegExp Test Tool,URL:   ,&Load Data¡ý,String:,Pattern:,Result:,&Test   ,Show &Result,Replace &All,&Set"
Private vCH, vEN
Dim isMove As Boolean

Private Sub cboPattern_Click()
    cboPattern.Tag = cboPattern.SelStart
End Sub

Private Sub Command1_Click()
    MsgBox Picture1.Tag
    txtResult.Text = Picture1.Tag
End Sub

Private Sub Form_Activate()
    setComboHeight cboUrl, 400
    setComboHeight cboPattern, 400
    setComboHeight cboReplace, 400
    cboPattern.SelStart = Len(cboPattern.Text)
    
'    setTextBoxHeight txtUrl, 20
'    setTextBoxHeight txtReplace, 20
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If Chr(KeyCode) = "R" Then '°´Ctrl+Rµ÷³öÕýÔò¿â
            frmReg.Show
        ElseIf Chr(KeyCode) = "F" Then '°´Ctrl+Fµ÷³ö²éÕÒ×Ö·û´®µÄ¶Ô»°¿ò
            Frame2.Visible = True
            txtKeyword.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Picture1.Tag = Picture1.Top / Me.ScaleHeight
    intClickTimes = 1

    Call initToForm
    Call setLanMain 'ÉèÖÃ½çÃæÓïÑÔ
End Sub

Private Sub Form_Click()
    Dim strCode$, strData$, vData, strPattern$, strReplace$, strReplacePart$, j&, intLine%, strTmp$, strXmlHttp$
    
    Timer1.Enabled = True
    intClickTimes = intClickTimes + 1
    
    Label3.Caption = intClickTimes
    If intClickTimes = 3 Then
        strData = txtSource.Text
        strData = Replace(strData, """", """""")
        If Right(strData, 2) = vbCrLf Then strData = Left(strData, Len(strData) - 2)
        vData = Split(strData, vbCrLf)
        strData = ""
        
        If txtUrl.BackColor = &HC0C0FF And txtUrl.Text <> "" Then
            strXmlHttp = IIf(isUseServerXMLHTTP, "MSXML2.ServerXMLHTTP", "Microsoft.XMLHTTP")
            
            If isFileExists(txtUrl.Text) Then
                strData = "fileStr(""" & txtUrl.Text & """)"
            Else
                strData = "getHtmlStr(""" & txtUrl.Text & """)"
            End If
        Else
            For j = 0 To UBound(vData)
                strTmp = Trim(vData(j))
                If strTmp <> "" Then
                    intLine = intLine + 1
                    If intLine = 1 Then
                        strData = """" & vData(j) & """  &  vbCrLf  & _" & vbCrLf
                    Else
                        strData = strData & Space(14) & """" & vData(j) & """  &  vbCrLf  & _" & vbCrLf
                    End If
                End If
                If intLine > 19 Then Exit For
            Next
            
            If strData <> "" Then strData = Left(strData, Len(strData) - Len("  &  vbCrLf  & _" & vbCrLf))
        End If
        
        If intLanMode = eCH Then
            'strCode = "'´Ë´úÂëÓÉ¡°[url=http://blog.csdn.net/sysdzw/article/details/6141844][color=#000000]ÕýÔò²âÊÔ¹¤¾ß [/color][/url] v" & App.Major & "." & App.Minor & "." & App.Revision & "¡±×Ô¶¯Éú³É£¬ÇëÖ±½Óµ÷ÓÃTestReg¹ý³Ì" & vbCrLf
            strCode = "'´Ë´úÂëÓÉ¡°ÕýÔò²âÊÔ¹¤¾ßV" & App.Major & "." & App.Minor & "." & App.Revision & "¡±×Ô¶¯Éú³É£¬ÇëÖ±½Óµ÷ÓÃTestReg¹ý³Ì" & vbCrLf
        Else
            strCode = "'This code was generated by ""RegTestTool v" & App.Major & "." & App.Minor & "." & App.Revision & """, please call the sub TestReg" & vbCrLf
        End If
        
        strCode = strCode & "Private Sub TestReg()" & vbCrLf & _
                  "    Dim strData As String" & vbCrLf & _
                  "    Dim reg As Object" & vbCrLf
                
        If txtReplace.BackColor <> &HC0C0FF Then
            strCode = strCode & "    Dim matchs As Object, i As Integer, j As Integer" & vbCrLf
        End If
        
        If strData <> "" Then
            strCode = strCode & "    strData = " & strData & "" & vbCrLf
        End If
            
        strPattern = cboPattern.Text
        strPattern = Replace(strPattern, """", """""")
        strReplace = txtReplace.Text
        strReplace = Replace(strReplace, """", """""")
        strCode = strCode & "    Set reg = CreateObject(""vbscript.regExp"")" & vbCrLf & _
                "    reg.Global = " & CStr(reg.Global = True) & vbCrLf & _
                "    reg.IgnoreCase = " & CStr(reg.IgnoreCase = True) & vbCrLf & _
                "    reg.MultiLine = " & CStr(reg.MultiLine = True) & vbCrLf
        
        '/////Ìæ»»¼°ÏÔÊ¾½á¹û/////
        If txtReplace.BackColor <> &HC0C0FF Then '²¶»ñÊý¾Ýö
            If InStr(strPattern, ">NEXT>") > 0 Then strPattern = Split(strPattern, ">NEXT>")(0)
            strCode = strCode & "    reg.Pattern = """ & strPattern & """" & vbCrLf & _
                                "    Set matchs = reg.Execute(strData)" & vbCrLf & _
                                "    For i = 0 To matchs.Count - 1" & vbCrLf
                                
            If isShowNumber Then
                If isShowFa Then 'ÊÇ·ñÏÔÊ¾¸¸Æ¥Åä
                    strCode = strCode & "        Debug.Print i + 1 & ""."" & matchs(i)" & vbCrLf
                End If
                If isShowSubs Then
                    strCode = strCode & "        For j = 0 To matchs(i).SubMatches.Count - 1" & vbCrLf & _
                                        "           Debug.Print ""("" & j + 1 & "")."" & matchs(i).SubMatches(j) & "" "";" & vbCrLf & _
                                        "        Next" & vbCrLf & _
                                        "        If matchs(i).SubMatches.Count > 0 Then Debug.Print" & vbCrLf
                End If
            Else
                If isShowFa Then 'ÊÇ·ñÏÔÊ¾¸¸Æ¥Åä
                    strCode = strCode & "        Debug.Print matchs(i)" & vbCrLf
                End If
                If isShowSubs Then
                    strCode = strCode & "        For j = 0 To matchs(i).SubMatches.Count - 1" & vbCrLf & _
                                        "           Debug.Print matchs(i).SubMatches(j) & "" "";" & vbCrLf & _
                                        "        Next" & vbCrLf & _
                                        "        If matchs(i).SubMatches.Count > 0 Then Debug.Print" & vbCrLf
                End If
            End If
            strCode = strCode & "    Next" & vbCrLf & _
                                "End Sub"
        Else 'Ìæ»»µÄÇé¿ö
            If InStr(txtReplace.Text, ">NEXT>") > 0 Then 'ÐèÒª¶à´ÎÌæ»»
                Dim v, v2, i%
                v = Split(strPattern, ">NEXT>")
                v2 = Split(strReplace, ">NEXT>")
                For i = 0 To UBound(v)
                    strReplacePart = """" & v2(i) & """"
                    strReplacePart = Replace(strReplacePart, "\r\n", """ & vbCrLf & """)
                    strReplacePart = Replace(strReplacePart, "\r", """ & vbCr & """)
                    strReplacePart = Replace(strReplacePart, "\n", """ & vbCr & """)
                    strReplacePart = Replace(strReplacePart, """"" & ", "")
                    strReplacePart = Replace(strReplacePart, " & """"", "")
                    strCode = strCode & "    reg.Pattern = """ & v(i) & """" & vbCrLf & _
                                        "    strData = reg.Replace(strData, " & strReplacePart & ")" & vbCrLf
                Next
                strCode = strCode & "    Debug.Print strData" & vbCrLf
            Else
                strReplacePart = """" & Replace(txtReplace.Text, """", """""") & """"
                strReplacePart = Replace(strReplacePart, "\r\n", """ & vbCrLf & """)
                strReplacePart = Replace(strReplacePart, "\r", """ & vbCr & """)
                strReplacePart = Replace(strReplacePart, "\n", """ & vbCr & """)
                strReplacePart = Replace(strReplacePart, """"" & ", "")
                strReplacePart = Replace(strReplacePart, " & """"", "")
                strCode = strCode & "    reg.Pattern = """ & strPattern & """" & vbCrLf & _
                                    "    Debug.Print reg.Replace(strData, " & strReplacePart & ")" & vbCrLf
            End If
            strCode = strCode & "End Sub" & vbCrLf
            
        End If
        
        '/////¹¹½¨ÏÂÔØÒ³Ãæ´úÂëº¯Êý/////
        If txtUrl.BackColor = &HC0C0FF Then
            If isFileExists(txtUrl.Text) Then
                strCode = strCode & vbCrLf & _
                        "Public Function fileStr(ByVal strFileName As String) As String" & vbCrLf & _
                        "    Dim fileHandl%" & vbCrLf & _
                        "    fileHandl = FreeFile" & vbCrLf & _
                        "    Open strFileName For Input As #fileHandl" & vbCrLf & _
                        "    fileStr = StrConv(InputB$(LOF(fileHandl), #fileHandl), vbUnicode)" & vbCrLf & _
                        "    Close #fileHandl" & vbCrLf & _
                        "End Function"
            Else
                If strDownMode = "Normal" Then
                    strCode = strCode & vbCrLf & _
                            "Private Function getHtmlStr(strUrl As String) As String" & vbCrLf & _
                            "    Dim XmlHttp As Object" & vbCrLf & _
                            "    Set XmlHttp = CreateObject(""" & strXmlHttp & """)" & vbCrLf & _
                            "    XmlHttp.Open ""GET"", strUrl, False" & vbCrLf & _
                            "    XmlHttp.SetRequestHeader ""If-Modified-Since"", ""0""" & vbCrLf & _
                            "    XmlHttp.send" & vbCrLf & _
                            "    getHtmlStr = StrConv(XmlHttp.ResponseBody, vbUnicode)" & vbCrLf & _
                            "    Set XmlHttp = Nothing" & vbCrLf & _
                            "End Function"
                Else 'UTF8,big5µÈ
                    strCode = strCode & vbCrLf & _
                        "Public Function getHtmlStr(strUrl As String) As String" & vbCrLf & _
                        "    Dim XmlHttp As Object" & vbCrLf & _
                        "    Set XmlHttp = CreateObject(""" & strXmlHttp & """)" & vbCrLf & _
                        "    XmlHttp.Open ""GET"", strUrl, False" & vbCrLf & _
                        "    XmlHttp.SetRequestHeader ""If-Modified-Since"", ""0""" & vbCrLf & _
                        "    XmlHttp.send" & vbCrLf & _
                        "    getHtmlStr = BytesToBstr(XmlHttp.ResponseBody, """ & strDownMode & """)" & vbCrLf & _
                        "    Set XmlHttp = Nothing" & vbCrLf & _
                        "End Function" & vbCrLf
                        
                    strCode = strCode & vbCrLf & _
                        "Private Function BytesToBstr(strBody, codeBase) As String" & vbCrLf & _
                        "    Dim objStream As Object" & vbCrLf & _
                        "    Set objStream = CreateObject(""Adodb.Stream"")" & vbCrLf & _
                        "    objStream.Type = 1" & vbCrLf & _
                        "    objStream.Mode = 3" & vbCrLf & _
                        "    objStream.Open" & vbCrLf & _
                        "    objStream.Write strBody" & vbCrLf & _
                        "    objStream.position = 0" & vbCrLf & _
                        "    objStream.Type = 2" & vbCrLf & _
                        "    objStream.Charset = codeBase" & vbCrLf & _
                        "    BytesToBstr = objStream.ReadText" & vbCrLf & _
                        "    objStream.Close" & vbCrLf & _
                        "    Set objStream = Nothing" & vbCrLf & _
                        "End Function"
                End If
            End If
        End If
        
        Clipboard.Clear
        Clipboard.SetText strCode
        intClickTimes = 1
        
        If intLanMode = eCH Then
            MsgBox "ÒÑ¾­³É¹¦½«´úÂëÉú³É²¢¸´ÖÆµ½¼ôÇÐ°å£¡", vbInformation, vCH(0)
        Else
            MsgBox "The code has been succesfully generated and has been copyied to your clipboard!", vbInformation, vEN(0)
        End If
    End If
    Label3.Caption = intClickTimes
End Sub

Private Sub Form_Resize()
On Error GoTo err1
    If Me.WindowState = 1 Then Exit Sub
    Call formResize
err1:
End Sub
'µ÷Õû´°Ìå
Private Sub formResize()
'On Error GoTo err1
    If txtSource.Tag = "Picture1_DblClick" Then '×î´ó»¯·µ»ØÊ±ºò±£Ö¤±ÈÀý²»±ä
        If Picture1.Tag <> 0.5 Then Picture1.Tag = txtSource.Top / Me.ScaleHeight
    End If
    If Picture1.Tag < 0 Or Picture1.Tag > 1 Then Picture1.Tag = 0.5
    
    Picture1.Top = Picture1.Tag * Me.ScaleHeight 'Ã¿´Î¶¼¸ù¾Ý±ÈÀý¼ÆËãÉÏÏÂÁ½¸ö¿òµÄ»®·Ö±È
    Picture1.Width = Me.ScaleWidth - Picture1.Left - 45
    txtUrl.Move txtUrl.Left, txtUrl.Top, Me.ScaleWidth - txtUrl.Left - cmdDownLoad.Width - 45 - 90 - 290, cboUrl.Height
    cboUrl.Move txtUrl.Left, txtUrl.Top, Me.ScaleWidth - txtUrl.Left - cmdDownLoad.Width - 45 - 90
    cmdDownLoad.Left = Me.ScaleWidth - cmdDownLoad.Width - 90
    txtSource.Move txtSource.Left, txtSource.Top, Picture1.Width
    If Picture1.Top - txtSource.Top <= 0 Then txtSource.Height = 0 Else txtSource.Height = Picture1.Top - txtSource.Top
    cboPattern.Move cboPattern.Left, Picture1.Top + 90, Picture1.Width
    cmdTest.Top = cboPattern.Top + cboPattern.Height + 45
    cmdSearch.Top = cmdTest.Top
    cmdReplace.Top = cmdTest.Top
    txtReplace.Move txtReplace.Left, cmdReplace.Top + 45, (Me.ScaleWidth - cmdReplace.Left - cmdReplace.Width) * 0.5 - 290, cboReplace.Height
    cboReplace.Move txtReplace.Left, txtReplace.Top, (Me.ScaleWidth - cmdReplace.Left - cmdReplace.Width) * 0.5
    cmdSet.Move cboReplace.Left + cboReplace.Width + 90, cmdTest.Top
    
    txtResult.Move txtResult.Left, cmdTest.Top + cmdTest.Height + 45, Picture1.Width
    Frame1.Move txtSource.Left + (txtSource.Width - Frame1.Width) / 2, txtSource.Top + (txtSource.Height - Frame1.Height) / 2
    Frame2.Move txtSource.Left + (txtSource.Width - Frame2.Width) / 2, txtSource.Top + (txtSource.Height - Frame2.Height) / 2
    If Me.ScaleHeight - cmdTest.Top - cmdTest.Height - 90 <= 0 Then txtResult.Height = 0 Else txtResult.Height = Me.ScaleHeight - cmdTest.Top - cmdTest.Height - 90
    Label6.Top = txtResult.Top
    Label2.Top = cboPattern.Top + 45
err1:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call saveToIniFile
    Call saveToDat
    End
End Sub

Private Sub cmdDownLoad_Click()
    cmdDownLoad.Enabled = False
    txtUrl.Text = Trim(txtUrl.Text)
    If txtUrl.Text <> "" Then
        If isFileExists(txtUrl.Text) Then
            txtSource.LoadFile txtUrl.Text
        Else
            txtSource.Text = GetHtmlByMicrosoftXMLHTTP(txtUrl.Text, strDownMode)
        End If
    End If
    cmdDownLoad.Enabled = True
    Call saveToDat
End Sub
'ÅÐ¶ÏÎÄ¼þÊÇ·ñ´æÔÚ
Private Function isFileExists(ByVal strFileName$) As Boolean
    On Error GoTo err1
    isFileExists = Dir(strFileName) <> ""
    Exit Function
err1:
End Function

Private Sub cmdReplace_Click()
    
    Call saveToIniFile
    Call saveToDat
    
    On Error GoTo err1
    Dim strTmp$, s$
    strTmp = Replace(txtReplace.Text, "\r\n", vbCrLf)
    strTmp = Replace(strTmp, "\r", vbCr)
    strTmp = Replace(strTmp, "\n", vbLf)
    strTmp = Replace(strTmp, "\t", vbTab)
    If InStr(strTmp, ">NEXT>") > 0 Then
        Dim v, v2, i%
        v = Split(cboPattern.Text, ">NEXT>")
        v2 = Split(strTmp, ">NEXT>")
        strTmp = txtSource.Text
        For i = 0 To UBound(v)
            reg.Pattern = v(i)
            strTmp = reg.Replace(strTmp, v2(i))
        Next
    Else
        reg.Pattern = cboPattern.Text
        strTmp = reg.Replace(txtSource.Text, strTmp)
    End If
    
    If Right(strTmp, 2) = vbCrLf Then strTmp = Left(strTmp, Len(strTmp) - 2)
    txtResult.Text = strTmp
    
    Exit Sub
err1:
    txtResult.Text = Err.Number & " " & Err.Description
End Sub

Private Sub cmdSearch_Click()
    Dim i%, j%, s$
    
    
    Call saveToIniFile
    Call saveToDat
    
    On Error GoTo err1
    reg.Pattern = cboPattern.Text
    If InStr(cboPattern.Text, ">NEXT>") > 0 Then reg.Pattern = Split(cboPattern.Text, ">NEXT>")(0)
    Set matchs = reg.Execute(txtSource.Text)
    If matchs.Count > 0 Then
        For i = 0 To matchs.Count - 1
            If isShowFa Then 'ÊÇ·ñÏÔÊ¾¸¸Æ¥Åä
                If isShowNumber Then
                    s = s & i + 1 & "." & matchs(i).Value
                Else
                    s = s & matchs(i).Value
                End If
                If isShowOfLine Then
                    s = s & vbTab & "ËùÔÚÐÐ: " & UBound(Split(Left(txtSource.Text, matchs(i).FirstIndex + 1), vbCrLf)) + 1 & vbCrLf
                Else
                    s = s & vbCrLf
                End If
            End If
            
            If isShowSubs Then
                If matchs(i).SubMatches.Count > 0 Then
                    For j = 0 To matchs(i).SubMatches.Count - 1
                        If isShowNumber Then
                            s = s & "(" & j + 1 & ")." & matchs(i).SubMatches(j) & " "
                        Else
                            s = s & matchs(i).SubMatches(j) & " "
                        End If
                        
                    Next
                    s = Left(s, Len(s) - 1)
                End If
                If Right(s, 2) <> vbCrLf Then s = s & vbCrLf
            End If
        Next
    End If
    txtResult.Text = s
    
    Exit Sub
err1:
    txtResult.Text = Err.Number & " " & Err.Description
End Sub

Private Sub cmdTest_Click()
    reg.Pattern = cboPattern.Text
    
    On Error GoTo err1
    Call saveToIniFile
    Call saveToDat
    txtResult.Text = reg.Test(txtSource.Text)
    
    Exit Sub
err1:
    txtResult.Text = Err.Number & " " & Err.Description
End Sub

Private Sub saveToDat()
    Dim dat1$, dat2$, dat3$, dat4$
    Dim s$
    If Dir(strAppPath & "History.dat") <> "" Then
        s = fileStr(strAppPath & "History.dat")
        dat1 = getRegMatchSub1(s, "//ÍøÖ·\r\n([\s\S]*?)\r\n#\r\n")
        dat2 = getRegMatchSub1(s, "//ÕýÔò±í´ïÊ½\r\n([\s\S]*?)\r\n#\r\n")
        dat3 = getRegMatchSub1(s, "//Ìæ»»ÄÚÈÝ\r\n([\s\S]*?)\r\n#\r\n")
        dat4 = getRegMatch1(s, "//×Ö·û´®\r\n[\s\S]*?\r\n#\r\n")
    End If
    
    dat1 = addLineToString(dat1, txtUrl.Text)
    dat2 = addLineToString(dat2, cboPattern.Text)
    dat3 = addLineToString(dat3, txtReplace.Text)
    dat4 = txtSource.Text
    
    writeToFile strAppPath & "History.dat", "//ÍøÖ·" & vbCrLf & dat1 & vbCrLf & "#" & vbCrLf & _
                                                "//ÕýÔò±í´ïÊ½" & vbCrLf & dat2 & vbCrLf & "#" & vbCrLf & _
                                                "//Ìæ»»ÄÚÈÝ" & vbCrLf & dat3 & vbCrLf & "#" & vbCrLf & _
                                                "//×Ö·û´®" & vbCrLf & dat4 & vbCrLf & "#"
    Call initToForm(False)
End Sub
Private Sub initToForm(Optional ByVal isAddSource As Boolean = True)
    Dim dat1$, dat2$, dat3$, dat4$
    Dim s$, v, i%
    If Dir(strAppPath & "History.dat") <> "" Then
        s = fileStr(strAppPath & "History.dat")
        dat1 = getRegMatchSub1(s, "//ÍøÖ·\r\n([\s\S]*?)\r\n#\r\n")
        If dat1 <> "" Then
            v = Split(dat1, vbCrLf)
            cboUrl.Clear
            For i = 0 To UBound(v)
                cboUrl.AddItem v(i)
            Next
            cboUrl.ListIndex = 0
            txtUrl.Text = cboUrl.Text
        End If
        
        dat2 = getRegMatchSub1(s, "//ÕýÔò±í´ïÊ½\r\n([\s\S]*?)\r\n#\r\n")
        If dat2 <> "" Then
            v = Split(dat2, vbCrLf)
            cboPattern.Clear
            For i = 0 To UBound(v)
                cboPattern.AddItem v(i)
            Next
            cboPattern.ListIndex = 0
        End If
        
        dat3 = getRegMatchSub1(s, "//Ìæ»»ÄÚÈÝ\r\n([\s\S]*?)\r\n#\r\n")
        '"//Ìæ»»ÄÚÈÝ" & vbCrLf & dat3 & vbCrLf & "#" & vbCrLf &
        If dat3 <> "" Then
            v = Split(dat3, vbCrLf)
            cboReplace.Clear
            For i = 0 To UBound(v)
                cboReplace.AddItem v(i)
            Next
            cboReplace.ListIndex = 0
            txtReplace.Text = cboReplace.Text
        End If
        
        If isAddSource Then
            dat4 = getRegMatchSub1(s, "//×Ö·û´®\r\n([\s\S]*?)\r\n#\r\n")
            txtSource.Text = dat4
        End If
    End If
End Sub

Private Function addLineToString(ByVal strData$, ByVal strNew$) As String
    Dim s$, v, i%
    v = Split(strData, vbCrLf)
    For i = 0 To UBound(v)
        If v(i) = strNew Then
            s = v(0)
            v(0) = v(i)
            v(i) = s
            addLineToString = Join(v, vbCrLf)
            Exit Function
        End If
    Next
    If strData <> "" Then
        addLineToString = strNew & vbCrLf & strData
    Else
        addLineToString = strNew & strData
    End If
End Function

Private Sub cmdSet_Click()
    frmSet.Show 1
End Sub
Public Sub setLanMain()
    vCH = Split(LAN_CH, ",")
    vEN = Split(LAN_EN, ",")
    Dim i%
    For i = 0 To UBound(vCH)
        vCH(i) = Trim(vCH(i))
        vEN(i) = Trim(vEN(i))
    Next
    Select Case intLanMode
        Case eCH
            Me.Caption = vCH(0)
            Label5.Caption = vCH(1)
            cmdDownLoad.Caption = vCH(2)
            Label1.Caption = vCH(3)
            Label2.Caption = vCH(4)
            Label6.Caption = vCH(5)
            cmdTest.Caption = vCH(6)
            cmdSearch.Caption = vCH(7)
            cmdReplace.Caption = vCH(8)
            cmdSet.Caption = vCH(9)
        Case eEN
            Me.Caption = vEN(0)
            Label5.Caption = vEN(1)
            cmdDownLoad.Caption = vEN(2)
            Label1.Caption = vEN(3)
            Label2.Caption = vEN(4)
            Label6.Caption = vEN(5)
            cmdTest.Caption = vEN(6)
            cmdSearch.Caption = vEN(7)
            cmdReplace.Caption = vEN(8)
            cmdSet.Caption = vEN(9)
    End Select
End Sub

Private Sub Frame1_Click()
    Frame1.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.BackColor = &H80FF&
    Timer3.Enabled = False
End Sub

Private Sub Label4_Click()
    Call Label7_Click
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer3.Enabled = False
End Sub

Private Sub Label7_Click()
    Dim i&, j&, s$, strKeyWord$
    Frame2.Visible = False
    
    strKeyWord = LCase(txtKeyword.Text)
    s = LCase(txtSource.Text)
    txtSource.SelStart = 0
    txtSource.SelLength = Len(s)
    txtSource.SelColor = vbBlack
    For i = 1 To Len(s)
        DoEvents
        If Mid(s, i, Len(strKeyWord)) = strKeyWord Then
            txtSource.SelStart = i - 1
            txtSource.SelLength = Len(strKeyWord)
            txtSource.SelColor = vbRed
            txtSource.SelBold = True
        End If
    Next
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.BackColor = &H8080FF
    Timer3.Enabled = False
End Sub

Private Sub Picture1_DblClick()
'    Picture1.Tag = IIf(Picture1.Tag = 0.07226107, 0.5, 0.07226107)
    Picture1.Tag = IIf(Picture1.Top = txtSource.Top, 0.5, txtSource.Top / Me.ScaleHeight)
    Call formResize
    txtSource.Tag = "Picture1_DblClick"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMove = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isMove = True Then
        txtSource.Tag = ""
        Picture1.Top = Picture1.Top + Y
        Picture1.Tag = Picture1.Top / Me.ScaleHeight
    End If
    Call formResize
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMove = False
End Sub

Private Sub Timer1_Timer()
    intClickTimes = 1
    Timer1.Enabled = False
    Label3.Caption = intClickTimes
'    MsgBox Label3
End Sub

Private Sub Timer2_Timer()
    Frame1.Visible = False
    Timer2.Enabled = False
    Call txtSource_Change
End Sub

Private Sub Timer3_Timer()
    Frame2.Visible = False
    Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
    If cboPattern.Text <> "" And cboPattern.SelStart <> 0 Then cboPattern.Tag = cboPattern.SelStart
End Sub

Private Sub txtKeyword_Change()
    Timer3.Enabled = False
End Sub

Private Sub txtKeyword_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call Label7_Click
End Sub

Private Sub txtKeyword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.BackColor = &H80FF&
    Timer3.Enabled = False
End Sub

'"CTRL"  ->  shift=2;  "ENTER" ->  keycode=13;  "DEL"   ->  keycode=46
Private Sub cboPattern_KeyDown(KeyCode As Integer, Shift As Integer)
'    MsgBox KeyCode & " " & Shift: Exit Sub
    
    If KeyCode = 13 Then '°´»Ø³µ¼ü
        Dim i%
        i = cboPattern.SelStart
        If Shift = 2 Then '°´CTRL¼ü
            cboPattern.Text = Replace(cboPattern.Text, """""", """")
        Else
            cboPattern.Text = Replace(cboPattern.Text, """", """""")
        End If
        cboPattern.SelStart = i
    End If
End Sub

Private Sub txtReplace_Click()
    txtReplace.BackColor = &H80000005
End Sub

Private Sub txtReplace_DblClick()
    txtReplace.BackColor = &HC0C0FF
End Sub

Private Sub cboReplace_Click()
    txtReplace.Text = cboReplace.Text
End Sub

Private Sub txtSource_Change()
    If chkDelHuiche.Value Then txtSource.Text = Replace(txtSource.Text, vbCrLf, "")
    If chkDelTab.Value Then txtSource.Text = Replace(txtSource.Text, vbTab, "")
End Sub

Private Sub txtSource_DblClick()
    Frame1.Visible = True
    Timer2.Enabled = True
    Frame1.Move txtSource.Left + (txtSource.Width - Frame1.Width) / 2, txtSource.Top + (txtSource.Height - Frame1.Height) / 2
End Sub

Private Sub txtSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Frame1.Visible Then
        Timer2.Enabled = True
    End If
    If Frame2.Visible Then
        Timer3.Enabled = True
    End If
    Label4.BackColor = &H80FF&
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Enabled = False
End Sub

Private Sub cmdDownLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame1.Visible = False
End Sub

Private Sub txtUrl_LostFocus()
    txtUrl.SelStart = 0
End Sub

Private Sub txtUrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame1.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame1.Visible = False
End Sub

Private Sub txtSource_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err1
    Dim strDragFile As String
    If Data.GetFormat(1) Then Exit Sub
    strDragFile = Data.Files.Item(Data.Files.Count)
    txtUrl.Text = strDragFile
    txtSource.LoadFile strDragFile
err1:
    Exit Sub
End Sub

Private Sub txtUrl_Click()
    txtUrl.BackColor = &H80000005
End Sub
Private Sub txtUrl_DblClick()
    txtUrl.BackColor = &HC0C0FF
End Sub

Private Sub txtUrl_GotFocus()
    txtUrl.SelStart = 0
    txtUrl.SelLength = Len(txtUrl.Text)
End Sub

Private Sub txtUrl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call cmdDownLoad_Click
End Sub

Private Sub txtUrl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err1
    Dim strDragFile As String
    If Data.GetFormat(1) Then Exit Sub
    strDragFile = Data.Files.Item(Data.Files.Count)
    txtUrl.Text = strDragFile
err1:
    Exit Sub
End Sub

Private Sub cboUrl_Click()
    txtUrl.Text = cboUrl.Text
End Sub
