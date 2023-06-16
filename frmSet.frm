VERSION 5.00
Begin VB.Form frmSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmSet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4080
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   5
         ItemData        =   "frmSet.frx":000C
         Left            =   1200
         List            =   "frmSet.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   8
         ItemData        =   "frmSet.frx":0027
         Left            =   1200
         List            =   "frmSet.frx":0031
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   3
         ItemData        =   "frmSet.frx":0052
         Left            =   1200
         List            =   "frmSet.frx":005C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   7
         ItemData        =   "frmSet.frx":006D
         Left            =   1200
         List            =   "frmSet.frx":007A
         TabIndex        =   7
         Text            =   "cboItem"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   6
         ItemData        =   "frmSet.frx":0093
         Left            =   1200
         List            =   "frmSet.frx":009D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   4
         ItemData        =   "frmSet.frx":00AE
         Left            =   1200
         List            =   "frmSet.frx":00B8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   2
         ItemData        =   "frmSet.frx":00C9
         Left            =   1200
         List            =   "frmSet.frx":00D3
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   0
         ItemData        =   "frmSet.frx":00E4
         Left            =   1200
         List            =   "frmSet.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   2295
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Index           =   1
         ItemData        =   "frmSet.frx":00FF
         Left            =   1200
         List            =   "frmSet.frx":0109
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "显示编号:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   345
         TabIndex        =   21
         Top             =   1995
         Width           =   765
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "软件语言:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   345
         TabIndex        =   20
         Top             =   3045
         Width           =   765
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "显示父匹配:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   165
         TabIndex        =   19
         Top             =   1275
         Width           =   945
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "下载模式:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   345
         TabIndex        =   18
         Top             =   2685
         Width           =   765
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "显示所在行:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   165
         TabIndex        =   17
         Top             =   2325
         Width           =   945
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "显示子匹配:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   165
         TabIndex        =   16
         Top             =   1635
         Width           =   945
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MultiLine:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   285
         TabIndex        =   15
         Top             =   885
         Width           =   825
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IgnoreCase:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   165
         Width           =   1050
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Global:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   495
         TabIndex        =   13
         Top             =   525
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "点击这里了解更多"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      MouseIcon       =   "frmSet.frx":011A
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4440
      Width           =   1680
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const LAN_CH = "显示父匹配:,显示子匹配:,显示编号:  ,显示所在行:,下载模式:,软件语言:,设置,确定,取消   ,点击这里了解更多,正则测试工具    ,网　址:,下载代码↓ ,字符串:,表达式: ,结　果:,完全匹配,检索结果    ,全部替换    ,系统设置"
Private Const LAN_EN = "ShowFather:,Showsub:   ,ShowNum:   ,ShowOfLine:,DownMode:,Language:,Set ,&OK ,&Cancel,About me        ,RegExp Test Tool,URL:   ,&Download↓,String:,Pattern:,Result:,&Test   ,&Show Result,&Replace All,&Set"
Private vCH, vEN
Dim isLoad As Boolean

Private Sub cboItem_Click(Index As Integer)
    If Index = 8 Then
        If isLoad Then
            isLoad = False
        Else
            Call flushLanguage
            Call setLan
        End If
    End If
End Sub

Sub flushLanguage()
    reg.IgnoreCase = cboItem(0).List(cboItem(0).ListIndex)
    reg.Global = cboItem(1).List(cboItem(1).ListIndex)
    reg.MultiLine = cboItem(2).List(cboItem(2).ListIndex)
    isShowFa = cboItem(3).List(cboItem(3).ListIndex)
    isShowSubs = cboItem(4).List(cboItem(4).ListIndex)
    isShowNumber = cboItem(5).List(cboItem(5).ListIndex)
    isShowOfLine = cboItem(6).List(cboItem(6).ListIndex)
    strDownMode = cboItem(7).Text
    intLanMode = cboItem(8).ListIndex
    
    Call saveToIniFile
    frmMain.setLanMain
End Sub

Private Sub Command1_Click()
    Call flushLanguage
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    isLoad = True
    vCH = Split(LAN_CH, ",")
    vEN = Split(LAN_EN, ",")
    
    Call setLan
    
'    On Error GoTo err1
    Me.Icon = frmMain.Icon
    cboItem(0).ListIndex = IIf(reg.IgnoreCase, 0, 1)
    cboItem(1).ListIndex = IIf(reg.Global, 0, 1)
    cboItem(2).ListIndex = IIf(reg.MultiLine, 0, 1)
    cboItem(3).ListIndex = IIf(isShowFa, 0, 1)
    cboItem(4).ListIndex = IIf(isShowSubs, 0, 1)
    cboItem(5).ListIndex = IIf(isShowNumber, 0, 1)
    cboItem(6).ListIndex = IIf(isShowOfLine, 0, 1)
    cboItem(7).Text = strDownMode
    cboItem(8).ListIndex = intLanMode

    Exit Sub
err1:
    Call initFromApp
End Sub

Private Sub setLan()
    Dim i%
    Select Case intLanMode
        Case eCH
            For i = 3 To 8
                lblItem(i).Caption = Trim(vCH(i - 3))
            Next
            Me.Caption = Trim(vCH(6))
            Command1.Caption = Trim(vCH(7))
            Command2.Caption = Trim(vCH(8))
            Label4.Caption = Trim(vCH(9))
        Case eEN
            For i = 3 To 8
                lblItem(i).Caption = Trim(vEN(i - 3))
            Next
            Me.Caption = Trim(vEN(6))
            Command1.Caption = Trim(vEN(7))
            Command2.Caption = Trim(vEN(8))
            Label4.Caption = Trim(vEN(9))
    End Select
End Sub

Private Sub Label4_Click()
    ShellExecute hwnd, "open", strDecode16Hex("%6874%747" & Chr(48) & Chr(37) & Chr(51) & Chr(65) & "2F2F77%77" & "%77%2E7" & "379%6D%65" & "%6E7" & "461%6C2" & "E%636F6" & "D"), "", "", 1
End Sub
'解密
Public Function strDecode16Hex(strSource$)
    Dim i As Long
    Dim bytSource() As Byte
    strSource = Replace(strSource, "%", "")
    ReDim bytSource(Len(strSource) / 2 - 1)
    For i = 0 To Len(strSource) / 2 - 1
        bytSource(i) = "&H" & (Mid(strSource, i * 2 + 1, 2))
    Next
    strDecode16Hex = StrConv(bytSource, vbUnicode)
End Function

