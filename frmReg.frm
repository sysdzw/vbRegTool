VERSION 5.00
Begin VB.Form frmReg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "正则库存(双击需要的项目自动填入表达式处)"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   15390
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkStay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "正则插入后保留窗口"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4080
      TabIndex        =   3
      Top             =   97
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "覆盖当前表达式"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "插入到光标处"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   8640
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   14415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "扩充正则表达式库"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6915
      TabIndex        =   4
      Top             =   97
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6915
      Top             =   60
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   9240
      Picture         =   "frmReg.frx":000C
      ToolTipText     =   "点击重新载入正则库"
      Top             =   112
      Width           =   210
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkStay_Click()
    isStayRegForm = (chkStay.Value = 1)
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    
    If isInsertReg Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    chkStay.Value = IIf(isStayRegForm, 1, 0)
    
    Call loadRegLib
End Sub

Private Sub Option1_Click()
    isInsertReg = Option1.Value
End Sub

Private Sub Option2_Click()
    isInsertReg = Not Option2.Value
End Sub

Private Function loadRegLib()
    Dim s$, v, i&
    s = fileStr(strAppPath & "reg.txt")
    v = Split(s, vbCrLf)
    List1.Clear
    For i = 0 To UBound(v)
        If v(i) <> "" And Left(v(i), 2) <> "//" Then List1.AddItem v(i)
    Next
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFFFFFF
    List1.ForeColor = &H404040
    Shape1.FillColor = &HC000&
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFFFFFF
    List1.ForeColor = &H404040
    Shape1.FillColor = &HC000&
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.ForeColor = &H8635AD
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &H40C0&
    Shape1.FillColor = &HC0FFC0
End Sub

Private Sub Form_Resize()
    List1.Move 45, List1.Top, Me.ScaleWidth - 90, Me.ScaleHeight - List1.Top - 45
End Sub

Private Sub Image1_Click()
    Call loadRegLib
End Sub

Private Sub Label1_Click()
    Shell "notepad.exe """ & strAppPath & "reg.txt" & """", 1
End Sub

Private Sub List1_DblClick()
    Dim s$, v
    s = List1.List(List1.ListIndex)
    v = Split(s, vbTab)
    s = v(1)
    If InStr(s, "~~") > 0 Then
        v = Split(s, "~~")
        frmMain.cboPattern.Text = v(0)
        frmMain.txtReplace.Text = v(1)
    Else
        If Option1.Value Then
            frmMain.cboPattern.SelStart = frmMain.cboPattern.Tag
            frmMain.cboPattern.SelText = s
        Else
            frmMain.cboPattern.Text = s
        End If
    End If
    If chkStay.Value = 0 Then Unload Me
End Sub

