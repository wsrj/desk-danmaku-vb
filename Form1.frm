VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "弹幕神器"
   ClientHeight    =   5505
   ClientLeft      =   5910
   ClientTop       =   3150
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   240
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "发送(&E)"
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   4800
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "外观"
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6615
      Begin VB.TextBox txtWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2880
         Width           =   1335
      End
      Begin ComctlLib.Slider sldWidth 
         Height          =   615
         Left            =   3360
         TabIndex        =   11
         Top             =   2760
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   50
         Min             =   100
         Max             =   1000
         SelStart        =   100
         TickStyle       =   2
         TickFrequency   =   100
         Value           =   100
      End
      Begin VB.TextBox txtColor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton btnColorPicker 
         Caption         =   "选择颜色(&I)…"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtFont 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton btnFontPicker 
         Caption         =   "选择字体(&P)…"
         Height          =   495
         Left            =   4560
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtDanmaku 
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "哔哩哔哩( ゜- ゜)つロ干杯~"
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "弹幕最大宽度"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "弹幕颜色"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "弹幕字体"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "弹幕文字"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "发送弹幕"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00993300&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuContact 
         Caption         =   "联系作者(&C)"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于弹幕神器(&A)…"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DanmakuColor '颜色变量
Dim DanmakuWidth '宽度变量
Dim DanmakuFontName '字体变量
Dim DanmakuFontSize '字号变量
Dim DanmakuFontText '显示字体文本

Private Sub btnSend_Click()
Form2.Show
Form2.Move Screen.Width, 0
Timer1.Enabled = True
End Sub

Private Sub Form_Initialize()
sldWidth.Max = Screen.Width
sldWidth.TickFrequency = sldWidth.Max / 10
sldWidth.Value = 100
Form2.Width = 100
txtWidth.Text = 100 & " px"
Timer1.Enabled = False
End Sub

Private Sub btnColorPicker_Click()
CommonDialog1.DialogTitle = "选择颜色"
CommonDialog1.ShowColor '弹出选择颜色
DanmakuColor = CommonDialog1.Color
txtColor.Text = DanmakuColor '在标签中显示颜色
Form2.Label1.ForeColor = DanmakuColor
End Sub

Private Sub btnFontPicker_Click()
CommonDialog1.DialogTitle = "选择字体"
CommonDialog1.ShowFont '弹出选择字体
DanmakuFontName = CommonDialog1.FontName
DanmakuFontSize = CommonDialog1.FontSize
DanmakuFontText = "字体：" & CommonDialog1.FontName & "，字号：" & CommonDialog1.FontSize
txtFont.Text = DanmakuFontText '在标签中显示字体
Form2.Label1.FontName = DanmakuFontName
Form2.Label1.FontSize = DanmakuFontSize
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAbout_Click()
ShellAboutA Me.hWnd, "关于弹幕神器", "这是一个能在桌面上发弹幕的程序"
End Sub

Private Sub mnuContact_Click()
MsgBox "请访问 https://space.bilibili.com/3493134929496963", vbInformation + vbOKOnly, "联系作者"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOpen_Click()
MsgBox "这个功能还没有开发", vbInformation + vbkonly, "提示"
End Sub

Private Sub mnuSave_Click()
CommonDialog1.DialogTitle = "保存配置"
CommonDialog1.DefaultExt = "*.ini"
CommonDialog1.ShowSave
file = CommonDialog1.FileName
If DanmakuFontSize <> "" And DanmakuFontText <> "" Then
    ResultWriteFont = WritePrivateProfileStringA("danmaku", "font", DanmakuFontName, file)
    ResultWriteSize = WritePrivateProfileStringA("danmaku", "size", DanmakuFontSize, file)
    ResultWriteWidth = WritePrivateProfileStringA("danmaku", "width", DanmakuWidth, file)
    If ResultWriteFont <> 0 And ResultWriteSize <> 0 And ResultWriteWidth <> 0 Then
        MsgBox "保存配置成功！", vbInformation + vbOKOnly, "成功"
    Else
        MsgBox "保存配置失败 (ResultWriteFont=" & ResultWriteFont & ", ResultWriteSize=" & ResultWriteSize & ")", vbCritical, "失败"
    End If
Else
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub sldWidth_Change()
DanmakuWidth = sldWidth.Value
txtWidth.Text = DanmakuWidth & " px"
Form2.Width = DanmakuWidth
Form2.Label1.Width = DanmakuWidth
End Sub

Sub Timer1_Timer()
Form2.Move Form2.Left - 1000, 0
If Form2.Left < 0 - Form2.Width Then
    Form1.Timer1.Enabled = False
    Unload Form2
End If
End Sub

Private Sub txtDanmaku_Change()
Form2.Label1.Caption = txtDanmaku.Text
End Sub

