VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   240
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "发送(&E)"
      Height          =   495
      Left            =   5280
      TabIndex        =   12
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
   Begin VB.Frame fraAppearance 
      Caption         =   "外观"
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6615
      Begin MSComctlLib.Slider sldWidth 
         Height          =   615
         Left            =   3840
         TabIndex        =   13
         ToolTipText     =   "长方形代表屏幕宽度，颜色填充的部分为弹幕所占宽度"
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         _Version        =   393216
         Max             =   1000
         SelectRange     =   -1  'True
         TickStyle       =   2
         Value           =   100
      End
      Begin VB.TextBox txtWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "弹幕最大能达到多宽，如果超出该值就会换行"
         Top             =   2880
         Width           =   1335
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
         ToolTipText     =   "你选择的颜色"
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
         ToolTipText     =   "你选择的字体与字号"
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
         ToolTipText     =   "输入你要发的弹幕内容"
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
      Caption         =   "桌面弹幕"
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
         Enabled         =   0   'False
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
   Begin VB.Menu mnuTool 
      Caption         =   "工具(&T)"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuOption 
         Caption         =   "选项(&O)…"
         Enabled         =   0   'False
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
Attribute VB_Name = "frmMain"
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
If DanmakuColor <> "" And DanmakuFontName <> "" And DanmakuFontSize <> "" Then
    frmContainer.Label1.ForeColor = DanmakuColor
    frmContainer.Label1.FontName = DanmakuFontName
    frmContainer.Label1.FontSize = DanmakuFontSize
    frmContainer.Show
    frmContainer.Move Screen.Width, 0
    Timer1.Enabled = True
Else
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub Form_Initialize()
sldWidth.Max = Screen.Width
sldWidth.TickFrequency = sldWidth.Max / 10
sldWidth.Value = 100
frmContainer.Width = 100
txtWidth.Text = 100
Timer1.Enabled = False
End Sub

Private Sub btnColorPicker_Click()
CommonDialog1.DialogTitle = "选择颜色"
CommonDialog1.ShowColor '弹出选择颜色
DanmakuColor = CommonDialog1.Color
txtColor.Text = DanmakuColor '在标签中显示颜色
frmContainer.Label1.ForeColor = DanmakuColor
End Sub

Private Sub btnFontPicker_Click()
CommonDialog1.DialogTitle = "选择字体"
CommonDialog1.ShowFont '弹出选择字体
DanmakuFontName = CommonDialog1.FontName
DanmakuFontSize = CommonDialog1.FontSize
If DanmakuFontName <> "" And DanmakuFontSize <> "" Then
    DanmakuFontText = "字体：" & CommonDialog1.FontName & "，字号：" & CommonDialog1.FontSize
    txtFont.Text = DanmakuFontText '在标签中显示字体
    frmContainer.Label1.FontName = DanmakuFontName
    frmContainer.Label1.FontSize = DanmakuFontSize
Else
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAbout_Click()
ShellAboutA Me.hWnd, _
    "关于弹幕神器#弹幕神器", _
    "这是一个能在桌面上发弹幕的程序！" & vbCrLf & _
    "联系作者：" & vbCrLf & _
    "https://space.bilibili.com/3493134929496963" & vbCrLf & _
    "software.wsrj@outlook.com" & vbCrLf & _
    "https://github.com/wsrj", _
    Me.Icon.Handle
End Sub

Private Sub mnuContact_Click()
MsgBox "请访问 https://space.bilibili.com/3493134929496963", vbInformation + vbOKOnly, "联系作者"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOpen_Click()
CommonDialog1.ShowOpen
file = CommonDialog1.FileName
buffer = Space$(32768)
DanmakuFontName = GetPrivateProfileStringA("danmaku", "font", vbNullString, buffer, 32768, file)
DanmakuFontSize = GetPrivateProfileStringA("danmaku", "size", vbNullString, buffer, 32768, file)
DanmakuWidth = GetPrivateProfileStringA("danmaku", "width", vbNullString, buffer, 32768, file)
DanmakuColor = GetPrivateProfileStringA("danmaku", "color", vbNullString, buffer, 32768, file)
frmContainer.Width = DanmakuWidth
frmContainer.Label1.Font = DanmakuFontName
frmContainer.Label1.FontSize = DanmakuFontSize
frmContainer.Label1.ForeColor = DanmakuColor
If DanmakuWidth < 100 Then
    MsgBox "弹幕宽度不能低于 100。", vbInformation + vbOKOnly, "提示"
    DanmakuWidth = 100
End If
txtWidth.Text = DanmakuWidth
txtFont.Text = "字体：" & DanmakuFontName & "，字号：" & DanmakuFontSize
txtColor.Text = DanmakuColor
End Sub

Private Sub mnuOption_Click()
frmOption.Show vbModal, frmMain
End Sub

Private Sub mnuSave_Click()
CommonDialog1.DialogTitle = "保存配置"
CommonDialog1.DefaultExt = "*.ini"
CommonDialog1.ShowSave
file = CommonDialog1.FileName
' 如果弹幕属性值都不为空
If DanmakuFontSize <> "" And DanmakuFontText <> "" And DanmakuColor <> "" Then
    ' 写入四个属性值
    ResultWriteFont = WritePrivateProfileStringA("danmaku", "font", DanmakuFontName, file)
    ResultWriteSize = WritePrivateProfileStringA("danmaku", "size", DanmakuFontSize, file)
    ResultWriteWidth = WritePrivateProfileStringA("danmaku", "width", DanmakuWidth, file)
    ResultWriteColor = WritePrivateProfileStringA("danmaku", "color", DanmakuColor, file)
    ' 如果返回值都不为 0 则成功
    If ResultWriteFont <> 0 And ResultWriteSize <> 0 And ResultWriteWidth <> 0 And ResultWriteColor <> 0 Then
        MsgBox "保存配置成功！", vbInformation + vbOKOnly, "成功"
    Else
        MsgBox "保存配置失败 (ResultWriteFont=" & ResultWriteFont & ", ResultWriteSize=" & ResultWriteSize & "ResultWriteColor=" & ResultWriteColor & ")", vbCritical, "失败"
    End If
Else
    ' 如果有空的则弹出提示
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub sldWidth_Change()
' 如果值小于 100 则自动更改为 100
If sldWidth.Value < 100 Then
    sldWidth.Value = 100
End If
sldWidth.SelStart = 0
' 设置填充部分长度
sldWidth.SelLength = sldWidth.Value
DanmakuWidth = sldWidth.Value
txtWidth.Text = DanmakuWidth
frmContainer.Width = DanmakuWidth
frmContainer.Label1.Width = DanmakuWidth
End Sub

Sub Timer1_Timer()
fraAppearance.Enabled = False
sldWidth.Enabled = False
btnSend.Enabled = False
' 每个时钟周期向左移动 50 像素
frmContainer.Move frmContainer.Left - 50, Screen.Height / 20
If frmContainer.Left < 0 - frmContainer.Width Then
    frmMain.Timer1.Enabled = False
    frmContainer.Left = Screen.Width ' 移动回初始位置
    fraAppearance.Enabled = True
    sldWidth.Enabled = True
    btnSend.Enabled = True
End If
End Sub

Private Sub txtDanmaku_Change()
frmContainer.Label1.Caption = txtDanmaku.Text
End Sub
