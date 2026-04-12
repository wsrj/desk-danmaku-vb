VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "弹幕神器"
   ClientHeight    =   7335
   ClientLeft      =   5910
   ClientTop       =   3150
   ClientWidth     =   7080
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
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   Begin VB.Frame fraWidth 
      Caption         =   "弹幕最大宽度（单位：二十分之一点）"
      Height          =   2535
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   6615
      Begin VB.TextBox txtWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "弹幕最大能达到多宽，如果超出该值就会换行"
         Top             =   480
         Width           =   6015
      End
      Begin MSComctlLib.Slider sldWidth 
         Height          =   615
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "长方形代表屏幕宽度，颜色填充的部分为弹幕所占宽度"
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1085
         _Version        =   393216
         Max             =   1000
         SelectRange     =   -1  'True
         TickStyle       =   2
         Value           =   100
      End
      Begin VB.Label Label5 
         Caption         =   "注：长方形代表屏幕宽度，颜色填充的部分为弹幕所占宽度。"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   5895
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   6960
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1005
            MinWidth        =   176
            Text            =   "CapsLk"
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock 键状态"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   1032
            MinWidth        =   176
            Text            =   "NumLk"
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock 键状态"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   767
            MinWidth        =   176
            Text            =   "Insert"
            TextSave        =   "INS"
            Object.ToolTipText     =   "Insert 键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "2026/4/12"
            Object.ToolTipText     =   "日期"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "19:22"
            Object.ToolTipText     =   "时间"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   240
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "发送(&E)"
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   6240
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
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6615
      Begin VB.TextBox txtColor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "你选择的颜色"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton btnColorPicker 
         Caption         =   "选择颜色(&I)…"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtFont 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "弹幕颜色"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "弹幕字体"
         ForeColor       =   &H00000000&
         Height          =   255
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label lblEgg 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开(&O)"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSend_Click()
If DanmakuColor <> "" And DanmakuFontName <> "" And DanmakuFontSize <> "" Then
    frmContainer.Label1.ForeColor = DanmakuColor
    frmContainer.Label1.FontName = DanmakuFontName
    frmContainer.Label1.FontSize = DanmakuFontSize
    frmContainer.Show
    frmContainer.Move Screen.Width, 0
    Timer1.Enabled = True
    ConOut "发送弹幕：Text=" & txtDanmaku.text & _
           ", Color=" & DanmakuColor & _
           ", Font=" & DanmakuFontName & "," & DanmakuFontSize
Else
    ConOut "错误：缺少必填项"
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub Form_Initialize()
ShowEgg = 0
sldWidth.Max = Screen.Width
sldWidth.TickFrequency = sldWidth.Max / 10
sldWidth.Value = 100
frmContainer.Width = 100
txtWidth.text = "100 twip"
Timer1.Enabled = False
ConOut "提示："
ConOut "当前显示器 缇(二十分之一点)/像素 值（横向）：" & Screen.TwipsPerPixelX
ConOut "当前显示器 缇/像素 值（纵向）：              " & Screen.TwipsPerPixelY
ConOut "", False
End Sub

Private Sub btnColorPicker_Click()
CommonDialog1.DialogTitle = "选择颜色"
CommonDialog1.ShowColor '弹出选择颜色
ConOut "选择颜色"
DanmakuColor = CommonDialog1.Color
txtColor.text = DanmakuColor '在标签中显示颜色
frmContainer.Label1.ForeColor = DanmakuColor
ConOut "选择的颜色：" & DanmakuColor
End Sub

Private Sub btnFontPicker_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "选择字体"
CommonDialog1.ShowFont '弹出选择字体
ConOut "选择字体"
DanmakuFontName = CommonDialog1.FontName
DanmakuFontSize = CommonDialog1.FontSize
If DanmakuFontName <> " " And DanmakuFontName <> "" And DanmakuFontSize <> "" Then
    DanmakuFontText = "字体：" & CommonDialog1.FontName & "，字号：" & CommonDialog1.FontSize
    txtFont.text = DanmakuFontText '在标签中显示字体
    frmContainer.Label1.FontName = DanmakuFontName
    frmContainer.Label1.FontSize = DanmakuFontSize
    ConOut "选择了字体：" & DanmakuFontName & "," & DanmakuFontSize
Else
    ConOut "错误：缺少必填项"
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
ConOut "结束"
If ShowConsole Then FreeConsole
End If
End
End Sub

Private Sub lblEgg_Click()
' 调用彩蛋功能
Egg ShowEgg
End Sub

Private Sub lblEgg_DblClick()
' 调用彩蛋功能
Egg ShowEgg
End Sub

Private Sub mnuAbout_Click()
ConOut "打开“关于”对话框"
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
ConOut "单击“联系作者”"
MsgBox "请访问 https://space.bilibili.com/3493134929496963", vbInformation + vbOKOnly, "联系作者"
End Sub

Private Sub mnuExit_Click()
ConOut "单击“退出”"
End
End Sub

'下面这坨使就让它封存吧
'Private Sub mnuOpen_Click()
'CommonDialog1.ShowOpen
'file = CommonDialog1.FileName
'buffer = Space$(32768)
'DanmakuFontName = GetPrivateProfileStringA("danmaku", "font", vbNullString, buffer, 32768, file)
'DanmakuFontSize = GetPrivateProfileStringA("danmaku", "size", vbNullString, buffer, 32768, file)
'DanmakuWidth = GetPrivateProfileStringA("danmaku", "width", vbNullString, buffer, 32768, file)
'DanmakuColor = GetPrivateProfileStringA("danmaku", "color", vbNullString, buffer, 32768, file)
'frmContainer.Width = DanmakuWidth
'frmContainer.Label1.Font = DanmakuFontName
'frmContainer.Label1.FontSize = DanmakuFontSize
'frmContainer.Label1.ForeColor = DanmakuColor
'If DanmakuWidth < 100 Then
'    MsgBox "弹幕宽度不能低于 100。", vbInformation + vbOKOnly, "提示"
'    DanmakuWidth = 100
'End If
'txtWidth.Text = DanmakuWidth
'txtFont.Text = "字体：" & DanmakuFontName & "，字号：" & DanmakuFontSize
'txtColor.Text = DanmakuColor
'End Sub
'现在这坨使被下面这些接替了

Private Sub mnuOpen_Click()
Dim file As String
CommonDialog1.DialogTitle = "打开配置(XML)"
CommonDialog1.DefaultExt = "*.xml"
CommonDialog1.ShowOpen
file = CommonDialog1.FileName
If file = "" Or file = " " Then Exit Sub
txtDanmaku.text = ReadXML(file, "/config/text")
txtColor.text = ReadXML(file, "/config/color")
DanmakuFontSize = ReadXML(file, "/config/font/size")
DanmakuFontName = ReadXML(file, "/config/font/name")
DanmakuFontText = "字体：" & CommonDialog1.FontName & "，字号：" & CommonDialog1.FontSize
txtFont.text = DanmakuFontText '在标签中显示字体
frmContainer.Label1.FontName = DanmakuFontName
frmContainer.Label1.FontSize = DanmakuFontSize
ConOut "读取配置文件：" & file
End Sub

Private Sub mnuSave_Click()
Dim xmlDoc As msxml2.DOMDocument60
Dim nodeConfig, nodeText, nodeColor, nodeFont, nodeFontSize, nodeFontName
Dim file As String
Set xmlDoc = New msxml2.DOMDocument60
CommonDialog1.DialogTitle = "保存配置(XML)"
CommonDialog1.DefaultExt = "*.xml"
CommonDialog1.ShowSave
file = CommonDialog1.FileName
If file = "" Or file = " " Then Exit Sub
' 如果弹幕属性值都不为空
If DanmakuFontSize <> "" And DanmakuFontText <> "" And DanmakuColor <> "" Then
    ' 写入四个属性值
    Set nodeConfig = xmlDoc.createElement("config")
    xmlDoc.appendChild nodeConfig
    Set nodeText = xmlDoc.createElement("text")
    nodeText.text = txtDanmaku.text
    nodeConfig.appendChild nodeText
    Set nodeColor = xmlDoc.createElement("color")
    nodeColor.text = DanmakuColor
    nodeConfig.appendChild nodeColor
    Set nodeFont = xmlDoc.createElement("font")
    nodeConfig.appendChild nodeFont
    Set nodeFontSize = xmlDoc.createElement("size")
    nodeFontSize.text = DanmakuFontSize
    nodeFont.appendChild nodeFontSize
    Set nodeFontName = xmlDoc.createElement("name")
    nodeFontName.text = DanmakuFontName
    nodeFont.appendChild nodeFontName
    xmlDoc.save file
    ConOut "保存配置文件：" & file
    MsgBox "已保存配置文件", vbOKOnly + vbInformation, "提示"
Else
    ' 如果有空的则弹出提示
    ConOut "错误：缺少必填项"
    MsgBox "你还没有输入完", vbInformation + vbOKOnly, "提示"
End If
End Sub

Private Sub sldWidth_Change()
' 如果值小于 100 则自动更改为 100
If sldWidth.Value < 100 Then
    ConOut "错误：弹幕宽度不能小于 100 缇"
    sldWidth.Value = 100
End If
sldWidth.SelStart = 0
' 设置填充部分长度
sldWidth.SelLength = sldWidth.Value
DanmakuWidth = sldWidth.Value
txtWidth.text = DanmakuWidth & " twip"
ConOut "更改的弹幕宽度：" & DanmakuWidth & " twip"
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
frmContainer.Label1.Caption = txtDanmaku.text
End Sub

Private Function Egg(Count As Long)
' 如果计次为 0 就使它为 1，如果为 1 就使它为 2，
' 如果为 2 就显示彩蛋并清零计次，
' 用到上面就是单击一次再双击一次（顺序反过来也可以）
' 就显示彩蛋
If Count = 0 Then
    Count = 1
ElseIf Count = 1 Then
    Count = 2
ElseIf Count = 2 Then
    MsgBox "你发现了彩蛋！", vbOKOnly + vbApplicationModal + vbInformation, "彩蛋"
    frmEgg.Show vbModal, Me
    Count = 0
End If
End Function
