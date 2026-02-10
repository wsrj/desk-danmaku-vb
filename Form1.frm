VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "弹幕神器"
   ClientHeight    =   5295
   ClientLeft      =   10320
   ClientTop       =   5655
   ClientWidth     =   7140
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
   ScaleHeight     =   5295
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
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
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox txtWidth 
         Appearance      =   0  'Flat
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DanmakuColor '颜色变量
Dim DanmakuFont '字体变量
Dim DanmakuWidth '宽度变量

Private Sub Form_Initialize()
sldWidth.Max = Screen.Width
sldWidth.TickFrequency = sldWidth.Max / 10
End Sub

Private Sub btnColorPicker_Click()
CommonDialog1.ShowColor '弹出选择颜色
DanmakuColor = CommonDialog1.Color
txtColor.Text = DanmakuColor '在标签中显示颜色
End Sub

Private Sub btnFontPicker_Click()
CommonDialog1.ShowFont '弹出选择字体
DanmakuFont = "字体：" & CommonDialog1.FontName & "，字号：" & CommonDialog1.FontSize
txtFont.Text = DanmakuFont '在标签中显示字体
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub sldWidth_Change()
DanmakuWidth = sldWidth.Value
txtWidth = DanmakuWidth
End Sub
