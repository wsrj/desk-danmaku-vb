VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "µ¯Ä»ÉñÆ÷"
   ClientHeight    =   5295
   ClientLeft      =   10320
   ClientTop       =   5655
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Íâ¹Û"
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
      Begin ComctlLib.Slider Slider1 
         Height          =   615
         Left            =   1920
         TabIndex        =   11
         Top             =   2760
         Width           =   4215
         _ExtentX        =   7435
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
         Caption         =   "Ñ¡ÔñÑÕÉ«(&I)¡­"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtFont 
         Appearance      =   0  'Flat
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
         Caption         =   "Ñ¡Ôñ×ÖÌå(&P)¡­"
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
         Text            =   "ßÙÁ¨ßÙÁ¨( ©b- ©b)¤Ä¥í¸É±­~"
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "µ¯Ä»×î´ó¿í¶È"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "µ¯Ä»ÑÕÉ«"
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
         Caption         =   "µ¯Ä»×ÖÌå"
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
         Caption         =   "µ¯Ä»ÎÄ×Ö"
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
      Caption         =   "·¢ËÍµ¯Ä»"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Caption         =   "ÎÄ¼þ(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "´ò¿ª(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "±£´æ(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "ÍË³ö(&X)"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DanmakuColor
Dim DanmakuFont
Private Sub btnColorPicker_Click()
CommonDialog1.ShowColor
DanmakuColor = CommonDialog1.Color
txtColor.Text = DanmakuColor
End Sub
Private Sub btnFontPicker_Click()
CommonDialog1.ShowFont
DanmakuFont = "×ÖÌå£º" & CommonDialog1.FontName & "£¬×ÖºÅ£º" & CommonDialog1.FontSize
txtFont.Text = DanmakuFont
End Sub
Private Sub mnuExit_Click()
End
End Sub
