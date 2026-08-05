VERSION 5.00
Begin VB.Form frmContainer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "弹幕"
   ClientHeight    =   3060
   ClientLeft      =   20550
   ClientTop       =   60
   ClientWidth     =   8055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8055
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' 设置为分层窗口
SetWindowLongA Me.hwnd, GWL_EXSTYLE, _
    GetWindowLongA(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
' 设置透明度
SetLayeredWindowAttributes Me.hwnd, vbWhite, 0, 1
'Me.Width = DanmakuWidth
Me.Height = Screen.Height - Screen.Height / 20
Label1.Caption = frmMain.txtDanmaku.text
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
End If
End Sub

Private Sub Form_Resize()
Label1.Move 0, 0, Me.Width, Me.Height
End Sub

