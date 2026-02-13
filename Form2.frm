VERSION 5.00
Begin VB.Form frmContainer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3060
   ClientLeft      =   20550
   ClientTop       =   60
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8055
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
Private Sub Form_Load()
' 设置为分层窗口
SetWindowLongA Me.hWnd, -20, GetWindowLongA(Me.hWnd, -20) Or 524288
' 设置透明度
SetLayeredWindowAttributes Me.hWnd, vbWhite, 0, 1
Me.Width = DanmakuWidth
Me.Height = Screen.Height - Screen.Height / 20
Label1.Caption = frmMain.txtDanmaku.Text
End Sub
