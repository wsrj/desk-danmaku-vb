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
Me.Width = DanmakuWidth
Me.Height = Screen.Height - Screen.Height / 20
Label1.Caption = frmMain.txtDanmaku.Text
End Sub
