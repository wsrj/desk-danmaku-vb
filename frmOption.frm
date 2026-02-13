VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选项"
   ClientHeight    =   4275
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   5535
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
   ScaleHeight     =   4275
   ScaleWidth      =   5535
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "应用(&A)"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox chkContainerBorder 
      Caption         =   "显示弹幕的边框（高级）(&D)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnApply_Click()
If chkContainerBorder.Value = 1 Then
    ShowContainerBorder = True
ElseIf chkContainerBorder.Value = 0 Then
    ShowContainerBorder = False
End If
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOK_Click()
If chkContainerBorder.Value = 1 Then
    ShowContainerBorder = True
ElseIf chkContainerBorder.Value = 0 Then
    ShowContainerBorder = False
End If
Unload Me
End Sub

Private Sub Form_Load()
If ShowContainerBorder = True Then
    chkContainerBorder.Value = 1
Else
    chkContainerBorder.Value = 0
End If
End Sub
