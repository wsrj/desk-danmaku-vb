VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEgg 
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "彩蛋"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6375
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
   ScaleHeight     =   5055
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar staBottom 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Text            =   "CapsLk"
            TextSave        =   "CapsLk"
            Object.ToolTipText     =   "Caps Lock 状态"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   2752
            Text            =   "还没有按有效的键"
            TextSave        =   "还没有按有效的键"
            Object.ToolTipText     =   "按下的键"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3810
            Text            =   "使用 [W] [A] [S] [D] 操控"
            TextSave        =   "使用 [W] [A] [S] [D] 操控"
            Object.ToolTipText     =   "提示"
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmEgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
staBottom.Panels(2).text = Chr$(KeyAscii)
' 键盘大概是这样：
'  Q  [W]  E   R   T …
' [A] [S] [D]  F   G …
' …
Select Case Chr$(KeyAscii)
    ' 按下向上
    Case "W", "w"
        If Shape1.Top >= 0 Then
            Shape1.Top = Shape1.Top - Screen.TwipsPerPixelY
        End If
    ' 按下向下
    Case "S", "s"
        If Shape1.Top <= Me.Height - Shape1.Height Then
            Shape1.Top = Shape1.Top + Screen.TwipsPerPixelY
        End If
    ' 按下向左
    Case "A", "a"
        If Shape1.Left >= 0 Then
            Shape1.Left = Shape1.Left - Screen.TwipsPerPixelX
        End If
    ' 按下向右
    Case "D", "d"
        If Shape1.Left < Me.Width - Shape1.Width Then
            Shape1.Left = Shape1.Left + Screen.TwipsPerPixelX
        End If
    Case Else
        Exit Sub
End Select
CheckPos Shape1
End Sub

Private Function CheckPos(Optional object As Shape = Shape1)
' 检查左边距
If object.Left < 0 Then
    object.Left = 0
    staBottom.Panels(3).text = "超出窗口范围"
ElseIf object.Left > Me.Width - object.Width Then
    object.Left = Me.Width - object.Width
    staBottom.Panels(3).text = "超出窗口范围"
Else
    staBottom.Panels(3).text = "使用 [W] [A] [S] [D] 操控"
End If
' 检查顶边距
If object.Top < 0 Then
    object.Top = 0
    staBottom.Panels(3).text = "超出窗口范围"
ElseIf object.Top > Me.Height - object.Height - staBottom.Height Then
    object.Top = Me.Height - object.Height - staBottom.Height
    staBottom.Panels(3).text = "超出窗口范围"
Else
    staBottom.Panels(3).text = "使用 [W] [A] [S] [D] 操控"
End If
End Function

