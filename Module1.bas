Attribute VB_Name = "Module1"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
   (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = -20
Private Const LWA_ALPHA = &H2

Sub main()
Form1.Show
Form2.Show
End Sub
