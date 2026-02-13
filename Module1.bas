Attribute VB_Name = "Module1"
Public Declare Function GetPrivateProfileStringA Lib "kernel32" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileStringA Lib "kernel32" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
Public Declare Function ShellAboutA Lib "shell32" ( _
    ByVal hWnd As Long, _
    ByVal szApp As String, _
    ByVal szOtherStuff As String, _
    ByVal hIcon As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SetWindowLongA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLongA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlph As Byte, _
    ByVal dwFlags As Long) As Long
'Const GWL_EXSTYLE = (-20)
'Const WS_EX_LAYERED = &H80000
'Const LWA_ALPHA = &H2

Sub Main()
frmMain.Show
frmContainer.Show
frmContainer.Move Screen.Width, Screen.Height / 20
frmMain.Timer1.Interval = 10
End Sub
