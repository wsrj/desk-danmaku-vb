Attribute VB_Name = "Module1"
'Public Declare Function GetPrivateProfileStringA Lib "kernel32" _
'    Alias "GetPrivateProfileString" ( _
'    ByVal lpApplicationName As String, _
'    ByVal lpKeyName As Any, _
'    ByVal lpDefault As String, _
'    ByVal lpReturnedString As String, _
'    ByVal nSize As Long, _
'    ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileStringA Lib "kernel32" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
Public Declare Function ShellAboutA Lib "shell32" ( _
    hWnd As Long, _
    szApp As String, _
    szOtherStuff As String) As Long

Sub Main()
Form2.Show
Form1.Show
Form2.Move Screen.Width, 0
Form1.Timer1.Interval = 200
End Sub
