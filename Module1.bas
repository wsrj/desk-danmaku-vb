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
'Public Declare Function TaskDialog Lib "commctl32" ( _
'    ByVal hwndOwner As Long, _
'    ByVal hInstance As String, _
'    ByVal pszWindowTitle As String, _
'    ByVal pszMainInstruction As String, _
'    ByVal pszContent As String, _
'    ByVal dwCommonButtons As Long, _
'    ByVal pszIcon As Long, _
'    pnButton As Long _
') As Long

Sub Main()
Form2.Show
Form1.Show
Form2.Move Screen.Width, Screen.Height / 20
Form1.Timer1.Interval = 10
End Sub
