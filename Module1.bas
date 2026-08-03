Attribute VB_Name = "Module1"
Option Explicit

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
Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Declare Function SetConsoleTitleA Lib "kernel32" (ByVal lpConsoleTitle As String) As Long

Public DanmakuColor As Long      '颜色变量
Public DanmakuWidth As Long      '宽度变量
Public DanmakuFontName As String '字体变量
Public DanmakuFontSize As Long   '字号变量
Public DanmakuFontText As String '显示字体文本
Public ShowEgg As Long
Public Cmd As String
Public ShowConsole As Boolean

'Const GWL_EXSTYLE = (-20)
'Const WS_EX_LAYERED = &H80000
'Const LWA_ALPHA = &H2
Public Const HELP_TEXT As String = "用法：" & vbCrLf & _
    "弹幕神器.exe [/? | /C]" & vbCrLf & _
    "/?, -h, --help - 显示此帮助信息" & vbCrLf & _
    "/C, -c, --console - 启用控制台输出（注意，弹幕神器并不是 CUI 子系统程序）"

Public Sub Main()
Cmd = Command
Select Case Cmd
    Case "/?", "-?", "-h", "--help"
        ShowConsole = False
        MsgBox HELP_TEXT, vbOKOnly + vbInformation, "弹幕神器"
        End
    Case "/C", "/c", "-c", "--console"
        ShowConsole = True
        AllocConsole
        SetConsoleTitleA "控制台 - 弹幕神器"
        ConOut "欢迎使用弹幕神器！", False
        ConOut "命令行参数：" & Cmd
    Case ""
        ShowConsole = False
    Case Else
        Select Case MsgBox("参数不正确。" & vbCrLf & HELP_TEXT & vbCrLf & vbCrLf & "是否打开主界面？", vbYesNo + vbCritical, "错误 - 弹幕神器")
            Case vbYes
                ShowConsole = False
            Case vbNo
                ShowConsole = False
'                MsgBox HELP_TEXT, vbOKOnly + vbInformation, "帮助"
                End
        End Select
End Select
frmMain.Show
frmContainer.Show
frmContainer.Move Screen.Width, Screen.Height / 20
frmMain.Timer1.Interval = 10
End Sub

Public Sub ConOut(text As String, Optional printTime As Boolean = True)
If Not ShowConsole Then Exit Sub
Open "CONOUT$" For Output As #1
If printTime = True Then
    Print #1, "[" & Now & "] " & text
ElseIf printTime = False Then
    Print #1, text
End If
Close #1
End Sub

Public Function ReadXML(xmlFile As String, nodeName As String) As String
Dim xmlDoc As msxml2.DOMDocument60, node As Object
Set xmlDoc = New msxml2.DOMDocument60
xmlDoc.Load xmlFile
Set node = xmlDoc.selectSingleNode(nodeName)
If Not node Is Nothing Then
    ReadXML = node.text
    Set xmlDoc = Nothing
    Exit Function
Else
    MsgBox "配置文件格式不正确。", vbOKOnly + vbCritical, "错误"
End If
End Function
