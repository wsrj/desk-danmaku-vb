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
Public Declare Function SetConsoleTitleA Lib "kernel32" ( _
    ByVal lpConsoleTitle As String) As Long

Public danmakuColor As String    '弹幕颜色
Public danmakuWidth As Long      '弹幕宽度
Public danmakuFontName As String '弹幕字体
Public danmakuFontSize As String '弹幕字号
Public danmakuFontText As String '显示字体文本
Public eggCount As Long          '彩蛋计数
Public cmdLine As String         '命令行参数
Public isShowConsole As Boolean  '是否显示控制台

Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const LWA_ALPHA As Long = &H2
Public Const HELP_TEXT As String = "用法：" & vbCrLf & _
    "弹幕神器.exe [/? | /C]" & vbCrLf & _
    "/?, -h, --help - 显示此帮助信息" & vbCrLf & _
    "/C, -c, --console - 启用控制台输出（注意，弹幕神器并不是 CUI 子系统程序）"
Public Const CONTACT_URL As String = _
    "https://space.bilibili.com/3493134929496963"

Public Sub Main()
cmdLine = Command
Select Case cmdLine
    Case "/?", "-?", "-h", "--help"
        isShowConsole = False
        MsgBox HELP_TEXT, vbOKOnly + vbInformation, "弹幕神器"
        End
    Case "/C", "/c", "-c", "--console"
        isShowConsole = True
        AllocConsole
        SetConsoleTitleA "控制台 - 弹幕神器"
        ConOut "欢迎使用弹幕神器！", False
        ConOut "命令行参数：" & cmdLine
    Case vbNullString
        isShowConsole = False
    Case Else
        Select Case MsgBox("参数不正确。" & vbCrLf & HELP_TEXT & _
                    vbCrLf & vbCrLf & "是否打开主界面？", _
                    vbYesNo + vbCritical, "错误 - 弹幕神器")
            Case vbYes
                isShowConsole = False
            Case vbNo
                isShowConsole = False
                End
        End Select
End Select
frmMain.Show
frmContainer.Show
frmContainer.Move Screen.Width, Screen.Height / 20
frmMain.Timer1.Interval = 10
End Sub

Public Sub ConOut(text As String, Optional isPrintTime As Boolean = True)
If Not isShowConsole Then Exit Sub
Open "CONOUT$" For Output As #1
If isPrintTime = True Then
    Print #1, "[" & Now & "] " & text
ElseIf isPrintTime = False Then
    Print #1, text
End If
Close #1
End Sub

Public Function ReadXML(xmlFile As String, nodeName As String) As String
Dim xmlDoc As New msxml2.DOMDocument60, node As Object
xmlDoc.Load xmlFile
Set node = xmlDoc.selectSingleNode(nodeName)
If Not node Is Nothing Then
    ReadXML = node.text
    Set xmlDoc = Nothing
Else
    MsgBox "配置文件格式不正确。", vbOKOnly + vbCritical, "错误"
    ReadXML = vbNullString
    Set xmlDoc = Nothing
End If
End Function
