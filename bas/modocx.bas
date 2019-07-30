Attribute VB_Name = "modocx"
'********************************************************
'**
'**模 块 名：modBSkin
'**
'**说    明：通用模块
'**
'********************************************************
Option Explicit

Private Declare Function ReleaseCapture Lib "User32" () As Long '界面渲染
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'程序执行
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'移动窗体或有HWND的控件,写这个为了方便调用
Sub MoveForm(frm As Object)
    If TypeOf frm Is Form Then
        If frm.Width >= Screen.Width - 600 And _
            frm.Height >= Screen.Height - 600 Then Exit Sub
    End If

    Call ReleaseCapture
    SendMessage frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'系统当前路径
Public Function APP_PATH() As String
    ChDrive App.path
    ChDir App.path
    APP_PATH = DirFix(App.path)
End Function

'目录自动补"\"
Private Function DirFix(ByVal PathName As String) As String
    If Right(PathName, 1) = "\" Then DirFix = PathName Else DirFix = PathName + "\"
End Function

'判断文件夹是否存在
Public Function FolderExists(ByVal sFolder As String) As Boolean
On Error Resume Next
    If Replace(sFolder, vbCrLf, "") = "" Then
        FolderExists = False
        Exit Function
    End If
    If Dir(sFolder, vbDirectory) = "" Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

'判断文件是否存在
Public Function FileExists(ByVal sFile As String) As Boolean
On Error Resume Next
    If Replace(sFile, vbCrLf, "") = "" Then
        FileExists = False
        Exit Function
    End If
    If Dir(sFile) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

'通过文件路径获取文件名
Public Function GetFileFromPath(ByVal sPath As String) As String
    Dim nPos As Integer
    
    nPos = InStrRev(sPath, "\")
    If nPos > 0 Then
        GetFileFromPath = Mid$(sPath, nPos + 1)
    Else
        GetFileFromPath = sPath
    End If
End Function

'打开网页
Public Sub OpenURL(ByVal sUrl As String)
    ShellExecute 0&, "open", sUrl, vbNullString, vbNullString, vbNormalNoFocus
End Sub

'打开任意文件
Public Function OpenFiles(ByVal sFilePath As String)
    ShellExecute 0&, vbNullString, sFilePath, vbNullString, vbNullString, vbNormalNoFocus
End Function

'注册OCX控件
Public Function RegOCX(ByVal OCX As String)
    Dim ocxPath As String
    Dim LsStr As String
    
    LsStr = Environ("windir") & "\system32\" & OCX
    ocxPath = APP_PATH & OCX
    If Dir(LsStr) = "" Then FileCopy ocxPath, LsStr

    Shell "regsvr32.exe " & APP_PATH & OCX, vbHide
End Function

'反注册OCX控件
Public Sub UniOCX(ByVal OCX As String)
    Shell "regsvr32.exe /u " & APP_PATH & OCX, vbHide
End Sub

'注册ActiveX EXE
Public Sub ActiveX(ByVal EXE As String)
    If GetFileFromPath(EXE) = "" Then Exit Sub
    Shell Chr(34) & EXE & Chr(34) & " /regserver", vbHide
End Sub

'反注册ActiveX EXE
Public Sub UnActiveX(ByVal EXE As String)
    If GetFileFromPath(EXE) = "" Then Exit Sub
    Shell Chr(34) & EXE & Chr(34) & " /unregserver", vbHide
End Sub


