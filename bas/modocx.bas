Attribute VB_Name = "modocx"
'********************************************************
'**
'**ģ �� ����modBSkin
'**
'**˵    ����ͨ��ģ��
'**
'********************************************************
Option Explicit

Private Declare Function ReleaseCapture Lib "User32" () As Long '������Ⱦ
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'����ִ��
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'�ƶ��������HWND�Ŀؼ�,д���Ϊ�˷������
Sub MoveForm(frm As Object)
    If TypeOf frm Is Form Then
        If frm.Width >= Screen.Width - 600 And _
            frm.Height >= Screen.Height - 600 Then Exit Sub
    End If

    Call ReleaseCapture
    SendMessage frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'ϵͳ��ǰ·��
Public Function APP_PATH() As String
    ChDrive App.path
    ChDir App.path
    APP_PATH = DirFix(App.path)
End Function

'Ŀ¼�Զ���"\"
Private Function DirFix(ByVal PathName As String) As String
    If Right(PathName, 1) = "\" Then DirFix = PathName Else DirFix = PathName + "\"
End Function

'�ж��ļ����Ƿ����
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

'�ж��ļ��Ƿ����
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

'ͨ���ļ�·����ȡ�ļ���
Public Function GetFileFromPath(ByVal sPath As String) As String
    Dim nPos As Integer
    
    nPos = InStrRev(sPath, "\")
    If nPos > 0 Then
        GetFileFromPath = Mid$(sPath, nPos + 1)
    Else
        GetFileFromPath = sPath
    End If
End Function

'����ҳ
Public Sub OpenURL(ByVal sUrl As String)
    ShellExecute 0&, "open", sUrl, vbNullString, vbNullString, vbNormalNoFocus
End Sub

'�������ļ�
Public Function OpenFiles(ByVal sFilePath As String)
    ShellExecute 0&, vbNullString, sFilePath, vbNullString, vbNullString, vbNormalNoFocus
End Function

'ע��OCX�ؼ�
Public Function RegOCX(ByVal OCX As String)
    Dim ocxPath As String
    Dim LsStr As String
    
    LsStr = Environ("windir") & "\system32\" & OCX
    ocxPath = APP_PATH & OCX
    If Dir(LsStr) = "" Then FileCopy ocxPath, LsStr

    Shell "regsvr32.exe " & APP_PATH & OCX, vbHide
End Function

'��ע��OCX�ؼ�
Public Sub UniOCX(ByVal OCX As String)
    Shell "regsvr32.exe /u " & APP_PATH & OCX, vbHide
End Sub

'ע��ActiveX EXE
Public Sub ActiveX(ByVal EXE As String)
    If GetFileFromPath(EXE) = "" Then Exit Sub
    Shell Chr(34) & EXE & Chr(34) & " /regserver", vbHide
End Sub

'��ע��ActiveX EXE
Public Sub UnActiveX(ByVal EXE As String)
    If GetFileFromPath(EXE) = "" Then Exit Sub
    Shell Chr(34) & EXE & Chr(34) & " /unregserver", vbHide
End Sub


