Attribute VB_Name = "Module1"
Public Workerw As Long, Tx As Long, Tx_B As Boolean
Public Function EnumWindowsProcA(ByVal Hwnd As Long, ByVal lParam As Long) As Boolean
If win32api.FindWindowExA(Hwnd, 0&, "SHELLDLL_DefView", "") <> 0& Then
Workerw = win32api.FindWindowExA(&O0, Hwnd, "WorkerW", "")
'Call win32api.ShowWindow(Workerw, SW_SHOW)
Call win32api.ShowWindow(Workerw, SW_HIDE)
End If
'特殊处理：
'__________关于腾讯桌面整理导致的问题
If Not Tx_B Then
Tx = win32api.FindWindowExA(Hwnd, 0&, "TXMiniSkin", "桌面整理")
If Tx <> 0 Then Call win32api.ShowWindow(Tx, SW_HIDE): Tx_B = True
End If
EnumWindowsProcA = True
End Function


