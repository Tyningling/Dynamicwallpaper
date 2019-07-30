VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ling_wallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const myname As String = "動態壁紙類模塊"
'   壁紙窗體設定類模塊 [這個工程沒有引用視頻組件，所以只是一個壁紙模塊]
'想法啓動于2019/7/28
'By: 風凌 |小林
' Blog: https://www.cnblogs.com/lingqingxue/
'   Blog: http://inkhin.com/
'       E-mail: 1919988942@qq.com
'__________________________________________
'
'   原理： 設置父子窗體和創建置頂窗體
'   通過一個C#項目翻譯而來,當然,您可以通過引用播放組件,使它成爲動態壁紙
'   具體的流程與説明我會在我的BLog中指出。[謝謝支持]
'__________________________________________
'Public Enum SendMessageTimeoutFlags
'SMTO_ABORTIFHUNG = 2
'SMTO_BLOCK = 1
'SMTO_ERRORONEXIT = &H20
'SMTO_NORMAL = 0
'SMTO_NOTIMEOUTIFNOTHUNG = 8
'End Enum

'Private Const def_MessageA = &H52C '0x52c
'Private Const def_MessageB = &H3E8 '0x3e8
Private Const UNITS = 4294967296#
Private Const MAX_LONG = -2147483648#

Private mlTimer As Long 'Watetime‘s var
Private Desktop As Long, hwnd As Long
Private WithEvents l_Timer As ling_Timer
Attribute l_Timer.VB_VarHelpID = -1
Private l_Image As B_Image

Private Sub Class_Initialize()
Set l_Image = New B_Image
    '測試時引用代碼，暫時不會用到。  小林修改于 19/7/30 12：00
    'Set l_Timer = New ling_Timer
    'l_Timer.Enabled = False
    'l_Timer.numInterval = 1
End Sub
Private Sub Class_Terminate()
Set l_Image = Nothing
    'Set l_Timer = Nothing
End Sub
Public Sub 設置桌面(ByRef b_hwnd As Long)
'設置參數
hwnd = b_hwnd 'Set hwnd
'設置窗體樣式
            Call Frm_MAX(hwnd)
'設置窗體
            Desktop = BackDesktop
'將自身設定為父窗口
            Call win32api.SetParent(hwnd, Desktop)
'模擬刷新桌面
            Call win32api.ShowWindow(hwnd, win32api.SW_HIDE)
            Call win32api.ShowWindow(hwnd, win32api.SW_SHOW)
'l_Timer.Enabled = True
End Sub
'消息会生成两个WorkerW 顶级窗口 所以要枚举不包含“SHELLDLL_DefView”这个的 WorkerW 窗口 隐藏掉。
Private Function BackDesktop() As Long
Dim Desktop_h As Long
Desktop_h = win32api.FindWindowA("Progman", "Program Manager") '取得桌面的顶级窗口
Dim result As Long
result = 0&
Dim state As Long
state = 0
state = win32api.SendMessageTimeoutA(Desktop_h, &H52C, 0&, 0&, 0, &H328, VarPtr(result))
If state <> 1 Then MsgBox "創建壁紙窗體時發生異常！"
'VarPtr(Workerw) = 0&
Workerw = 0&
' Win32.User32.ShowWindow(workerw, Win32.User32.SW_HIDE);
Call EnumWindows(AddressOf EnumWindowsProcA, 0&)
BackDesktop = Desktop_h
End Function
Public Sub Frm_MAX(hwnd As Long)
'設置無邊框樣式
Dim style As Long
style = win32api.GetWindowLongA(hwnd, GWL_STYLE)
style = style And Not win32api.WS_CAPTION
style = style And Not win32api.WS_SYSMENU
style = style And Not win32api.WS_MAXIMIZEBOX
style = style And Not win32api.WS_MINIMIZEBOX
style = style And Not win32api.WS_SIZEBOX
win32api.SetWindowLongA hwnd, win32api.GWL_STYLE, style
Dim swpFlag As Long
swpFlag = win32api.SWP_NOMOVE Or win32api.SWP_NOSIZE Or win32api.SWP_NOZORDER Or win32api.SWP_NOREDRAW Or win32api.SWP_NOOWNERZORDER Or win32api.SWP_NOCOPYBITS Or win32api.SWP_NOACTIVATE Or win32api.SWP_FRAMECHANGED Or win32api.WM_NCCALCSIZE
win32api.SetWindowPos hwnd, 0, 0, 0, 0, 0, swpFlag
'設置窗體尺寸
win32api.MoveWindow hwnd, 0, 0, Screen.Width, Screen.Height, ByVal 1
End Sub
'ling_2019/7/30備注: Taskbar_visible將幫助任務欄隱藏 /// 看清來似乎沒有什麽用處
Public Sub Taskbar_Visble()
Dim hwnd As Long
hwnd = win32api.FindWindowA("Shell_TrayWnd", "")
If win32api.IsWindowVisible(hwnd) <> 0 Then
win32api.ShowWindow hwnd, 0
Else
win32api.ShowWindow hwnd, 1
End If
End Sub
Public Sub 动态壁纸设置(ByRef Aplayer As Player, hwnd As Long, ByVal file As String)
Aplayer.SetConfig 2, App.path & "\Bin\codecs" '載入解碼器
Aplayer.Open file '加載設置視頻
Adaptive_Aspect_ratio Aplayer, hwnd '自適應橫縱比
End Sub
Public Sub 壁紙設置(ByVal hdc As Long, ByVal hwnd As Long, file)
'加載圖片
'GDI+加载图片
Dim wh As RECT
Call l_Image.Init(hdc)
Call win32api.GetWindowRect(hwnd, wh)
Call l_Image.add(file, 0, 0, wh.Right - wh.Left, wh.Bottom - wh.Top, True)
Call l_Image.Drawshow
End Sub
Public Sub Adaptive_Aspect_ratio(ByRef Aplayer As Player, hwnd As Long)
Dim Rect_play As RECT
WateTime 300
'獲取客戶區大小:
Call win32api.GetWindowRect(hwnd, Rect_play)
Dim cxClient As Long
cxClient = Rect_play.Right - Rect_play.Left
Dim cyClient As Long
cyClient = Rect_play.Bottom - Rect_play.Top
'設置橫縱比
Aplayer.SetConfig 204, Me.Get_Aspect_ratio((cxClient), cyClient)
End Sub
'由寬高大小計算橫縱比
Public Function Get_Aspect_ratio(ByVal w As Integer, ByVal h As Integer) As String '返回橫縱比
Dim tmp_A As Integer, tmp_B As Integer, R As Integer
tmp_A = w
tmp_B = h
R = tmp_B Mod tmp_A
Do While R <> 0
tmp_B = tmp_A
tmp_A = R
R = tmp_B Mod tmp_A
Loop
Get_Aspect_ratio = CStr(w / tmp_A) & ";" & CStr(h / tmp_A)
End Function




'————————————————————————————————————————————————————————————————————————————————
Public Sub WateTime(MilliSecond As Long) '忘记从哪搬过来的了，十分抱歉！
    '时间延时函数，单位毫秒，此函数不会凝滞窗体，不会造成窗体假死
    On Error GoTo ErrHandler
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    
    mlTimer = CreateWaitableTimerA(0, True, App.EXEName & "Timer" & format$(Now(), "NNSS"))
    If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If
    dblDelay = CDbl(MilliSecond) * 10000#
    ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    Do
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0
    CloseHandle mlTimer
    mlTimer = 0
    Exit Sub
    
ErrHandler:                                                 'clsWaitableTimer
    Err.Raise Err.Number, Err.Source, "延迟时间" & Err.Description
End Sub






