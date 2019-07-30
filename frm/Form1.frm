VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   9960
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command11 
      Caption         =   "設置窗體為壁紙窗體"
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "設置屏幕比例為視頻比例"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "音量-10"
      Height          =   615
      Left            =   8160
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "音量+10"
      Height          =   615
      Left            =   7080
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "繼續/暫停播放壁紙視頻"
      Height          =   615
      Left            =   7080
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "設置視頻比例為4:3"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "設置視頻比例為16:9"
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   9240
      TabIndex        =   9
      Top             =   1560
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   壁紙示例
'想法啓動于2019/7/28
'By: 風凌 |小林
' Blog: https://www.cnblogs.com/lingqingxue/
'   Blog: http://inkhin.com/
'       E-mail: 1919988942@qq.com
'__________________________________________
'
'   相關的細節和流程會在我的BLog中指出。[謝謝支持]
'__________________________________________
Option Explicit
Dim a As New ling_wallpaper
Private Sub Command4_Click()
Set a = New ling_wallpaper
Form2.Show
Form2.WindowState = 2
With Form2.Picture1
.Width = Form2.Width
.Height = Form2.Height
.Left = 0
.Top = 0
End With
With Form2.Player1
.Width = Form2.Width
.Height = Form2.Height
.Left = 0
.Top = 0
End With
Form2.Player1.Visible = True
a.設置桌面 Form2.hwnd
a.动态壁纸设置 Form2.Player1, Form2.Picture1.hwnd, "D:\Desktop\data\石大图形学\2-3-第三节 扫描转换圆弧.mp4"
End Sub
Private Sub Command1_Click()
'Set Me.ShowInTaskbar = False 請在窗體預先設置 [只讀屬性]
Set a = New ling_wallpaper
Form2.Show
Form2.WindowState = 2
With Form2.Picture1
.Width = Form2.Width
.Height = Form2.Height
.Left = 0
.Top = 0
End With
With Form2.Player1
.Width = Form2.Width
.Height = Form2.Height
.Left = 0
.Top = 0
End With
a.設置桌面 Form2.hwnd
a.壁紙設置 Form2.Picture1.hdc, Form2.Picture1.hwnd, App.path & "\bg.png"
End Sub
Private Sub Command10_Click()
a.Adaptive_Aspect_ratio Form2.Player1, Form2.Picture1.hwnd
End Sub
Private Sub Command11_Click()
a.設置桌面 Form2.hwnd
End Sub
Private Sub Command2_Click()
'119 - Loop play                      int      R/W         获取或者设置循环播放, 0-自动, 1-循环, 2-不循环, 默认0 (自动模式中, GIF 会自动循环, 其他格式默认不循环)
'120 - No close when complete         int      R/W         获取或者设置是否播放完成不自动 Close (自动 Close 会返回 PS_READY 状态)，0-自动 Close，1-不自动 Close，默认 0，设置为1时，播放结束不自动 Close，调用者还可以 SetPositon 继续播放，但还是会发送 OnEvent(PLAYCOMPLETE) 事件
If Form2.Player1.GetConfig(119) = 0 Then
Form2.Player1.SetConfig 119, 1
Form2.Player1.SetConfig 120, 1
Else
Form2.Player1.SetConfig 119, 0
Form2.Player1.SetConfig 120, 0
End If
End Sub
Private Sub Command3_Click()
Unload Form2
End Sub
Private Sub Command5_Click()
Form2.Player1.SetConfig 204, "16;9"
End Sub
Private Sub Command6_Click()
Form2.Player1.SetConfig 204, "4;3"
End Sub
'    Enum PLAY_STATE
 '   {
  '      PS_READY      = 0,  // 准备就绪
   '     PS_OPENING    = 1,  // 正在打开
    '    PS_PAUSING    = 2,  // 正在暂停
     '   PS_PAUSED     = 3,  // 暂停中
      '  PS_PLAYING    = 4,  // 正在开始播放
       ' PS_PLAY       = 5,  // 播放中
       ' PS_CLOSING    = 6,  // 正在开始关闭
    '};
'
Private Sub Command7_Click()
If Form2.Player1.GetState = 5 Then
    Form2.Player1.Pause
Else
    Form2.Player1.Play
End If
End Sub

Private Sub Command8_Click()
Form2.Player1.SetVolume (Form2.Player1.GetVolume + 10)
Label1.Caption = Form2.Player1.GetVolume
End Sub
Private Sub Command9_Click()
Form2.Player1.SetVolume (Form2.Player1.GetVolume - 10)
Label1.Caption = Form2.Player1.GetVolume
End Sub
Private Sub Form_Load()
'OpenFiles App.path & "/install.bat"
Me.Command1.Caption = "小林的壁紙_測試開始"
Me.Command4.Caption = "小林的動態壁紙_測試開始"
Me.Command3.Caption = "壁紙窗體解體"
Me.Command2.Caption = "设置/取消 循环播放"
MsgBox "小林提醒您,當前使用的Aplayer版本是  ：" & Form2.Player1.GetVersion
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set a = Nothing
End
End Sub
