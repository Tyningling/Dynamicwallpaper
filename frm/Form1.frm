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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command11 
      Caption         =   "�O�ô��w��ڼ����w"
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "�O����Ļ������ҕ�l����"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "����-10"
      Height          =   615
      Left            =   8160
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����+10"
      Height          =   615
      Left            =   7080
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�^�m/��ͣ���űڼ�ҕ�l"
      Height          =   615
      Left            =   7080
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�O��ҕ�l������4:3"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�O��ҕ�l������16:9"
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
'   �ڼ�ʾ��
'�뷨������2019/7/28
'By: �L�� |С��
' Blog: https://www.cnblogs.com/lingqingxue/
'   Blog: http://inkhin.com/
'       E-mail: 1919988942@qq.com
'__________________________________________
'
'   ���P�ļ��������̕����ҵ�BLog��ָ����[�x�x֧��]
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
a.�O������ Form2.hwnd
a.��̬��ֽ���� Form2.Player1, Form2.Picture1.hwnd, "D:\Desktop\data\ʯ��ͼ��ѧ\2-3-������ ɨ��ת��Բ��.mp4"
End Sub
Private Sub Command1_Click()
'Set Me.ShowInTaskbar = False Ո�ڴ��w�A���O�� [ֻ�x����]
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
a.�O������ Form2.hwnd
a.�ڼ��O�� Form2.Picture1.hdc, Form2.Picture1.hwnd, App.path & "\bg.png"
End Sub
Private Sub Command10_Click()
a.Adaptive_Aspect_ratio Form2.Player1, Form2.Picture1.hwnd
End Sub
Private Sub Command11_Click()
a.�O������ Form2.hwnd
End Sub
Private Sub Command2_Click()
'119 - Loop play                      int      R/W         ��ȡ��������ѭ������, 0-�Զ�, 1-ѭ��, 2-��ѭ��, Ĭ��0 (�Զ�ģʽ��, GIF ���Զ�ѭ��, ������ʽĬ�ϲ�ѭ��)
'120 - No close when complete         int      R/W         ��ȡ���������Ƿ񲥷���ɲ��Զ� Close (�Զ� Close �᷵�� PS_READY ״̬)��0-�Զ� Close��1-���Զ� Close��Ĭ�� 0������Ϊ1ʱ�����Ž������Զ� Close�������߻����� SetPositon �������ţ������ǻᷢ�� OnEvent(PLAYCOMPLETE) �¼�
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
  '      PS_READY      = 0,  // ׼������
   '     PS_OPENING    = 1,  // ���ڴ�
    '    PS_PAUSING    = 2,  // ������ͣ
     '   PS_PAUSED     = 3,  // ��ͣ��
      '  PS_PLAYING    = 4,  // ���ڿ�ʼ����
       ' PS_PLAY       = 5,  // ������
       ' PS_CLOSING    = 6,  // ���ڿ�ʼ�ر�
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
Me.Command1.Caption = "С�ֵıڼ�_�yԇ�_ʼ"
Me.Command4.Caption = "С�ֵĄӑB�ڼ�_�yԇ�_ʼ"
Me.Command3.Caption = "�ڼ����w���w"
Me.Command2.Caption = "����/ȡ�� ѭ������"
MsgBox "С��������,��ǰʹ�õ�Aplayer�汾��  ��" & Form2.Player1.GetVersion
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set a = Nothing
End
End Sub
