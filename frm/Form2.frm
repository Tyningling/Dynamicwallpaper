VERSION 5.00
Object = "{97830570-35FE-4195-83DE-30E79B718713}#1.0#0"; "APlayer_3.9.10.806.dll"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin APlayer3LibCtl.Player Player1 
         Height          =   2535
         Left            =   0
         OleObjectBlob   =   "Form2.frx":0000
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
Player1.Close
End Sub
