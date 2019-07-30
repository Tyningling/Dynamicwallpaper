VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ling_Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Timer()
Public numInterval As Long
Private blnEnabled As Boolean
Private Sub Class_Initialize()
    numInterval = 0
    blnEnabled = False
End Sub
Private Sub Class_Terminate()
    '
End Sub
Public Property Get Enabled() As Boolean
    Enabled = blnEnabled
End Property
Public Property Let Enabled(ByVal blnValue As Boolean)
    If Not blnEnabled And blnValue Then
        blnEnabled = blnValue
    
        Call Sub_Timer
    Else
        blnEnabled = blnValue
    End If
    
End Property
Private Sub Sub_Timer()
    Static numTime As Single
    numTime = Timer
    
    Do
        If Not blnEnabled Then Exit Do
        
        If (Timer - numTime) > numInterval Then
            numTime = Timer
            RaiseEvent Timer
        End If
        
        DoEvents
    Loop
End Sub


