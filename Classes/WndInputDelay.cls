VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndInputDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWndInputSender
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)
Private m_Milsecs As Long

Friend Sub New_(Milliseconds As Long)
    m_Milsecs = Milliseconds
End Sub

Public Sub Send()
    Sleep m_Milsecs
End Sub

' v ' ############################## ' v '    Implements IWndInputSender    ' v ' ############################## ' v '
Private Sub IWndInputSender_Send()
    Me.Send
End Sub

