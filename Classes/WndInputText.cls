VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndInputText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWndInputSender
Private m_hWndSend As LongPtr
Private m_hWndRecv As LongPtr
Private m_Name     As String
Private m_Text     As String
Private m_DelayMS  As Integer
Private m_KInputs  As Collection '(Of IWndInputSender(Of WndInputKeybd | WndInputDelay))

Friend Sub New_(Name As String, Text As String, hWndSender As LongPtr, hWndReceiver As LongPtr)
    m_hWndSend = hWndSender: m_hWndRecv = hWndReceiver
    m_Name = Name: Me.Text = Text
End Sub
Friend Sub NewC(other As WndInputText)
    With other
        m_Name = .Name
        m_Text = .Text
        m_DelayMS = .DelayMS
    End With
End Sub
Friend Function Clone() As WndInputText
    Set Clone = New WndInputText
    WndInputText.NewC Me
End Function

Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal Value As String)
    m_Name = Value
End Property

Public Property Get Text() As String
    Text = m_Text
End Property
Public Property Let Text(ByVal Value As String)
    m_Text = Value
    Set m_KInputs = CreateWndInpKeybdsFromText(m_Text)
End Property

Public Property Get HasDelay() As String
    HasDelay = 0 < m_DelayMS
End Property

Public Property Get DelayMS() As Integer
    DelayMS = m_DelayMS
End Property
Public Property Let DelayMS(ByVal Value As Integer)
    m_DelayMS = Value
End Property

Public Sub Send()
    If m_KInputs Is Nothing Then Exit Sub
    If m_KInputs.Count = 0 Then Exit Sub
    Dim wi As IWndInputSender
    For Each wi In m_KInputs: wi.Send: Next
End Sub

Private Function CreateWndInpKeybdsFromText(aText As String) As Collection
    'now we generate a new WndInputs object
    Dim l As Long: l = Len(m_Text)
    If l = 0 Then Exit Sub
    Dim ins As WndInputs: Set ins = MNew.WndInputs(m_hWndSend, m_hWndRecv, "KeybInputs")
    Dim i As Long, ch As String, ki As WndInputKeybd, col As Collection
    For i = 1 To l
        ch = Mid(m_Text, i, 1)
        Set col = CreateWndInpKeybdsFromChar(ch)
        For Each ki In col: ins.Add ki: Next
    Next

End Function

Private Function CreateWndInpKeybdsFromChar(ByVal char As String) As Collection
    Dim col As Collection: Set col = New Collection
    Dim ic  As Integer: ic = AscW(ch)
    Dim icu As Integer: icu = AscW(UCase(ch))
    Dim icl As Integer: icl = AscW(LCase(ch))
    If ic = icu Then col.Add MNew.WndInputKeybd(VK_SHIFT, 0, KEYEVENTF_KEYDOWN)
    
    Select Case True
    Case ic = icu
    End Select
    Set CreateWndInpKeybdsFromChar = col
End Function


' v ' ############################## ' v '    Implements IWndInputSender    ' v ' ############################## ' v '
Private Sub IWndInputSender_Send()
    Me.Send
End Sub

