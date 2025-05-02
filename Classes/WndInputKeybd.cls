VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndInputKeybd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-keybdinput
Private Type KEYBDINPUT
    dwType      As EInputType      ' 4
    wVk         As Integer         ' 2
    wScan       As Integer         ' 2
    dwFlags     As EKeyEventFlags  ' 4
    Time        As Long            ' 4
    dwExtraInfo As LongPtr         ' 4/8
End Type                     ' Sum: 20

Private Type PKEYBDINPUT
    SAPtr As TUDTPtr
    Arr() As KEYBDINPUT
End Type

Private m_Value As PKEYBDINPUT
Private m_Data  As KEYBDINPUT

Private Sub Class_Initialize()
    m_Data.dwType = EInputType.INPUT_KEYBOARD
    MPtr.New_UDTPtr m_Value.SAPtr, FADF_AUTO Or FADF_FIXEDSIZE Or FADF_RECORD, 16 '24
    MPtr.SAPtr(ArrPtr(m_Value.Arr)) = m_Value.SAPtr.pSA
    m_Value.SAPtr.pvData = VarPtr(m_Data)
End Sub

Private Sub Class_Terminate()
    MPtr.ZeroSAPtr ArrPtr(m_Value.Arr)
End Sub
Friend Sub New_(VirtKey As Integer, Scan As Integer, Flags As EKeyEventFlags, aTime As Long)
    'If pData Then m_Value.SAPtr.pvData = pData
    With Me: .VirtKeyCode = VirtKey: .Scan = Scan: .Flags = Flags: .Time = aTime: End With
End Sub
'Copy to MNew:
'Public Function WndInputKeybd(byval pData As LongPtr, byval VirtKey As EVirtualKeyCodes, byval Scan As Integer, byval Flags As Long, byval aTime As Long) As WndInputKeybd
'    Set WndInputKeybd = New WndInputKeybd: WndInputKeybd.New_ pData, dX, dY, MouseData, Flags, aTime
'End Function

Friend Sub NewC(other As WndInputKeybd)
    With other: New_ .VirtKeyCode, .Scan, .Flags, .Time: End With
End Sub
Public Function Clone() As WndInputKeybd
    Set Clone = New WndInputKeybd: Clone.NewC Me
End Function

Public Property Get Key() As String
    Key = "Ki: " & MVirtualKeys.EVirtualKeyCodes_ToStr(Me.VirtKeyCode)
End Property

Public Property Get Ptr() As LongPtr
    Ptr = m_Value.SAPtr.pvData
End Property
Public Property Let Ptr(ByVal Value As LongPtr)
    m_Value.SAPtr.pvData = Value
End Property

Public Sub SetData()
    m_Value.Arr(0) = m_Data
End Sub

'A virtual-key code. The code must be a value in the range 1 to 254. If the dwFlags member specifies KEYEVENTF_UNICODE, wVk must be 0.
Public Property Get VirtKeyCode() As Integer
    VirtKeyCode = m_Value.Arr(0).wVk
End Property
Public Property Let VirtKeyCode(ByVal Value As Integer)
    m_Value.Arr(0).wVk = Value
End Property

'A hardware scan code for the key. If dwFlags specifies KEYEVENTF_UNICODE, wScan specifies a Unicode character which is to be sent to the foreground application.
Public Property Get Scan() As Integer
    Scan = m_Value.Arr(0).wScan
End Property
Public Property Let Scan(ByVal Value As Integer)
    m_Value.Arr(0).wScan = Value
End Property

'Specifies various aspects of a keystroke. This member can be certain combinations of the following values:
'Value                           |   Meaning
'KEYEVENTF_EXTENDEDKEY = &H0001  |   If specified, the wScan scan code consists of a sequence of two bytes, where the first byte has a value of 0xE0. See Extended-Key Flag for more info.
'KEYEVENTF_KEYUP       = &H0002  |   If specified, the key is being released. If not specified, the key is being pressed.
'KEYEVENTF_UNICODE     = &H0004  |   If specified, the system synthesizes a VK_PACKET keystroke. The wVk parameter must be zero. This flag can only be combined with the KEYEVENTF_KEYUP flag. For more information, see the Remarks section.
'KEYEVENTF_SCANCODE    = &H0008  |   If specified, wScan identifies the key and wVk is ignored.
Public Property Get Flags() As EKeyEventFlags
    Flags = m_Value.Arr(0).dwFlags
End Property
Public Property Let Flags(ByVal Value As EKeyEventFlags)
    m_Value.Arr(0).dwFlags = Value
End Property

'The time stamp for the event, in milliseconds. If this parameter is zero, the system will provide its own time stamp.
Public Property Get Time() As Long
    Time = m_Value.Arr(0).Time
End Property
Public Property Let Time(ByVal Value As Long)
    m_Value.Arr(0).Time = Value
End Property

'An additional value associated with the keystroke. Use the GetMessageExtraInfo function to obtain this information.
Public Property Get ExtraInfo() As LongPtr
    ExtraInfo = m_Value.Arr(0).dwExtraInfo
End Property
Public Property Let ExtraInfo(ByVal Value As LongPtr)
    m_Value.Arr(0).dwExtraInfo = Value
End Property

Public Function ToStr() As String
    Dim s As String: s = TypeName(Me) & "{" & vbCrLf
    With Me
        s = s & "    VirtKeyCode : " & .VirtKeyCode & " (" & MVirtualKeys.EVirtualKeyCodes_ToStr(.VirtKeyCode) & ")" & vbCrLf
        s = s & "    Scan        : " & .Scan & vbCrLf
        s = s & "    Flags       : " & MVirtualKeys.EKeyEventFlags_ToStr(.Flags) & vbCrLf
        s = s & "    Time        : " & .Time & vbCrLf
    End With
    s = s & "}" & vbCrLf
    ToStr = s
End Function

Public Sub Edit()
    MApp.EditInputKeybd Me
End Sub

'Remarks
'INPUT_KEYBOARD supports nonkeyboard-input methods—such as handwriting recognition or voice recognition—as if it were text input by using the KEYEVENTF_UNICODE flag.
'If KEYEVENTF_UNICODE is specified, SendInput sends a WM_KEYDOWN or WM_KEYUP message to the foreground thread's message queue with wParam equal to VK_PACKET.
'Once GetMessage or PeekMessage obtains this message, passing the message to TranslateMessage posts a WM_CHAR message with the Unicode character originally specified by wScan.
'This Unicode character will automatically be converted to the appropriate ANSI value if it is posted to an ANSI window.
'Set the KEYEVENTF_SCANCODE flag to define keyboard input in terms of the scan code.
'This is useful for simulating a physical keystroke regardless of which keyboard is currently being used.
'You can also pass the KEYEVENTF_EXTENDEDKEY flag if the scan code is an extended key.
'The virtual key value of a key can change depending on the current keyboard layout or what other keys were pressed, but the scan code will always be the same.
