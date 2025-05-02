VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndInputMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput
Private Type MOUSEINPUT
    dwType      As EInputType ' 4
    dX          As Long       ' 4
    dY          As Long       ' 4
    MouseData   As Long       ' 4
    dwFlags     As Long       ' 4
    Time        As Long       ' 4
    dwExtraInfo As LongPtr    ' 4/8
End Type                ' Sum: 28

Private Type PMOUSEINPUT
    SAPtr As TUDTPtr
    Arr() As MOUSEINPUT
End Type

Private m_Value As PMOUSEINPUT
Private m_Data  As MOUSEINPUT

Private Sub Class_Initialize()
    m_Data.dwType = EInputType.INPUT_MOUSE
    MPtr.New_UDTPtr m_Value.SAPtr, FADF_AUTO Or FADF_FIXEDSIZE Or FADF_RECORD, 24
    MPtr.SAPtr(ArrPtr(m_Value.Arr)) = m_Value.SAPtr.pSA
    m_Value.SAPtr.pvData = VarPtr(m_Data)
End Sub

Private Sub Class_Terminate()
    MPtr.ZeroSAPtr ArrPtr(m_Value.Arr)
End Sub
'pData As LongPtr,
Friend Sub New_(dX As Long, dY As Long, MouseData As Long, Flags As Long, aTime As Long)
    If pData Then m_Value.SAPtr.pvData = pData
    With Me: .dX = dX: .dY = dY: .MouseData = MouseData: .Flags = Flags: .Time = aTime: End With
End Sub
'Copy to MNew:
'Public Function WndInputMouse(ByVal pData As LongPtr, ByVal dX As Long, ByVal dY As Long, ByVal MouseData As Long, ByVal Flags As Long, ByVal aTime As Long) As WndInputMouse
'    Set WndInputMouse = New WndInputMouse: WndInputMouse.New_ pData, dX, dY, MouseData, Flags, aTime
'End Function

Friend Sub NewC(other As WndInputMouse)
    With other: Me.New_ .dX, .dY, .MouseData, .Flags, .Time: End With
End Sub
Public Function Clone() As WndInputMouse
    Set Clone = New WndInputMouse: Clone.NewC Me
End Function

Public Property Get Ptr() As LongPtr
    Ptr = m_Value.SAPtr.pvData
End Property
Public Property Let Ptr(ByVal Value As LongPtr)
    m_Value.SAPtr.pvData = Value
End Property

Public Sub SetData()
    m_Value.Arr(0) = m_Data
End Sub

'dx-dy:
'The absolute position of the mouse, or the amount of motion since the last mouse event was generated, depending on the value of the
'dwFlags member. Absolute data is specified as the x coordinate of the mouse; relative data is specified as the number of pixels moved.
Public Property Get dX() As Long
    dX = m_Value.Arr(0).dX
End Property
Public Property Let dX(ByVal Value As Long)
    m_Value.Arr(0).dX = Value
End Property

Public Property Get dY() As Long
    dY = m_Value.Arr(0).dY
End Property
Public Property Let dY(ByVal Value As Long)
    m_Value.Arr(0).dY = Value
End Property

'If dwFlags contains MOUSEEVENTF_WHEEL, then mouseData specifies the amount of wheel movement.
'A positive value indicates that the wheel was rotated forward, away from the user;
'a negative value indicates that the wheel was rotated backward, towards the user.
'One wheel click is defined as WHEEL_DELTA, which is 120.
'If dwFlags does not contain MOUSEEVENTF_WHEEL, MOUSEEVENTF_XDOWN, or MOUSEEVENTF_XUP, then mouseData should be zero.
'If dwFlags contains MOUSEEVENTF_XDOWN or MOUSEEVENTF_XUP, then mouseData specifies which X buttons were pressed or released.
'This value may be any combination of the following flags.
'XBUTTON1 = &H0001  'Set if the first X button is pressed or released.
'XBUTTON2 = &H0002  'Set if the second X button is pressed or released.
Public Property Get MouseData() As Long
    MouseData = m_Value.Arr(0).MouseData
End Property
Public Property Let MouseData(ByVal Value As Long)
    m_Value.Arr(0).MouseData = Value
End Property

'A set of bit flags that specify various aspects of mouse motion and button clicks.
'The bits in this member can be any reasonable combination of the following values. (see above Enum EMouseEventFlags)
'The bit flags that specify mouse button status are set to indicate changes in status, not ongoing conditions.
'For example, if the left mouse button is pressed and held down, MOUSEEVENTF_LEFTDOWN is set when the left button is first
'pressed, but not for subsequent motions. Similarly MOUSEEVENTF_LEFTUP is set only when the button is first released.
'You cannot specify both the MOUSEEVENTF_WHEEL flag and either MOUSEEVENTF_XDOWN or MOUSEEVENTF_XUP flags simultaneously
'in the dwFlags parameter, because they both require use of the mouseData field.
Public Property Get Flags() As EMouseEventFlags
    Flags = m_Value.Arr(0).dwFlags
End Property
Public Property Let Flags(ByVal Value As EMouseEventFlags)
    m_Value.Arr(0).dwFlags = Value
End Property

'The time stamp for the event, in milliseconds. If this parameter is 0, the system will provide its own time stamp.
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

Friend Function ToStr() As String
    Dim s As String: s = TypeName(Me) & "{" & vbCrLf
    With Me
        s = s & "    dX       : " & .dX & vbCrLf
        s = s & "    dY       : " & .dY & vbCrLf
        s = s & "    Mousedata: " & .MouseData & vbCrLf
        s = s & "    dwFlags  : " & .Flags & vbCrLf
        s = s & "    time     : " & .Time & vbCrLf
    End With
    ToStr = s & "}"
End Function

Public Sub Edit()
    MApp.EditInputMouse Me
End Sub

'Remarks
'If the mouse has moved, indicated by MOUSEEVENTF_MOVE, dx and dy specify information about that movement.
'The information is specified as absolute or relative integer values.
'If MOUSEEVENTF_ABSOLUTE value is specified, dx and dy contain normalized absolute coordinates between 0 and 65,535.
'The event procedure maps these coordinates onto the display surface.
'Coordinate (    0,    0) maps onto the upper-left corner of the display surface;
'coordinate (65535,65535) maps onto the lower-right corner. In a multimonitor system, the coordinates map to the primary monitor.
'If MOUSEEVENTF_VIRTUALDESK is specified, the coordinates map to the entire virtual desktop.
'If the MOUSEEVENTF_ABSOLUTE value is not specified, dx and dy specify movement relative to the previous mouse event (the last reported position).
'Positive values mean the mouse moved right (or down); negative values mean the mouse moved left (or up).
'Relative mouse motion is subject to the effects of the mouse speed and the two-mouse threshold values.
'A user sets these three values with the Pointer Speed slider of the Control Panel's Mouse Properties sheet.
'You can obtain and set these values using the SystemParametersInfo function.
'The system applies two tests to the specified relative mouse movement.
'If the specified distance along either the x or y axis is greater than the first  mouse threshold value, and the mouse speed is not zero, the system doubles the distance.
'If the specified distance along either the x or y axis is greater than the second mouse threshold value, and the mouse speed is equal to two, the system doubles the distance
'that resulted from applying the first threshold test.
'It is thus possible for the system to multiply specified relative mouse movement along the x or y axis by up to four times.
