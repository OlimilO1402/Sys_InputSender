VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type WPoint
    X As Long
    Y As Long
End Type

Private Type WRect
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdesktopwindow
Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdc
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getcursorpos
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As WPoint) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-windowfrompoint
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-childwindowfrompoint
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As WRect) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-drawfocusrect
Private Declare Function DrawFocusRect Lib "user32" (ByVal hhdc As LongPtr, lpRect As Any) As Long 'WinAPIRect) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtextw
Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As LongPtr, ByVal cch As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-releasedc
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long 'Ptr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-releasecapture
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setcapture
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long

Private WithEvents mTimer  As Timer
Attribute mTimer.VB_VarHelpID = -1
Private WithEvents mButton As CommandButton
Attribute mButton.VB_VarHelpID = -1
'Private m_Start As Boolean
Private m_OldWRect As WRect
Private m_Desktop_hWnd As LongPtr
Private m_Desktop_hDC  As LongPtr
Private m_hWnd As LongPtr

Public Event Found(ByVal aHWnd As LongPtr, ByVal WndCaption As String)

Private Sub Class_Initialize()
    m_Desktop_hWnd = GetDesktopWindow
    m_Desktop_hDC = GetDC(m_Desktop_hWnd)
End Sub

Private Sub Class_Terminate()
    Dim hr As Long: hr = ReleaseDC(m_Desktop_hWnd, m_Desktop_hDC)
End Sub

Friend Sub New_(aTimer As Timer, aButton As CommandButton)
    Set mTimer = aTimer: mTimer.Interval = 100: mTimer.Enabled = False
    Set mButton = aButton
End Sub

'Copy To MNew:
'Public Function WndPicker(aTimer As Timer, aButton As CommandButton) As WndPicker
'    Set WndPicker = New WndPicker: WndPicker.New_ aTimer, aButton
'End Function

Private Sub mButton_Click()
    mTimer.Enabled = Not mTimer.Enabled
    SetCapture mButton.hwnd
    If mTimer.Enabled Then Exit Sub
    ReleaseCapture ' mButton.hwnd
    If Not WRect_IsNull(m_OldWRect) Then
        'delete the old rect
        Dim hr As Long: hr = DrawFocusRect(m_Desktop_hDC, m_OldWRect)
    End If
    Dim sCap As String: sCap = String(1024, vbNullChar)
    Dim sLen As Long: sLen = GetWindowTextW(m_hWnd, StrPtr(sCap), Len(sCap))
    sCap = Left(sCap, sLen)
    RaiseEvent Found(m_hWnd, sCap)
End Sub

Private Sub mTimer_Timer()
    Dim pt   As WPoint
    Dim hr   As Long:      hr = GetCursorPos(pt)
    Dim hwnd As LongPtr: hwnd = WindowFromPoint(pt.X, pt.Y)
    If hwnd Then
        Dim hWCh As LongPtr: hWCh = ChildWindowFromPoint(hwnd, pt.X, pt.Y)
        m_hWnd = IIf(hWCh, hWCh, hwnd)
        Dim wr As WRect
        hr = GetWindowRect(m_hWnd, wr)
        If WRect_Equals(wr, m_OldWRect) Then Exit Sub
        If Not WRect_IsNull(m_OldWRect) Then
            'delete the old rect
            hr = DrawFocusRect(m_Desktop_hDC, m_OldWRect)
        End If
        'draw the new rect
        hr = DrawFocusRect(m_Desktop_hDC, wr)
        m_OldWRect = wr
    End If
End Sub

Private Function WRect_Equals(this As WRect, other As WRect) As Boolean
    With this
        If .Top <> other.Top Then Exit Function
        If .Left <> other.Left Then Exit Function
        If .Right <> other.Right Then Exit Function
        If .Bottom <> other.Bottom Then Exit Function
    End With
    WRect_Equals = True
End Function

Private Function WRect_IsNull(this As WRect) As Boolean
    Dim NullRect As WRect: WRect_IsNull = WRect_Equals(this, NullRect)
End Function
