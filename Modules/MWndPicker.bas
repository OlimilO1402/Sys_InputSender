Attribute VB_Name = "MWndPicker"
Option Explicit

Public Type WRect
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type WPoint
    X As Long
    Y As Long
End Type

Public Type WSize
    Width  As Long
    Height As Long
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtextw
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As WRect) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-windowfrompoint
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

'Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal pt As POINTAPI) As LongPtr
'Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

Private m_OldRect As WRect

Public Function WRect_Equals(this As WRect, other As WRect) As Boolean
    With this
        If .Left <> other.Left Then Exit Function
        If .Top <> other.Top Then Exit Function
        If .Right <> other.Right Then Exit Function
        If .Bottom <> other.Bottom Then Exit Function
    End With
    WRect_Equals = True
End Function

Public Function GetCaptionFromMouse(ByVal X As Long, ByVal Y As Long) As String
    
End Function
