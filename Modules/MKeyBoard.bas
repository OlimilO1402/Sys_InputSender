Attribute VB_Name = "MKeyBoard"
Option Explicit

'Private Const WM_KEYDOWN    As Long = &H100

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput
Private Type MOUSEINPUT
    dX          As Long    ' 4
    dY          As Long    ' 4
    MouseData   As Long    ' 4
    dwFlags     As Long    ' 4
    Time        As Long    ' 4
    dwExtraInfo As LongPtr ' 4/8
End Type             ' Sum: 24

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-keybdinput
Private Type KEYBDINPUT
    wVk         As Integer ' 2
    wScan       As Integer ' 2
    dwFlags     As Long    ' 4
    Time        As Long    ' 4
    dwExtraInfo As LongPtr ' 4/8
End Type             ' Sum: 16

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-hardwareinput
Private Type HARDWAREINPUT
    uMsg        As Long    ' 4
    WParamL     As Integer ' 2
    WParamH     As Integer ' 2
End Type              ' Sum: 8

'Public Enum EInputType
'    INPUT_MOUSE
'    INPUT_KEYBOARD
'    INPUT_HARDWARE
'End Enum
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-input
Private Type IINPUT
    dwType As EInputType
    mi     As MOUSEINPUT
    'ki As KEYBDINPUT
    'hi As HARDWAREINPUT
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput
Private Declare Function SendInput Lib "user32" (ByVal cInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getforegroundwindow
'Private Declare Function GetForegroundWindow Lib "user32" () As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As LongPtr) As Long


'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
'Private Declare Function GetFocus Lib "user32" () As Long
'Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub SendKey(ByVal aHWnd As LongPtr, ByVal Key As EVirtualKeyCodes)
    SetForegroundWindow aHWnd
    Dim IIn   As IINPUT
    IIn.dwType = INPUT_KEYBOARD
    Dim KeyIn As KEYBDINPUT
    With KeyIn
        .wVk = Key
        .dwFlags = 2
    End With
    LSet IIn.mi = KeyIn
    Dim liin As Long: liin = LenB(IIn)
    Dim hr As Long: hr = SendInput(1, IIn, liin)
End Sub

Public Sub SendMouse(ByVal aHWnd As LongPtr, ByVal X As Long, ByVal Y As Long)
    SetForegroundWindow aHWnd
    Dim IIn   As IINPUT
    IIn.dwType = INPUT_MOUSE
    Dim MouseIn As MOUSEINPUT
    With MouseIn
        .dX = X
        .dY = Y
    End With
    IIn.mi = MouseIn
    Dim liin As Long: liin = LenB(IIn)
    Dim hr As Long: hr = SendInput(1, IIn, liin)
End Sub

'-------------------------------------------------------
' Objekt (Textfeld,....) welches den Fokus hat ermitteln
Private Function GetControl(ByVal aHWnd As Long) As Long
    'Dim GetCurThreadID As Long, Thread1 As Long: Thread1 = GetWindowThreadProcessId(aHWnd, GetCurThreadID)
    'Dim OtherThreadID  As Long, Thread2 As Long: Thread2 = GetWindowThreadProcessId(GetForegroundWindow, OtherThreadID)
    'Dim bMerker     As Boolean: If Thread1 <> Thread2 Then bMerker = AttachThreadInput(Thread2, Thread1, True)
    'GetControl = GetFocus()
End Function

Private Function GetNextWindow() As Long
    'Dim hNext  As Long 'HWND
    'Dim hFound As Long 'HWND
  
    'hNext = GetWindow(hwnd, GW_HWNDNEXT)
  
'    while (!hFound && hNext)
'    {
'        if (IsWindowVisible(hNext) && !IsIconic(hNext) && !IsCloaked(hNext) && !(GetWindowLong(hNext, GWL_EXSTYLE) & (WS_EX_TOOLWINDOW | WS_EX_TOPMOST)))
'        {
'            hFound = hNext;
'            break;
'        }
'
'        hNext = GetWindow(hNext, GW_HWNDNEXT);
'    }  hNext
'    HWND hFound{};
'
'    hNext = GetWindow(hwnd, GW_HWNDNEXT);
'
'    while (!hFound && hNext)
'    {
'        if (IsWindowVisible(hNext) && !IsIconic(hNext) && !IsCloaked(hNext) && !(GetWindowLong(hNext, GWL_EXSTYLE) & (WS_EX_TOOLWINDOW | WS_EX_TOPMOST)))
'        {
'            hFound = hNext;
'            break;
'        }
'
'        hNext = GetWindow(hNext, GW_HWNDNEXT);
'    }
End Function

'SwitchToThisWindow
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-switchtothiswindow

'void FlashNext(HWND hwnd)
'{
'    HWND hNext{};
'    HWND hFound{};
'
'    hNext = GetWindow(hwnd, GW_HWNDNEXT);
'
'    while (!hFound && hNext)
'    {
'        if (IsWindowVisible(hNext) && !IsIconic(hNext) && !IsCloaked(hNext) && !(GetWindowLong(hNext, GWL_EXSTYLE) & (WS_EX_TOOLWINDOW | WS_EX_TOPMOST)))
'        {
'            hFound = hNext;
'            break;
'        }
'
'        hNext = GetWindow(hNext, GW_HWNDNEXT);
'    }
'
'    if (hFound)
'    {
'        FLASHWINFO fw{ sizeof fw, hFound, FLASHW_TRAY, 1, 0 };
'        FlashWindowEx(&fw);
'        fw.dwFlags = FLASHW_STOP;
'        fw.uCount = 0;
'        FlashWindowEx(&fw);
'    }
'}

