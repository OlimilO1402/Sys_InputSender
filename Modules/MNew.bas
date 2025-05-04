Attribute VB_Name = "MNew"
Option Explicit

Public Function WndPicker(aTimer As Timer, aButton As CommandButton) As WndPicker
    Set WndPicker = New WndPicker: WndPicker.New_ aTimer, aButton
End Function

Public Function WndInputs(ByVal hWndSender As LongPtr, ByVal hWndReceiver As LongPtr) As WndInputs
    Set WndInputs = New WndInputs: WndInputs.New_ hWndSender, hWndReceiver
End Function

Public Function WndInputMouse(ByVal dX As Long, ByVal dY As Long, ByVal MouseData As Long, ByVal Flags As Long, ByVal aTime As Long) As WndInputMouse
    Set WndInputMouse = New WndInputMouse: WndInputMouse.New_ dX, dY, MouseData, Flags, aTime
End Function

Public Function WndInputKeybd(ByVal VirtKey As EVirtualKeyCodes, ByVal Scan As Integer, ByVal Flags As EKeyEventFlags, ByVal aTime As Long) As WndInputKeybd
    Set WndInputKeybd = New WndInputKeybd: WndInputKeybd.New_ CInt(VirtKey), Scan, Flags, aTime
End Function

Public Function WndInputHardw(ByVal aMessage As Long, ByVal WParamL As Integer, ByVal WParamH As Integer) As WndInputHardw
    Set WndInputHardw = New WndInputHardw: WndInputHardw.New_ aMessage, WParamL, WParamH
End Function

