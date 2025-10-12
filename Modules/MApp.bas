Attribute VB_Name = "MApp"
Option Explicit
Public Declare Function GetDesktopWindow Lib "user32" () As LongPtr

Sub Main()
    'FKeyboard.Show
    FMain.Show
End Sub

Public Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(FileName) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Public Property Get DefaultTmpPFN() As String
    Dim tmpPath As String
    tmpPath = Environ("TMP")
    'tmpPath = App.Path
    DefaultTmpPFN = tmpPath & "\InputSender.bin"
End Property

'Public Property Get DefaultExt() As String
'    DefaultExt = ".bin"
'End Property

Sub EditInputHardw(Obj As WndInputHardw)
    If FInputHardw.ShowDialog(Obj) = vbCancel Then Exit Sub
    FKeyboard.UpdateView
End Sub

Sub EditInputKeybd(Obj As WndInputKeybd)
    If FInputKeybd.ShowDialog(Obj) = vbCancel Then Exit Sub
    FKeyboard.UpdateView
End Sub

Sub EditInputMouse(Obj As WndInputMouse)
    If FInputMouse.ShowDialog(Obj) = vbCancel Then Exit Sub
    FKeyboard.UpdateView
End Sub

