Attribute VB_Name = "MApp"
Option Explicit

Sub Main()
    FKeyboard.Show
End Sub

Sub EditInputKeybd(Obj As WndInputKeybd)
    If FInputKeybd.ShowDialog(Obj) = vbCancel Then Exit Sub
    FKeyboard.UpdateView
End Sub

Sub EditInputMouse(Obj As WndInputMouse)
    If FInputMouse.ShowDialog(Obj) = vbCancel Then Exit Sub
    FKeyboard.UpdateView
End Sub

