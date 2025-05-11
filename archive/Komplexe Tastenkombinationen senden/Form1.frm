VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   """Copmuter suchen"" Dialogfeld per Tastenkombination [WIN] + [STRG] + [F] öffnen"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

' Deklaration
' Auslösen eines Tastatur_Events per Code
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, _
        ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

' Tastendruck absetzen
' keyUp = 0 Up-/Down  | = 1 nur Up | = -1 nur Down
Public Sub SendKeyStroke(ByRef hCode As Byte, Optional ByVal keyUp As Long = 0)
    Const KEYEVENTF_KEYUP = &H2&
    
    'nur KeyUp senden
    If keyUp = 1 Then
        keyUp = 0
    Else
        'KeyDown senden
        Call keybd_event(hCode, 0&, 0&, 0&)
    End If
    
    If keyUp = 0 Then
        'KeyUp senden
        Call keybd_event(hCode, 0&, KEYEVENTF_KEYUP, 0&)
        DoEvents
    End If
End Sub

Private Sub Command1_Click()
    'nur keyDown
    'H5B: Win-Taste
    SendKeyStroke &H5B, -1
    SendKeyStroke vbKeyControl, -1
    SendKeyStroke vbKeyF, -1
    
    'nur keyUp
    SendKeyStroke vbKeyF, 1
    SendKeyStroke vbKeyControl, 1
    SendKeyStroke &H5B, 1
End Sub
