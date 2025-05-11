VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
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
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal _
        bVk As Byte, ByVal bScan As Byte, ByVal dwFlags _
        As Long, ByVal dwExtraInfo As Long)

Const VK_LWIN = &H5B
Const VK_APPS = &H5D

Const KEYEVENTF_KEYUP = &H2

Private Sub Command1_Click()
  keybd_event VK_LWIN, 0, 0, 0
  keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
End Sub
