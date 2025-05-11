VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   945
   ClientLeft      =   3105
   ClientTop       =   2325
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   945
   ScaleWidth      =   2475
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1800
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hier Klicken"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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

Private Declare Function SetCursorPos Lib "user32" (ByVal _
        X As Long, ByVal Y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags _
        As Long, ByVal dx As Long, ByVal dy As Long, ByVal _
        cButtons As Long, ByVal dwExtraInfo As Long)

Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_ABSOLUTE = &H8000&

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim aX%, aY%, dx%, dy%

Private Sub Form_Load()
  Timer1.Enabled = False
  Timer1.Interval = 50
  Me.WindowState = 2
  Command1.Left = Screen.Width / 2
  Command1.Top = Screen.Height / 2
End Sub

Private Sub Command1_Click()
  Timer1.Enabled = True
  dx = Screen.Width / Screen.TwipsPerPixelX - 10
  dy = 5
End Sub

Private Sub Timer1_Timer()
  Dim Pt As POINTAPI
  
    Call GetCursorPos(Pt)
      aX = Pt.X
      aY = Pt.Y
      If aY > dy Then aY = aY - 15
      If aX < dx Then aX = aX + 20
      
      Call SetCursorPos(aX, aY)
      
      If aY <= dy And aX >= dx Then
        SetCursorPos dx, dy
        Timer1.Enabled = False
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
      End If
End Sub
