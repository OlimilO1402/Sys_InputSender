VERSION 5.00
Begin VB.Form FInputMouse 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit InputMouse"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Caption         =   "MouseInput-Coordinates"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox TxtXdX 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtYdY 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y | dy:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X | dx:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Screen-Coordinates"
      Height          =   1215
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton BtnGetScreenCoords 
         Caption         =   "Get Screen- Coordinates"
         Height          =   735
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtScrXdX 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtScrYdY 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y | dy:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X | dx:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.TextBox TxtMouseData 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   2  'Zentriert
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ListBox LstFlags 
      Height          =   4050
      Left            =   120
      Style           =   1  'Kontrollkästchen
      TabIndex        =   10
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label LblFlags 
      Caption         =   "Flags:"
      Height          =   2295
      Left            =   3000
      TabIndex        =   15
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label LblMusedata 
      AutoSize        =   -1  'True
      Caption         =   "MouseData:"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Timestamp:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   1920
      Width           =   1245
   End
End
Attribute VB_Name = "FInputMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mWndPicker As WndPicker
Attribute mWndPicker.VB_VarHelpID = -1
Private m_Result As VbMsgBoxResult
Private m_Object As WndInputMouse

Private Sub Form_Load()
    MVirtualKeys.EMouseEventFlags_ToList LstFlags
    Set mWndPicker = MNew.WndPicker(Timer1, BtnGetScreenCoords, False)
End Sub

Public Function ShowDialog(Obj As WndInputMouse) As VbMsgBoxResult
    Set m_Object = Obj.Clone
    UpdateView
    Me.Show vbModal
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    Obj.NewC m_Object
End Function

Sub UpdateView()
    With m_Object
        TxtXdX.Text = .dX
        TxtYdY.Text = .dY
        Dim Scr_X As Long, Scr_Y As Long
        If MVirtualKeys.MouseInpCoords_ToScreenCoords(.dX, .dY, Scr_X, Scr_Y) Then
            TxtScrXdX.Text = Scr_X
            TxtScrYdY.Text = Scr_Y
        End If
        TxtMouseData.Text = .MouseData
        MVirtualKeys.ListBox_EMouseEventFlags(Me.LstFlags) = .Flags
        TxtTime.Text = .Time
    End With
End Sub

Sub UpdateData()
    With m_Object
        .dX = TxtXdX.Text
        .dY = TxtYdY.Text
        .MouseData = TxtMouseData.Text
        .Flags = MVirtualKeys.ListBox_EMouseEventFlags(Me.LstFlags)
        .Time = CLng(TxtTime.Text)
    End With
End Sub

Private Sub BtnOK_Click()
    UpdateData
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

Private Sub LstFlags_Click()
    Dim e As EMouseEventFlags: e = MVirtualKeys.ListBox_EMouseEventFlags(Me.LstFlags)
    LblFlags.Caption = "Flags: " & EMouseEventFlags_ToHex(e) & vbCrLf & MVirtualKeys.EMouseEventFlags_ToStr(e)
End Sub

Private Sub mWndPicker_ScreenCoordinates(ByVal X As Long, Y As Long)
    Dim Mi_X As Long, Mi_Y As Long
    If MVirtualKeys.ScreenCoords_ToMouseInpCoords(X, Y, Mi_X, Mi_Y) Then
        m_Object.dX = Mi_X: m_Object.dY = Mi_Y
        'UpdateView
        With m_Object
            TxtXdX.Text = .dX
            TxtYdY.Text = .dY
            Dim Scr_X As Long, Scr_Y As Long
            If MVirtualKeys.MouseInpCoords_ToScreenCoords(.dX, .dY, Scr_X, Scr_Y) Then
                TxtScrXdX.Text = Scr_X
                TxtScrYdY.Text = Scr_Y
            End If
        End With
    End If
End Sub

Private Sub TxtXdX_LostFocus()
    Dim Mi_X As Long: Mi_X = TxtXdX.Text
    Dim X_pix As Long
    If MVirtualKeys.MouseInpX_ToScreenX(Mi_X, X_pix) Then
        TxtScrXdX.Text = X_pix
    End If
End Sub

Private Sub TxtScrXdX_LostFocus()
    Dim X_pix As Long: X_pix = TxtScrXdX.Text
    Dim Mi_X As Long
    If MVirtualKeys.ScreenX_ToMouseInpX(X_pix, Mi_X) Then
        With m_Object
            .dX = Mi_X
            TxtXdX.Text = .dX
        End With
    End If
End Sub

Private Sub TxtScrYdY_LostFocus()
    Dim Y_pix As Long: Y_pix = TxtScrYdY.Text
    Dim Mi_Y As Long
    If MVirtualKeys.ScreenY_ToMouseInpY(Y_pix, Mi_Y) Then
        With m_Object
            .dY = Mi_Y
            TxtYdY.Text = .dY
        End With
    End If
End Sub
