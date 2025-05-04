VERSION 5.00
Begin VB.Form FInputMouse 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit InputMouse"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6510
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
   ScaleHeight     =   3990
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtMouseData 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox TxtYdY 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtXdX 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ListBox LstFlags 
      Height          =   3765
      Left            =   4080
      Style           =   1  'Kontrollkästchen
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LblMusedata 
      AutoSize        =   -1  'True
      Caption         =   "MouseData:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X | dx:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y | dy:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Flags:"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Time (ms):"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   900
   End
End
Attribute VB_Name = "FInputMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Object As WndInputMouse

Private Sub Form_Load()
    MVirtualKeys.EMouseEventFlags_ToList LstFlags
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

