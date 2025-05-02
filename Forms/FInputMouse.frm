VERSION 5.00
Begin VB.Form FInputMouse 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit InputMouse"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4815
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
   ScaleHeight     =   3615
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox CmbKeyCodes 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox TxtScan 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   3015
   End
   Begin VB.ListBox LstFlags 
      Height          =   1485
      Left            =   1680
      Style           =   1  'Kontrollkästchen
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VKeyCode:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Scan (Unicode):"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Flags:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Time (ms):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
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
    MVirtualKeys.EVirtualKeyCodes_ToList CmbKeyCodes
    MVirtualKeys.EKeyEventFlags_ToList LstFlags
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
        CmbKeyCodes.Text = MVirtualKeys.EVirtualKeyCodes_ToStr(.VirtKeyCode)
        TxtScan.Text = "&H" & Hex(.Scan)
        MVirtualKeys.ListBox_EKeyEventFlags(Me.LstFlags) = .Flags
        TxtTime.Text = .Time
    End With
End Sub
Sub UpdateData()
    With m_Object
        .VirtKeyCode = MVirtualKeys.EVirtualKeyCodes_Parse(CmbKeyCodes.Text)
        .Scan = CLng(TxtScan.Text)
        .Flags = MVirtualKeys.ListBox_EKeyEventFlags(Me.LstFlags)
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

