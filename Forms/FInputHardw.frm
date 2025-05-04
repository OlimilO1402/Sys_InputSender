VERSION 5.00
Begin VB.Form FInputHardw 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit Input Hardware"
   ClientHeight    =   2175
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4095
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
   ScaleHeight     =   2175
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtMessage 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtWParamL 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox TxtWParamH 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label LblWParamL 
      AutoSize        =   -1  'True
      Caption         =   "WParamL:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   870
   End
   Begin VB.Label LblMessage 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
   Begin VB.Label LbWParamH 
      AutoSize        =   -1  'True
      Caption         =   "WParamH:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   915
   End
End
Attribute VB_Name = "FInputHardw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Object As WndInputHardw

Public Function ShowDialog(Obj As WndInputHardw) As VbMsgBoxResult
    Set m_Object = Obj.Clone
    UpdateView
    Me.Show vbModal
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    Obj.NewC m_Object
End Function

Sub UpdateView()
    With m_Object
        TxtMessage.Text = .Message
        TxtWParamL.Text = .WParamL
        TxtWParamH.Text = .WParamH
    End With
End Sub
Sub UpdateData()
    With m_Object
        .Message = CLng(TxtMessage.Text)
        .WParamL = CLng(TxtWParamL.Text)
        .WParamH = CLng(TxtWParamH.Text)
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

