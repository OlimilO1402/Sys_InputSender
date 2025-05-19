VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   5925
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton BtnShowKeyboard 
      Caption         =   "Show Keyboard"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Move mouse over desired window and press the Enter-key (do not click the mouse)"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton BtnWndPicker 
      Caption         =   "Select a Window"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Move mouse over desired window and press the Enter-key (do not click the mouse)"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LblWndTitle 
      AutoSize        =   -1  'True
      Caption         =   "(Window Title)"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mFKeyboard As FKeyboard
Attribute mFKeyboard.VB_VarHelpID = -1

Private Sub BtnShowKeyboard_Click()
    Set mFKeyboard = New FKeyboard
    mFKeyboard.Show
End Sub

Private Sub mFKeyboard_KeyDown(ByVal KeyCode As EVirtualKeyCodes)
    List1.AddItem "KeyDown: " & MVirtualKeys.EVirtualKeyCodes_ToStr(KeyCode)
End Sub

Private Sub mFKeyboard_KeyUp(ByVal KeyCode As EVirtualKeyCodes)
    List1.AddItem "KeyUp: " & MVirtualKeys.EVirtualKeyCodes_ToStr(KeyCode)
End Sub
