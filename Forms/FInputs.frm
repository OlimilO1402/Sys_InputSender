VERSION 5.00
Begin VB.Form FInputs 
   Caption         =   "Edit WndInputs"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
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
   ScaleHeight     =   5295
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      ToolTipText     =   "Delete the list"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TxtToStr 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   9
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "v"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Move down"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton BtnMoveUp 
      Caption         =   "^"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      ToolTipText     =   "Move up"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "-"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "Delete the selected object"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton BtnClone 
      Caption         =   "++"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      ToolTipText     =   "Clone the selected object"
      Top             =   960
      Width           =   495
   End
   Begin VB.ListBox LstWndInputs 
      Height          =   4140
      ItemData        =   "FInputs.frx":0000
      Left            =   0
      List            =   "FInputs.frx":0002
      TabIndex        =   4
      ToolTipText     =   "Select to view; doubleclick to edit"
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton BtnNewInputHardw 
      Caption         =   "+Hardware"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputMouse 
      Caption         =   "+Mouse"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Add New Mouse-Input"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputKeybd 
      Caption         =   "+Keyboard"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Add New Keyboard-Input"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label LblName 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_WInputs As WndInputs

Private Sub Form_Load()
    'Set m_WInputs = MNew.WndInputs(Me.hwnd, GetDesktopWindow)
End Sub

Public Function ShowDialog(Obj As WndInputKeybd) As VbMsgBoxResult
    Set m_Object = Obj.Clone
    UpdateView
    Me.Show vbModal
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    Obj.NewC m_Object
End Function



Private Sub BtnNewInputKeybd_Click():    NewInput FInputKeybd, New WndInputKeybd: End Sub
Private Sub BtnNewInputMouse_Click():    NewInput FInputMouse, New WndInputMouse: End Sub
Private Sub BtnNewInputHardw_Click():    NewInput FInputHardw, New WndInputHardw: End Sub
Private Sub NewInput(FInput, WndInput As Object)
    If FInput.ShowDialog(WndInput) = vbCancel Then Exit Sub
    m_WInputs.Add WndInput
    Me.LstWndInputs.AddItem WndInput.Key
End Sub
Private Sub BtnClone_Click()
    If LstWndInputs.ListCount = 0 Then Exit Sub
    Dim i As Long: i = LstWndInputs.ListIndex
    If i < 0 Then i = LstWndInputs.ListCount - 1 'MsgBox "Select an object first!": Exit Sub
    Dim ii: Set ii = m_WInputs.Item(i)
    m_WInputs.Add ii.Clone
    Me.LstWndInputs.AddItem ii.Key
End Sub
Private Sub BtnDelete_Click()
    If LstWndInputs.ListCount = 0 Then Exit Sub
    Dim i As Long: i = LstWndInputs.ListIndex
    If i < 0 Then MsgBox "Please select an object first!": Exit Sub
    m_WInputs.Remove i
    UpdateView
End Sub
Private Sub BtnMoveUp_Click()
    Dim i As Long: i = LstWndInputs.ListIndex
    If i < 0 Then MsgBox "Select an object first!": Exit Sub
    If i = 0 Then MsgBox "Can not go further up": Exit Sub
    m_WInputs.SwapItems i, i - 1
    UpdateView
    LstWndInputs.ListIndex = i - 1
End Sub
Private Sub BtnMoveDown_Click()
    Dim i As Long: i = LstWndInputs.ListIndex
    If i < 0 Then MsgBox "Select an object first!": Exit Sub
    If i = LstWndInputs.ListCount - 1 Then MsgBox "Can not go further down": Exit Sub
    m_WInputs.SwapItems i, i + 1
    UpdateView
    LstWndInputs.ListIndex = i + 1
End Sub

