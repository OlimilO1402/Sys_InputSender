VERSION 5.00
Begin VB.Form FInputKeybd 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit InputKeyboard"
   ClientHeight    =   4095
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
   ScaleHeight     =   4095
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbKCodeHex 
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox CmbKCodeDec 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox LstFlags 
      Height          =   1485
      Left            =   1680
      Style           =   1  'Kontrollkästchen
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox TxtScan 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox CmbKeyCodes 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Hex:"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Decimal:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Time (ms):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Flags:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Scan (Unicode):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VKeyCode:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "FInputKeybd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Object As WndInputKeybd

Private Sub CmbKCodeDec_Click()
    Dim i As Long: i = CmbKCodeDec.ListIndex
    Dim s As String: s = EVirtualKeyCodes_ToStr(i)
    If Len(s) Then
        CmbKCodeHex.ListIndex = i
    Else
        CmbKCodeHex.Text = ""
    End If
    CmbKeyCodes.Text = s
End Sub
Private Sub CmbKCodeDec_LostFocus()
    Dim s As String: s = CmbKCodeDec.Text
    If Not IsNumeric(s) Then
        CmbKCodeDec.Text = ""
        CmbKeyCodes.Text = ""
        Exit Sub
    End If
    Dim i As Long: i = CLng(s)
    s = EVirtualKeyCodes_ToStr(i)
    If Len(s) Then
        CmbKeyCodes.Text = s
        CmbKCodeHex.Text = "0x" & Hex(i)
    End If
End Sub

Private Sub CmbKCodeHex_Click()
    Dim i As Long: i = CmbKCodeHex.ListIndex
    Dim s As String: s = EVirtualKeyCodes_ToStr(i)
    If Len(s) = 0 Then Exit Sub
    CmbKCodeDec.ListIndex = i
    CmbKeyCodes.Text = s
End Sub
Private Sub CmbKCodeHex_LostFocus()
    Dim s As String: s = CmbKCodeHex.Text
    s = Replace(s, "0x", "&H")
Try: On Error GoTo Catch
    Dim i As Long: i = CLng(s)
    s = EVirtualKeyCodes_ToStr(i)
    If Len(s) Then
        CmbKeyCodes.Text = s
        CmbKCodeDec.Text = i
    End If
Catch:
End Sub

Private Sub CmbKeyCodes_Click()
    Dim s As String: s = CmbKeyCodes.Text
    Dim e As EVirtualKeyCodes: e = EVirtualKeyCodes_Parse(s)
    CmbKCodeDec.Text = e: CmbKCodeHex.Text = Hex(e)
End Sub

Private Sub Form_Load()
    FillCmbs
    MVirtualKeys.EVirtualKeyCodes_ToList CmbKeyCodes
    MVirtualKeys.EKeyEventFlags_ToList LstFlags
End Sub
Sub FillCmbs()
    Dim i As Long
    For i = 1 To 255
        CmbKCodeDec.AddItem i
        CmbKCodeHex.AddItem "0x" & Hex(i)
    Next
End Sub
Public Function ShowDialog(Obj As WndInputKeybd) As VbMsgBoxResult
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

Private Sub LstFlags_ItemCheck(Item As Integer)
    Select Case Item
    Case 0: LstFlags.Selected(2) = False
    Case 2: LstFlags.Selected(0) = False
    End Select
End Sub
