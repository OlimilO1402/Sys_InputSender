VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TabStrip bei MouseDown reagieren lassen"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'Kein
      Height          =   855
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lblText2 
         AutoSize        =   -1  'True
         Caption         =   "Nichts besonderes..."
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.Image imgLight 
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":000C
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'Kein
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":044E
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblText1 
         AutoSize        =   -1  'True
         Caption         =   "Was ist wohl in Registerkarte 2..."
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   2325
      End
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registerkarte 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registerkarte 2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c)2002 by Sebastian Schwarz
'mail: deepblue@amargorp.com

Option Explicit

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, _
    ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTUP = &H4

Private Sub tbsMain_Click()
Dim i As Long
    For i = 0 To tbsMain.Tabs.Count - 1
        fraTab(i).Visible = False
    Next i
    fraTab(tbsMain.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub tbsMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Click() Event auslösen
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
End Sub
