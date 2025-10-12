VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Input Actions"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   765
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
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Move down"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton BtnMoveUp 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Move up"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "Delete the selected object"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton BtnClone 
      Caption         =   "++"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Clone the selected object"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Delete the list"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      ToolTipText     =   "Send all inputs"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputHardw 
      Caption         =   "+Hardware"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputMouse 
      Caption         =   "+Mouse"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Add New Mouse-Input"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputKeybd 
      Caption         =   "+Keyboard"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Add New Keyboard-Input"
      Top             =   0
      Width           =   1215
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
      TabIndex        =   10
      Top             =   360
      Width           =   3855
   End
   Begin VB.ListBox LstWndInputs 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      ItemData        =   "FMain.frx":0CCA
      Left            =   0
      List            =   "FMain.frx":0CCC
      TabIndex        =   9
      ToolTipText     =   "Select to view; doubleclick to edit"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

