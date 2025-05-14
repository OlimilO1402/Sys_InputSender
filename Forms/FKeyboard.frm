VERSION 5.00
Begin VB.Form FKeyboard 
   Caption         =   "InputSender Keyboard, Mouse, Hardware"
   ClientHeight    =   4815
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   19695
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
   ScaleHeight     =   4815
   ScaleWidth      =   19695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   18720
      TabIndex        =   166
      ToolTipText     =   "Save to file tmp\InputSender.bin"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnRead 
      Caption         =   "Read"
      Height          =   375
      Left            =   17880
      TabIndex        =   165
      ToolTipText     =   "Read file tmp\InputSender.bin"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   17040
      TabIndex        =   164
      ToolTipText     =   "Delete the list"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "v"
      Height          =   495
      Left            =   15360
      TabIndex        =   144
      ToolTipText     =   "Move down"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BtnMoveUp 
      Caption         =   "^"
      Height          =   495
      Left            =   15360
      TabIndex        =   143
      ToolTipText     =   "Move up"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "-"
      Height          =   495
      Left            =   15360
      TabIndex        =   163
      ToolTipText     =   "Clone the selected object"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton BtnClone 
      Caption         =   "++"
      Height          =   495
      Left            =   15360
      TabIndex        =   142
      ToolTipText     =   "Clone the selected object"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   15840
      TabIndex        =   145
      ToolTipText     =   "Send all inputs"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputHardw 
      Caption         =   "+Hardware"
      Height          =   375
      Left            =   14640
      TabIndex        =   140
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputMouse 
      Caption         =   "+Mouse"
      Height          =   375
      Left            =   13440
      TabIndex        =   139
      ToolTipText     =   "Add New Mouse-Input"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputKeybd 
      Caption         =   "+Keyboard"
      Height          =   375
      Left            =   12240
      TabIndex        =   138
      ToolTipText     =   "Add New Keyboard-Input"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox LstWndInputs 
      Height          =   4140
      ItemData        =   "FKeyboard.frx":0000
      Left            =   12240
      List            =   "FKeyboard.frx":0002
      TabIndex        =   141
      ToolTipText     =   "Select to view; doubleclick to edit"
      Top             =   480
      Width           =   3135
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
      Left            =   15840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   146
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton BtnRec 
      Caption         =   "Record"
      Height          =   375
      Left            =   11040
      TabIndex        =   137
      Top             =   120
      Width           =   1095
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
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox ChkShowF13F24 
      Caption         =   "F13-F24"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   975
   End
   Begin VB.CheckBox ChkShowNumpad 
      Caption         =   "Numpad"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   1095
   End
   Begin VB.CheckBox ChkShowCursorKeys 
      Caption         =   "Cursor Keys"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   1455
   End
   Begin VB.Frame FraSpecialKeys 
      Caption         =   "Special Keys"
      Height          =   1095
      Left            =   9120
      TabIndex        =   23
      Top             =   960
      Width           =   3135
      Begin VB.CommandButton BtnPCSleep 
         Caption         =   "PC Sleep"
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnStartEmail 
         Caption         =   "Start Email"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Media Select"
         Height          =   375
         Index           =   181
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnStartCalculator 
         Caption         =   "Start Calc"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CheckBox ChkShowSpecialKeys 
      Caption         =   "Special Keys"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   1455
   End
   Begin VB.CheckBox ChkShowBrowserKeys 
      Caption         =   "Browser Keys"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   1575
   End
   Begin VB.CheckBox ChkShowMediaKeys 
      Caption         =   "Media Keys"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   1335
   End
   Begin VB.Frame FraBrowserKeys 
      Caption         =   "Browser Keys"
      Height          =   1095
      Left            =   4560
      TabIndex        =   15
      Top             =   960
      Width           =   4575
      Begin VB.CommandButton BtnKey 
         Caption         =   "Favorites"
         Height          =   375
         Index           =   171
         Left            =   3000
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Search"
         Height          =   375
         Index           =   170
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Stop"
         Height          =   375
         Index           =   169
         Left            =   1560
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Home"
         Height          =   375
         Index           =   172
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Forw. >"
         Height          =   375
         Index           =   167
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Refresh"
         Height          =   375
         Index           =   168
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "< Back"
         Height          =   375
         Index           =   166
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FraMediaKeys 
      Caption         =   "Media Keys"
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   4575
      Begin VB.CommandButton BtnKey 
         Caption         =   "Next Track >"
         Height          =   375
         Index           =   176
         Left            =   3000
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Vol. + >"
         Height          =   375
         Index           =   175
         Left            =   2040
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Mute"
         Height          =   375
         Index           =   173
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "< Vol. -"
         Height          =   375
         Index           =   174
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "< Prev Track"
         Height          =   375
         Index           =   177
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Stop"
         Height          =   375
         Index           =   178
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Play/Pause"
         Height          =   375
         Index           =   179
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox PnlKeyboard 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   12255
      TabIndex        =   28
      Top             =   2040
      Width           =   12255
      Begin VB.PictureBox PnlStandardKeys 
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   7335
         TabIndex        =   29
         Top             =   360
         Width           =   7335
         Begin VB.CommandButton BtnKey 
            Caption         =   "Strg"
            Height          =   375
            Index           =   163
            Left            =   6360
            TabIndex        =   131
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Win.r.m.t."
            Height          =   375
            Index           =   92
            Left            =   5400
            TabIndex        =   130
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "AltGr"
            Height          =   375
            Index           =   273
            Left            =   4800
            TabIndex        =   129
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "'------Space------'"
            Height          =   375
            Index           =   32
            Left            =   1920
            TabIndex        =   128
            Top             =   1920
            Width           =   2895
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Alt"
            Height          =   375
            Index           =   18
            Left            =   1320
            TabIndex        =   127
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Win"
            Height          =   375
            Index           =   91
            Left            =   720
            TabIndex        =   126
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Strg"
            Height          =   375
            Index           =   17
            Left            =   0
            TabIndex        =   125
            Top             =   1920
            Width           =   735
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "rShift"
            Height          =   375
            Index           =   272
            Left            =   5880
            TabIndex        =   119
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "-_"
            Height          =   375
            Index           =   189
            Left            =   5400
            TabIndex        =   118
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ":."
            Height          =   375
            Index           =   190
            Left            =   4920
            TabIndex        =   117
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ";,"
            Height          =   375
            Index           =   188
            Left            =   4440
            TabIndex        =   116
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "M"
            Height          =   375
            Index           =   77
            Left            =   3960
            TabIndex        =   115
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "N"
            Height          =   375
            Index           =   78
            Left            =   3480
            TabIndex        =   114
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "B"
            Height          =   375
            Index           =   66
            Left            =   3000
            TabIndex        =   113
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "V"
            Height          =   375
            Index           =   86
            Left            =   2520
            TabIndex        =   112
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "C"
            Height          =   375
            Index           =   67
            Left            =   2040
            TabIndex        =   111
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "X"
            Height          =   375
            Index           =   88
            Left            =   1560
            TabIndex        =   110
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Y"
            Height          =   375
            Index           =   89
            Left            =   1080
            TabIndex        =   109
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<>|"
            Height          =   375
            Index           =   226
            Left            =   600
            TabIndex        =   108
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "lShift"
            Height          =   375
            Index           =   16
            Left            =   0
            TabIndex        =   107
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "'#"
            Height          =   375
            Index           =   191
            Left            =   6120
            TabIndex        =   103
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ä"
            Height          =   375
            Index           =   222
            Left            =   5640
            TabIndex        =   102
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ö"
            Height          =   375
            Index           =   192
            Left            =   5160
            TabIndex        =   101
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "L"
            Height          =   375
            Index           =   76
            Left            =   4680
            TabIndex        =   100
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "K"
            Height          =   375
            Index           =   75
            Left            =   4200
            TabIndex        =   99
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "J"
            Height          =   375
            Index           =   74
            Left            =   3720
            TabIndex        =   98
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "H"
            Height          =   375
            Index           =   72
            Left            =   3240
            TabIndex        =   97
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "G"
            Height          =   375
            Index           =   71
            Left            =   2760
            TabIndex        =   96
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F"
            Height          =   375
            Index           =   70
            Left            =   2280
            TabIndex        =   95
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "D"
            Height          =   375
            Index           =   68
            Left            =   1800
            TabIndex        =   94
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "S"
            Height          =   375
            Index           =   83
            Left            =   1320
            TabIndex        =   93
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "A"
            Height          =   375
            Index           =   65
            Left            =   840
            TabIndex        =   92
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "v-Shift"
            Height          =   375
            Index           =   20
            Left            =   0
            TabIndex        =   91
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<--' Enter"
            Height          =   735
            Index           =   13
            Left            =   6480
            TabIndex        =   83
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "*+~"
            Height          =   375
            Index           =   187
            Left            =   6000
            TabIndex        =   82
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ü"
            Height          =   375
            Index           =   186
            Left            =   5520
            TabIndex        =   81
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "P"
            Height          =   375
            Index           =   80
            Left            =   5040
            TabIndex        =   80
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "O"
            Height          =   375
            Index           =   79
            Left            =   4560
            TabIndex        =   79
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "I"
            Height          =   375
            Index           =   73
            Left            =   4080
            TabIndex        =   78
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "U"
            Height          =   375
            Index           =   85
            Left            =   3600
            TabIndex        =   77
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Z"
            Height          =   375
            Index           =   90
            Left            =   3120
            TabIndex        =   76
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "T"
            Height          =   375
            Index           =   84
            Left            =   2640
            TabIndex        =   75
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "R"
            Height          =   375
            Index           =   82
            Left            =   2160
            TabIndex        =   74
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "E€"
            Height          =   375
            Index           =   69
            Left            =   1680
            TabIndex        =   73
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "W"
            Height          =   375
            Index           =   87
            Left            =   1200
            TabIndex        =   72
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Q@"
            Height          =   375
            Index           =   81
            Left            =   720
            TabIndex        =   71
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "|<-->|"
            Height          =   375
            Index           =   9
            Left            =   0
            TabIndex        =   70
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<- Back"
            Height          =   375
            Index           =   8
            Left            =   6240
            TabIndex        =   62
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "´`"
            Height          =   375
            Index           =   221
            Left            =   5760
            TabIndex        =   61
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "?ß\"
            Height          =   375
            Index           =   219
            Left            =   5280
            TabIndex        =   60
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "=0}"
            Height          =   375
            Index           =   48
            Left            =   4800
            TabIndex        =   59
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ")9]"
            Height          =   375
            Index           =   57
            Left            =   4320
            TabIndex        =   58
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "(8["
            Height          =   375
            Index           =   56
            Left            =   3840
            TabIndex        =   57
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "/7{"
            Height          =   375
            Index           =   55
            Left            =   3360
            TabIndex        =   56
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "&&6"
            Height          =   375
            Index           =   54
            Left            =   2880
            TabIndex        =   55
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "%5"
            Height          =   375
            Index           =   53
            Left            =   2400
            TabIndex        =   54
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "$4"
            Height          =   375
            Index           =   52
            Left            =   1920
            TabIndex        =   53
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "§3³"
            Height          =   375
            Index           =   51
            Left            =   1440
            TabIndex        =   52
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   """2²"
            Height          =   375
            Index           =   50
            Left            =   960
            TabIndex        =   51
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "!1"
            Height          =   375
            Index           =   49
            Left            =   480
            TabIndex        =   50
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "°^"
            Height          =   375
            Index           =   220
            Left            =   0
            TabIndex        =   49
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F12"
            Height          =   375
            Index           =   123
            Left            =   6840
            TabIndex        =   42
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F11"
            Height          =   375
            Index           =   122
            Left            =   6360
            TabIndex        =   41
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F10"
            Height          =   375
            Index           =   121
            Left            =   5880
            TabIndex        =   40
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F9"
            Height          =   375
            Index           =   120
            Left            =   5400
            TabIndex        =   39
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F8"
            Height          =   375
            Index           =   119
            Left            =   4560
            TabIndex        =   38
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F7"
            Height          =   375
            Index           =   118
            Left            =   4080
            TabIndex        =   37
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F6"
            Height          =   375
            Index           =   117
            Left            =   3600
            TabIndex        =   36
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F5"
            Height          =   375
            Index           =   116
            Left            =   3120
            TabIndex        =   35
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F4"
            Height          =   375
            Index           =   115
            Left            =   2280
            TabIndex        =   34
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F3"
            Height          =   375
            Index           =   114
            Left            =   1800
            TabIndex        =   33
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F2"
            Height          =   375
            Index           =   113
            Left            =   1320
            TabIndex        =   32
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F1"
            Height          =   375
            Index           =   112
            Left            =   840
            TabIndex        =   31
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Esc"
            Height          =   375
            Index           =   27
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox PnlF13F24 
         BorderStyle     =   0  'Kein
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   7335
         TabIndex        =   149
         Top             =   0
         Width           =   7335
         Begin VB.CommandButton BtnKey 
            Caption         =   "F24"
            Height          =   375
            Index           =   135
            Left            =   6840
            TabIndex        =   150
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F23"
            Height          =   375
            Index           =   134
            Left            =   6360
            TabIndex        =   151
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F22"
            Height          =   375
            Index           =   133
            Left            =   5880
            TabIndex        =   152
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F21"
            Height          =   375
            Index           =   132
            Left            =   5400
            TabIndex        =   153
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F20"
            Height          =   375
            Index           =   131
            Left            =   4560
            TabIndex        =   154
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F19"
            Height          =   375
            Index           =   130
            Left            =   4080
            TabIndex        =   155
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F18"
            Height          =   375
            Index           =   129
            Left            =   3600
            TabIndex        =   156
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F17"
            Height          =   375
            Index           =   128
            Left            =   3120
            TabIndex        =   157
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F16"
            Height          =   375
            Index           =   127
            Left            =   2280
            TabIndex        =   158
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F15"
            Height          =   375
            Index           =   126
            Left            =   1800
            TabIndex        =   159
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F14"
            Height          =   375
            Index           =   125
            Left            =   1320
            TabIndex        =   160
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F13"
            Height          =   375
            Index           =   124
            Left            =   840
            TabIndex        =   161
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox PnlCursorKeys 
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   7680
         ScaleHeight     =   2295
         ScaleWidth      =   1815
         TabIndex        =   148
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton BtnKey 
            Caption         =   ">"
            Height          =   375
            Index           =   39
            Left            =   1200
            TabIndex        =   134
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "v"
            Height          =   375
            Index           =   40
            Left            =   600
            TabIndex        =   133
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<"
            Height          =   375
            Index           =   37
            Left            =   0
            TabIndex        =   132
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "^"
            Height          =   375
            Index           =   38
            Left            =   600
            TabIndex        =   120
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Bild v"
            Height          =   375
            Index           =   34
            Left            =   1200
            TabIndex        =   86
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ende"
            Height          =   375
            Index           =   35
            Left            =   600
            TabIndex        =   85
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Entf"
            Height          =   375
            Index           =   46
            Left            =   0
            TabIndex        =   84
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Bild^"
            Height          =   375
            Index           =   33
            Left            =   1200
            TabIndex        =   65
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Pos1"
            Height          =   375
            Index           =   36
            Left            =   600
            TabIndex        =   64
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Einfg"
            Height          =   375
            Index           =   45
            Left            =   0
            TabIndex        =   63
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Paus"
            Height          =   375
            Index           =   19
            Left            =   1200
            TabIndex        =   45
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Roll"
            Height          =   375
            Index           =   145
            Left            =   600
            TabIndex        =   44
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Druck"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   42
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox PnlNumpad 
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   9720
         ScaleHeight     =   2295
         ScaleWidth      =   2415
         TabIndex        =   147
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton BtnKey 
            Caption         =   "Enter"
            Height          =   735
            Index           =   43
            Left            =   1800
            TabIndex        =   124
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ",Entf"
            Height          =   375
            Index           =   110
            Left            =   1200
            TabIndex        =   136
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "0 Einfg"
            Height          =   375
            Index           =   96
            Left            =   0
            TabIndex        =   135
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "3 Bv"
            Height          =   375
            Index           =   99
            Left            =   1200
            TabIndex        =   123
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "2 v"
            Height          =   375
            Index           =   98
            Left            =   600
            TabIndex        =   122
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "1 E"
            Height          =   375
            Index           =   97
            Left            =   0
            TabIndex        =   121
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "+"
            Height          =   735
            Index           =   107
            Left            =   1800
            TabIndex        =   90
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "6 >"
            Height          =   375
            Index           =   102
            Left            =   1200
            TabIndex        =   106
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "5"
            Height          =   375
            Index           =   101
            Left            =   600
            TabIndex        =   105
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "4 <"
            Height          =   375
            Index           =   100
            Left            =   0
            TabIndex        =   104
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "9 B^"
            Height          =   375
            Index           =   105
            Left            =   1200
            TabIndex        =   89
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "8 ^"
            Height          =   375
            Index           =   104
            Left            =   600
            TabIndex        =   88
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "7 P1"
            Height          =   375
            Index           =   103
            Left            =   0
            TabIndex        =   87
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "-"
            Height          =   375
            Index           =   109
            Left            =   1800
            TabIndex        =   69
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "*"
            Height          =   375
            Index           =   106
            Left            =   1200
            TabIndex        =   68
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "/"
            Height          =   375
            Index           =   111
            Left            =   600
            TabIndex        =   67
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Num"
            Height          =   375
            Index           =   144
            Left            =   0
            TabIndex        =   66
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CkRoll 
            Caption         =   "Roll"
            Height          =   375
            Left            =   1680
            TabIndex        =   48
            Tag             =   "145"
            Top             =   0
            Width           =   735
         End
         Begin VB.CheckBox CkShift 
            Caption         =   "Shift Lock"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   47
            Tag             =   "20"
            Top             =   0
            Width           =   735
         End
         Begin VB.CheckBox CkNum 
            Caption         =   "Num"
            Height          =   375
            Left            =   0
            TabIndex        =   46
            Tag             =   "144"
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.Label LblWndTitle 
      AutoSize        =   -1  'True
      Caption         =   "(Window Title)"
      Height          =   255
      Left            =   2040
      TabIndex        =   162
      Top             =   120
      Width           =   1245
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "FKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents MWndPicker As WndPicker
Attribute MWndPicker.VB_VarHelpID = -1
Private m_WInputs As WndInputs
Private m_InQueue As WndInputs
Private m_AtOnce  As Boolean 'right-mouse-click started, list in queue until left-mouse-click
Private Declare Function GetDesktopWindow Lib "user32.dll" () As LongPtr

Private Sub Form_Load()
    Me.Caption = "InputSender Keyboard, Mouse, Hardware: v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_WInputs = MNew.WndInputs(Me.hwnd, GetDesktopWindow)
    Set MWndPicker = New WndPicker: MWndPicker.New_ Timer1, BtnWndPicker
    SetBtnKeyTooltip
End Sub

Sub SetBtnKeyTooltip()
    Dim i As Long
    Dim btn As CommandButton
'    For i = BtnKey.LBound To BtnKey.UBound ' BtnKey.Count ' - 1
'        Set btn = BtnKey.Item(i)
'        On Error Resume Next
'        btn.ToolTipText = "VKey: " & i
'    Next
    For Each btn In BtnKey
        i = btn.Index
        btn.ToolTipText = "VKey: " & i & ", 0x" & Hex(i) & ", " & MVirtualKeys.EVirtualKeyCodes_ToStr(i)
    Next
    
    ' Right-Shift is the same VkCode as Left-Shift = 16 , but we can not have twice the same Index
    ' for a CommandButton so I decided to give the button for Right-Shift the index 256 + 16 = 272
    BtnKey(272).ToolTipText = "VKey: " & EVirtualKeyCodes.VK_SHIFT
    
    ' The key AltGr is actually two VkCodes: 17 + 18, so I decided
    ' to give the CommandButton for AltGr the Index 256 + 17 = 273
    BtnKey(273).ToolTipText = "VKey: " & EVirtualKeyCodes.VK_CONTROL & " " & EVirtualKeyCodes.VK_MENU
End Sub

Private Sub BtnClear_Click()
    m_WInputs.Clear
    UpdateView
End Sub

Private Sub BtnRead_Click()
    ReadFile
    UpdateView
End Sub
Private Sub BtnSave_Click()
    SaveFile
End Sub

Sub ReadFile()
    Dim PFN As String: PFN = MApp.DefaultTmpPFN
    If Not FileExists(PFN) Then MsgBox "File not found!" & vbCrLf & PFN: Exit Sub
    If MsgBox("File found, do you want to read it?" & vbCrLf & PFN, vbOKCancel) = vbCancel Then Exit Sub
    m_WInputs.Read PFN
End Sub
Sub SaveFile()
    If m_WInputs.Count = 0 Then MsgBox "List is empty, nothing to save!": Exit Sub
    Dim PFN As String: PFN = MApp.DefaultTmpPFN
    Select Case MsgBox("Do you want to save to the file?" & vbCrLf & PFN, vbOKCancel) ' vbYesNoCancel)
    Case vbCancel: 'Cancel = True: Exit Sub
    'Case vbNo:     Cancel = False: Exit Sub
    'Case vbYes:
    Case vbOK:
        If FileExists(PFN) Then Kill PFN
        m_WInputs.Save PFN
    End Select
End Sub

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
Private Sub BtnSend_Click()
    m_WInputs.Send
End Sub

Public Sub UpdateView()
    Dim i As Long: i = LstWndInputs.ListIndex
    m_WInputs.ToListBox Me.LstWndInputs
    If LstWndInputs.ListCount <= i Then Exit Sub
    LstWndInputs.ListIndex = i
End Sub

Private Sub ChkShowBrowserKeys_Click():    Form_Resize: End Sub
Private Sub ChkShowCursorKeys_Click():     Form_Resize: End Sub
Private Sub ChkShowF13F24_Click():         Form_Resize: End Sub
Private Sub ChkShowMediaKeys_Click():      Form_Resize: End Sub
Private Sub ChkShowNumpad_Click():         Form_Resize: End Sub
Private Sub ChkShowSpecialKeys_Click():    Form_Resize: End Sub

Private Sub LstWndInputs_Click()
    Dim i As Long: i = LstWndInputs.ListIndex
    If i < 0 Then Exit Sub
    Dim Obj: Set Obj = m_WInputs.Item(i)
    TxtToStr.Text = Obj.ToStr
End Sub
Private Sub LstWndInputs_DblClick()
    Dim i As Long: i = LstWndInputs.ListIndex
    If i < 0 Then Exit Sub
    Dim Obj: Set Obj = m_WInputs.Item(i)
    Obj.Edit
End Sub

Private Sub mnuFileNew_Click()
    m_WInputs.Clear
    UpdateView
End Sub

Private Sub mnuFileOpen_Click()
    Dim PFN As String
End Sub

Private Sub mnuFileSave_Click()
    '
End Sub

Private Sub mnuFileSaveAs_Click()
    '
End Sub

Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & App.ProductName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub mWndPicker_Found(ByVal aHWnd As LongPtr, ByVal WndCaption As String)
    'Set m_WInputs = MNew.WndInputs(Me.hwnd, aHWnd)
    m_WInputs.New_ Me.hwnd, aHWnd
    LblWndTitle.Caption = aHWnd & " " & WndCaption
End Sub

Private Sub BtnKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        m_AtOnce = True
        m_InQueue.AddInputKeybd MNew.WndInputKeybd(Index, 0, 0, 0)
    End If
End Sub

Private Sub BtnKey_Click(Index As Integer)
    Dim VKey0 As EVirtualKeyCodes: VKey0 = Index
    Dim VKey1 As EVirtualKeyCodes: VKey1 = -1
    Select Case VKey0
    
    Case EVirtualKeyCodes.VK_CAPITAL: CkShift.Value = IIf(CkShift.Value = vbChecked, vbUnchecked, vbChecked) 'just toggle
    Case EVirtualKeyCodes.VK_NUMLOCK:   CkNum.Value = IIf(CkNum.Value = vbChecked, vbUnchecked, vbChecked)   'just toggle
    Case EVirtualKeyCodes.VK_SCROLL:   CkRoll.Value = IIf(CkRoll.Value = vbChecked, vbUnchecked, vbChecked)  'just toggle
    
    Case 272:         VKey0 = EVirtualKeyCodes.VK_SHIFT
                    ' Right-Shift is the same VkCode as Left-Shift = 16 , but we can not have twice the same Index
                    ' for a CommandButton so I decided to give the button for Right-Shift the index 256 + 16 = 272
    
    Case 273:         VKey0 = EVirtualKeyCodes.VK_CONTROL
                      VKey1 = EVirtualKeyCodes.VK_MENU
                    ' The key AltGr is actually two VkCodes: 17 + 18, so I decided
                    ' to give the CommandButton for AltGr the Index 256 + 17 = 273
    End Select
    'maybe we also need a solution for Num-Lock
    If CkNum.Value = vbUnchecked Then
        Select Case VKey0
        Case EVirtualKeyCodes.VK_NUMPAD0: VKey0 = EVirtualKeyCodes.VK_INSERT
        Case EVirtualKeyCodes.VK_NUMPAD1: VKey0 = EVirtualKeyCodes.VK_END
        Case EVirtualKeyCodes.VK_NUMPAD2: VKey0 = EVirtualKeyCodes.VK_DOWN
        Case EVirtualKeyCodes.VK_NUMPAD3: VKey0 = EVirtualKeyCodes.VK_NEXT
        Case EVirtualKeyCodes.VK_NUMPAD4: VKey0 = EVirtualKeyCodes.VK_LEFT
        ' 5
        Case EVirtualKeyCodes.VK_NUMPAD6: VKey0 = EVirtualKeyCodes.VK_RIGHT
        Case EVirtualKeyCodes.VK_NUMPAD7: VKey0 = EVirtualKeyCodes.VK_HOME
        Case EVirtualKeyCodes.VK_NUMPAD8: VKey0 = EVirtualKeyCodes.VK_UP
        Case EVirtualKeyCodes.VK_NUMPAD9: VKey0 = EVirtualKeyCodes.VK_PRIOR
        Case EVirtualKeyCodes.VK_DECIMAL: VKey0 = EVirtualKeyCodes.VK_DELETE
        End Select
    End If
    
    If m_WInputs Is Nothing Then
        Exit Sub
    End If
    m_WInputs.Clear
    m_WInputs.AddInputKeybd MNew.WndInputKeybd(VKey0, 0, 0, 0)
    m_WInputs.AddInputKeybd MNew.WndInputKeybd(VKey0, 0, EKeyEventFlags.KEYEVENTF_KEYUP, 0)
    If VKey1 > 0 Then
        m_WInputs.AddInputKeybd MNew.WndInputKeybd(VKey1, 0, KEYEVENTF_KEYUP, 0)
    End If
    m_WInputs.Send
End Sub

Private Sub PnlKeyboard_KeyDown(KeyCode As Integer, Shift As Integer)
    BtnKey(KeyCode).Value = True
End Sub

'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Debug.Print "KeyDown KeyCode: " & KeyCode & "; Shift: " & Shift
'End Sub
'Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print "KeyUp   KeyCode: " & KeyCode & "; Shift: " & Shift
'End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    Dim B As Single: B = 2 * 8 * Screen.TwipsPerPixelX
    FraMediaKeys.Visible = ChkShowMediaKeys.Value = vbChecked
    FraBrowserKeys.Visible = ChkShowBrowserKeys.Value = vbChecked
    FraSpecialKeys.Visible = ChkShowSpecialKeys.Value = vbChecked
    PnlF13F24.Visible = ChkShowF13F24.Value = vbChecked
    PnlCursorKeys.Visible = ChkShowCursorKeys.Value = vbChecked
    PnlNumpad.Visible = ChkShowNumpad.Value = vbChecked
    If FraMediaKeys.Visible Then
        T = FraMediaKeys.Top
        W = FraMediaKeys.Width: H = FraMediaKeys.Height
        If W > 0 And H > 0 Then FraMediaKeys.Move l, T, W, H
        l = l + W
    End If
    If FraBrowserKeys.Visible Then
        T = FraBrowserKeys.Top
        W = FraBrowserKeys.Width: H = FraBrowserKeys.Height
        If W > 0 And H > 0 Then FraBrowserKeys.Move l, T, W, H
        l = l + W
    End If
    If FraSpecialKeys.Visible Then
        T = FraSpecialKeys.Top
        W = FraSpecialKeys.Width: H = FraSpecialKeys.Height
        If W > 0 And H > 0 Then FraSpecialKeys.Move l, T, W, H
        l = l + W
    End If
    l = 0: T = IIf(T, T + H, FraMediaKeys.Top)
    
    W = PnlKeyboard.Width
    H = PnlKeyboard.Height
    If W > 0 And H > 0 Then PnlKeyboard.Move l, T, W, H
    T = 0
    l = PnlStandardKeys.Left
    If PnlF13F24.Visible Then
        W = PnlF13F24.Width: H = PnlF13F24.Height
        If W > 0 And H > 0 Then PnlF13F24.Move l, T, W, H
        T = T + H
    End If
    l = PnlStandardKeys.Left
    W = PnlStandardKeys.Width
    H = PnlStandardKeys.Height
    PnlStandardKeys.Move l, T, W, H
    l = PnlStandardKeys.Left + W + B
    PnlStandardKeys.ZOrder 0
    W = 0
    If PnlCursorKeys.Visible Then
        W = PnlCursorKeys.Width
        H = PnlCursorKeys.Height
        If W > 0 And H > 0 Then PnlCursorKeys.Move l, T, W, H
        W = W + B
    End If
    l = l + W '+ b
    If PnlNumpad.Visible Then
        W = PnlNumpad.Width
        H = PnlNumpad.Height
        If W > 0 And H > 0 Then PnlNumpad.Move l, T, W, H
    End If
End Sub


