VERSION 5.00
Begin VB.Form FKeyboard 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Virtual Keyboard"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FKeyboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame FraSpecialKeys 
      Caption         =   "Special Keys"
      Height          =   1095
      Left            =   9120
      TabIndex        =   22
      Top             =   480
      Width           =   3135
      Begin VB.CommandButton BtnPCSleep 
         Caption         =   "PC Sleep"
         Height          =   375
         Left            =   1560
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnStartEmail 
         Caption         =   "Start Email"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Media Select"
         Height          =   375
         Index           =   181
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnStartCalculator 
         Caption         =   "Start Calc"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FraBrowserKeys 
      Caption         =   "Browser Keys"
      Height          =   1095
      Left            =   4560
      TabIndex        =   14
      Top             =   480
      Width           =   4575
      Begin VB.CommandButton BtnKey 
         Caption         =   "Favorites"
         Height          =   375
         Index           =   171
         Left            =   3000
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Search"
         Height          =   375
         Index           =   170
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Stop"
         Height          =   375
         Index           =   169
         Left            =   1560
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Home"
         Height          =   375
         Index           =   172
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Forw. >"
         Height          =   375
         Index           =   167
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Refresh"
         Height          =   375
         Index           =   168
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "< Back"
         Height          =   375
         Index           =   166
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FraMediaKeys 
      Caption         =   "Media Keys"
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   4575
      Begin VB.CommandButton BtnKey 
         Caption         =   "Next Track >"
         Height          =   375
         Index           =   176
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Vol. + >"
         Height          =   375
         Index           =   175
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Mute"
         Height          =   375
         Index           =   173
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "< Vol. -"
         Height          =   375
         Index           =   174
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "< Prev Track"
         Height          =   375
         Index           =   177
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Stop"
         Height          =   375
         Index           =   178
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnKey 
         Caption         =   "Play/Pause"
         Height          =   375
         Index           =   179
         Left            =   120
         TabIndex        =   7
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
      TabIndex        =   27
      Top             =   1680
      Width           =   12255
      Begin VB.PictureBox PnlStandardKeys 
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   7335
         TabIndex        =   28
         Top             =   360
         Width           =   7335
         Begin VB.CommandButton BtnKey 
            Caption         =   "Strg"
            Height          =   375
            Index           =   163
            Left            =   6360
            TabIndex        =   130
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Win.r.m.t."
            Height          =   375
            Index           =   92
            Left            =   5400
            TabIndex        =   129
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "AltGr"
            Height          =   375
            Index           =   273
            Left            =   4800
            TabIndex        =   128
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "'------Space------'"
            Height          =   375
            Index           =   32
            Left            =   1920
            TabIndex        =   127
            Top             =   1920
            Width           =   2895
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Alt"
            Height          =   375
            Index           =   18
            Left            =   1320
            TabIndex        =   126
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Win"
            Height          =   375
            Index           =   91
            Left            =   720
            TabIndex        =   125
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Strg"
            Height          =   375
            Index           =   17
            Left            =   0
            TabIndex        =   124
            Top             =   1920
            Width           =   735
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "rShift"
            Height          =   375
            Index           =   272
            Left            =   5880
            TabIndex        =   118
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "-_"
            Height          =   375
            Index           =   189
            Left            =   5400
            TabIndex        =   117
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ":."
            Height          =   375
            Index           =   190
            Left            =   4920
            TabIndex        =   116
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ";,"
            Height          =   375
            Index           =   188
            Left            =   4440
            TabIndex        =   115
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "M"
            Height          =   375
            Index           =   77
            Left            =   3960
            TabIndex        =   114
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "N"
            Height          =   375
            Index           =   78
            Left            =   3480
            TabIndex        =   113
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "B"
            Height          =   375
            Index           =   66
            Left            =   3000
            TabIndex        =   112
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "V"
            Height          =   375
            Index           =   86
            Left            =   2520
            TabIndex        =   111
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "C"
            Height          =   375
            Index           =   67
            Left            =   2040
            TabIndex        =   110
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "X"
            Height          =   375
            Index           =   88
            Left            =   1560
            TabIndex        =   109
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Y"
            Height          =   375
            Index           =   89
            Left            =   1080
            TabIndex        =   108
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<>|"
            Height          =   375
            Index           =   226
            Left            =   600
            TabIndex        =   107
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "lShift"
            Height          =   375
            Index           =   16
            Left            =   0
            TabIndex        =   106
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "'#"
            Height          =   375
            Index           =   191
            Left            =   6120
            TabIndex        =   102
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ä"
            Height          =   375
            Index           =   222
            Left            =   5640
            TabIndex        =   101
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ö"
            Height          =   375
            Index           =   192
            Left            =   5160
            TabIndex        =   100
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "L"
            Height          =   375
            Index           =   76
            Left            =   4680
            TabIndex        =   99
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "K"
            Height          =   375
            Index           =   75
            Left            =   4200
            TabIndex        =   98
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "J"
            Height          =   375
            Index           =   74
            Left            =   3720
            TabIndex        =   97
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "H"
            Height          =   375
            Index           =   72
            Left            =   3240
            TabIndex        =   96
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "G"
            Height          =   375
            Index           =   71
            Left            =   2760
            TabIndex        =   95
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F"
            Height          =   375
            Index           =   70
            Left            =   2280
            TabIndex        =   94
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "D"
            Height          =   375
            Index           =   68
            Left            =   1800
            TabIndex        =   93
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "S"
            Height          =   375
            Index           =   83
            Left            =   1320
            TabIndex        =   92
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "A"
            Height          =   375
            Index           =   65
            Left            =   840
            TabIndex        =   91
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "v-Shift"
            Height          =   375
            Index           =   20
            Left            =   0
            TabIndex        =   90
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<--' Enter"
            Height          =   735
            Index           =   13
            Left            =   6480
            TabIndex        =   82
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "*+~"
            Height          =   375
            Index           =   187
            Left            =   6000
            TabIndex        =   81
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ü"
            Height          =   375
            Index           =   186
            Left            =   5520
            TabIndex        =   80
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "P"
            Height          =   375
            Index           =   80
            Left            =   5040
            TabIndex        =   79
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "O"
            Height          =   375
            Index           =   79
            Left            =   4560
            TabIndex        =   78
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "I"
            Height          =   375
            Index           =   73
            Left            =   4080
            TabIndex        =   77
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "U"
            Height          =   375
            Index           =   85
            Left            =   3600
            TabIndex        =   76
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Z"
            Height          =   375
            Index           =   90
            Left            =   3120
            TabIndex        =   75
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "T"
            Height          =   375
            Index           =   84
            Left            =   2640
            TabIndex        =   74
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "R"
            Height          =   375
            Index           =   82
            Left            =   2160
            TabIndex        =   73
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "E€"
            Height          =   375
            Index           =   69
            Left            =   1680
            TabIndex        =   72
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "W"
            Height          =   375
            Index           =   87
            Left            =   1200
            TabIndex        =   71
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Q@"
            Height          =   375
            Index           =   81
            Left            =   720
            TabIndex        =   70
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "|<-->|"
            Height          =   375
            Index           =   9
            Left            =   0
            TabIndex        =   69
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<- Back"
            Height          =   375
            Index           =   8
            Left            =   6240
            TabIndex        =   61
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "´`"
            Height          =   375
            Index           =   221
            Left            =   5760
            TabIndex        =   60
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "?ß\"
            Height          =   375
            Index           =   219
            Left            =   5280
            TabIndex        =   59
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "=0}"
            Height          =   375
            Index           =   48
            Left            =   4800
            TabIndex        =   58
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ")9]"
            Height          =   375
            Index           =   57
            Left            =   4320
            TabIndex        =   57
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "(8["
            Height          =   375
            Index           =   56
            Left            =   3840
            TabIndex        =   56
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "/7{"
            Height          =   375
            Index           =   55
            Left            =   3360
            TabIndex        =   55
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "&&6"
            Height          =   375
            Index           =   54
            Left            =   2880
            TabIndex        =   54
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "%5"
            Height          =   375
            Index           =   53
            Left            =   2400
            TabIndex        =   53
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "$4"
            Height          =   375
            Index           =   52
            Left            =   1920
            TabIndex        =   52
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "§3³"
            Height          =   375
            Index           =   51
            Left            =   1440
            TabIndex        =   51
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   """2²"
            Height          =   375
            Index           =   50
            Left            =   960
            TabIndex        =   50
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "!1"
            Height          =   375
            Index           =   49
            Left            =   480
            TabIndex        =   49
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "°^"
            Height          =   375
            Index           =   220
            Left            =   0
            TabIndex        =   48
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F12"
            Height          =   375
            Index           =   123
            Left            =   6840
            TabIndex        =   41
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F11"
            Height          =   375
            Index           =   122
            Left            =   6360
            TabIndex        =   40
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F10"
            Height          =   375
            Index           =   121
            Left            =   5880
            TabIndex        =   39
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F9"
            Height          =   375
            Index           =   120
            Left            =   5400
            TabIndex        =   38
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F8"
            Height          =   375
            Index           =   119
            Left            =   4560
            TabIndex        =   37
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F7"
            Height          =   375
            Index           =   118
            Left            =   4080
            TabIndex        =   36
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F6"
            Height          =   375
            Index           =   117
            Left            =   3600
            TabIndex        =   35
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F5"
            Height          =   375
            Index           =   116
            Left            =   3120
            TabIndex        =   34
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F4"
            Height          =   375
            Index           =   115
            Left            =   2280
            TabIndex        =   33
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F3"
            Height          =   375
            Index           =   114
            Left            =   1800
            TabIndex        =   32
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F2"
            Height          =   375
            Index           =   113
            Left            =   1320
            TabIndex        =   31
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F1"
            Height          =   375
            Index           =   112
            Left            =   840
            TabIndex        =   30
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Esc"
            Height          =   375
            Index           =   27
            Left            =   0
            TabIndex        =   29
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
         TabIndex        =   138
         Top             =   0
         Width           =   7335
         Begin VB.CommandButton BtnKey 
            Caption         =   "F24"
            Height          =   375
            Index           =   135
            Left            =   6840
            TabIndex        =   139
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F23"
            Height          =   375
            Index           =   134
            Left            =   6360
            TabIndex        =   140
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F22"
            Height          =   375
            Index           =   133
            Left            =   5880
            TabIndex        =   141
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F21"
            Height          =   375
            Index           =   132
            Left            =   5400
            TabIndex        =   142
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F20"
            Height          =   375
            Index           =   131
            Left            =   4560
            TabIndex        =   143
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F19"
            Height          =   375
            Index           =   130
            Left            =   4080
            TabIndex        =   144
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F18"
            Height          =   375
            Index           =   129
            Left            =   3600
            TabIndex        =   145
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F17"
            Height          =   375
            Index           =   128
            Left            =   3120
            TabIndex        =   146
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F16"
            Height          =   375
            Index           =   127
            Left            =   2280
            TabIndex        =   147
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F15"
            Height          =   375
            Index           =   126
            Left            =   1800
            TabIndex        =   148
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F14"
            Height          =   375
            Index           =   125
            Left            =   1320
            TabIndex        =   149
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "F13"
            Height          =   375
            Index           =   124
            Left            =   840
            TabIndex        =   150
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
         TabIndex        =   137
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton BtnKey 
            Caption         =   ">"
            Height          =   375
            Index           =   39
            Left            =   1200
            TabIndex        =   133
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "v"
            Height          =   375
            Index           =   40
            Left            =   600
            TabIndex        =   132
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "<"
            Height          =   375
            Index           =   37
            Left            =   0
            TabIndex        =   131
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "^"
            Height          =   375
            Index           =   38
            Left            =   600
            TabIndex        =   119
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Bild v"
            Height          =   375
            Index           =   34
            Left            =   1200
            TabIndex        =   85
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Ende"
            Height          =   375
            Index           =   35
            Left            =   600
            TabIndex        =   84
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Entf"
            Height          =   375
            Index           =   46
            Left            =   0
            TabIndex        =   83
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Bild^"
            Height          =   375
            Index           =   33
            Left            =   1200
            TabIndex        =   64
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Pos1"
            Height          =   375
            Index           =   36
            Left            =   600
            TabIndex        =   63
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Einfg"
            Height          =   375
            Index           =   45
            Left            =   0
            TabIndex        =   62
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Paus"
            Height          =   375
            Index           =   19
            Left            =   1200
            TabIndex        =   44
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Roll"
            Height          =   375
            Index           =   145
            Left            =   600
            TabIndex        =   43
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
            TabIndex        =   42
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
         TabIndex        =   136
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton BtnKey 
            Caption         =   "Enter"
            Height          =   735
            Index           =   43
            Left            =   1800
            TabIndex        =   123
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   ",Entf"
            Height          =   375
            Index           =   110
            Left            =   1200
            TabIndex        =   135
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "0 Einfg"
            Height          =   375
            Index           =   96
            Left            =   0
            TabIndex        =   134
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "3 Bv"
            Height          =   375
            Index           =   99
            Left            =   1200
            TabIndex        =   122
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "2 v"
            Height          =   375
            Index           =   98
            Left            =   600
            TabIndex        =   121
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "1 E"
            Height          =   375
            Index           =   97
            Left            =   0
            TabIndex        =   120
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "+"
            Height          =   735
            Index           =   107
            Left            =   1800
            TabIndex        =   89
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "6 >"
            Height          =   375
            Index           =   102
            Left            =   1200
            TabIndex        =   105
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "5"
            Height          =   375
            Index           =   101
            Left            =   600
            TabIndex        =   104
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "4 <"
            Height          =   375
            Index           =   100
            Left            =   0
            TabIndex        =   103
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "9 B^"
            Height          =   375
            Index           =   105
            Left            =   1200
            TabIndex        =   88
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "8 ^"
            Height          =   375
            Index           =   104
            Left            =   600
            TabIndex        =   87
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "7 P1"
            Height          =   375
            Index           =   103
            Left            =   0
            TabIndex        =   86
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "-"
            Height          =   375
            Index           =   109
            Left            =   1800
            TabIndex        =   68
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "*"
            Height          =   375
            Index           =   106
            Left            =   1200
            TabIndex        =   67
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "/"
            Height          =   375
            Index           =   111
            Left            =   600
            TabIndex        =   66
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton BtnKey 
            Caption         =   "Num"
            Height          =   375
            Index           =   144
            Left            =   0
            TabIndex        =   65
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CkRoll 
            Caption         =   "Roll"
            Height          =   375
            Left            =   1680
            TabIndex        =   47
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
            TabIndex        =   46
            Tag             =   "20"
            Top             =   0
            Width           =   735
         End
         Begin VB.CheckBox CkNum 
            Caption         =   "Num"
            Height          =   375
            Left            =   0
            TabIndex        =   45
            Tag             =   "144"
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.CheckBox ChkShowF13F24 
      Caption         =   "F13-F24"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   975
   End
   Begin VB.CheckBox ChkShowNumpad 
      Caption         =   "Numpad"
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1095
   End
   Begin VB.CheckBox ChkShowCursorKeys 
      Caption         =   "Cursor Keys"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1455
   End
   Begin VB.CheckBox ChkShowSpecialKeys 
      Caption         =   "Special Keys"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1455
   End
   Begin VB.CheckBox ChkShowBrowserKeys 
      Caption         =   "Browser Keys"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1575
   End
   Begin VB.CheckBox ChkShowMediaKeys 
      Caption         =   "Media Keys"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1335
   End
End
Attribute VB_Name = "FKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event KeyDown(ByVal KeyCode As EVirtualKeyCodes)
Public Event KeyUp(ByVal KeyCode As EVirtualKeyCodes)
Public m_KeyDownList As Collection

'Private Declare Function GetDesktopWindow Lib "user32.dll" () As LongPtr

Private Sub Form_Load()
'    Me.Caption = "InputSender Keyboard, Mouse, Hardware: v" & App.Major & "." & App.Minor & "." & App.Revision
'    Set m_WInputs = MNew.WndInputs(Me.hWnd, GetDesktopWindow)
'    Set MWndPicker = New WndPicker: MWndPicker.New_ Timer1, BtnWndPicker
    Set m_KeyDownList = New Collection
    SetBtnKeyTooltip
End Sub

Sub SetBtnKeyTooltip()
    Dim i As Long
    Dim btn As CommandButton
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

'Public Sub UpdateView()
'    Dim i As Long: i = LstWndInputs.ListIndex
'    m_WInputs.ToListBox Me.LstWndInputs
'    If LstWndInputs.ListCount <= i Then Exit Sub
'    LstWndInputs.ListIndex = i
'End Sub

Private Sub ChkShowBrowserKeys_Click():    Form_Resize: End Sub
Private Sub ChkShowCursorKeys_Click():     Form_Resize: End Sub
Private Sub ChkShowF13F24_Click():         Form_Resize: End Sub
Private Sub ChkShowMediaKeys_Click():      Form_Resize: End Sub
Private Sub ChkShowNumpad_Click():         Form_Resize: End Sub
Private Sub ChkShowSpecialKeys_Click():    Form_Resize: End Sub

'Private Sub mnuHelpInfo_Click()
'    MsgBox App.CompanyName & " " & App.ProductName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
'End Sub

'Private Sub mWndPicker_Found(ByVal aHWnd As LongPtr, ByVal WndCaption As String)
'    'Set m_WInputs = MNew.WndInputs(Me.hwnd, aHWnd)
'    m_WInputs.New_ Me.hWnd, aHWnd
'    LblWndTitle.Caption = aHWnd & " " & WndCaption
'End Sub

Private Sub BtnKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        Dim VKey0 As EVirtualKeyCodes
        Dim VKey1 As EVirtualKeyCodes
        IndexToKeyCodes Index, VKey0, VKey1
        m_KeyDownList.Add VKey0
        RaiseEvent KeyDown(VKey0)
        If VKey1 > 0 Then
            m_KeyDownList.Add VKey1
            RaiseEvent KeyDown(VKey1)
        End If
    End If
End Sub

Private Sub IndexToCheckBoxes(ByVal Index As Long)
    Select Case Index
    Case EVirtualKeyCodes.VK_CAPITAL: CkShift.Value = IIf(CkShift.Value = vbChecked, vbUnchecked, vbChecked) 'just toggle
    Case EVirtualKeyCodes.VK_NUMLOCK:   CkNum.Value = IIf(CkNum.Value = vbChecked, vbUnchecked, vbChecked)   'just toggle
    Case EVirtualKeyCodes.VK_SCROLL:   CkRoll.Value = IIf(CkRoll.Value = vbChecked, vbUnchecked, vbChecked)  'just toggle
    End Select
End Sub
Private Function IndexToKeyCodes(ByVal Index As Long, ByRef VKey0_out As Long, Optional VKey1_out As Long = -1)
    Select Case Index
    Case 272: VKey0_out = EVirtualKeyCodes.VK_SHIFT
              ' Right-Shift is the same VkCode as Left-Shift = 16 , but we can not have twice the same Index
              ' for a CommandButton so I decided to give the button for Right-Shift the index 256 + 16 = 272
    
    Case 273: VKey0_out = EVirtualKeyCodes.VK_CONTROL
              VKey1_out = EVirtualKeyCodes.VK_MENU
              ' The key AltGr is actually two VkCodes: 17 + 18, so I decided
              ' to give the CommandButton for AltGr the Index 256 + 17 = 273
    Case Else
              VKey0_out = Index
              'And we also need a solution for Num-Lock
              If CkNum.Value = vbUnchecked Then
                  Select Case Index
                  Case EVirtualKeyCodes.VK_NUMPAD0: VKey0_out = EVirtualKeyCodes.VK_INSERT
                  Case EVirtualKeyCodes.VK_NUMPAD1: VKey0_out = EVirtualKeyCodes.VK_END
                  Case EVirtualKeyCodes.VK_NUMPAD2: VKey0_out = EVirtualKeyCodes.VK_DOWN
                  Case EVirtualKeyCodes.VK_NUMPAD3: VKey0_out = EVirtualKeyCodes.VK_NEXT
                  Case EVirtualKeyCodes.VK_NUMPAD4: VKey0_out = EVirtualKeyCodes.VK_LEFT
                  ' 5
                  Case EVirtualKeyCodes.VK_NUMPAD6: VKey0_out = EVirtualKeyCodes.VK_RIGHT
                  Case EVirtualKeyCodes.VK_NUMPAD7: VKey0_out = EVirtualKeyCodes.VK_HOME
                  Case EVirtualKeyCodes.VK_NUMPAD8: VKey0_out = EVirtualKeyCodes.VK_UP
                  Case EVirtualKeyCodes.VK_NUMPAD9: VKey0_out = EVirtualKeyCodes.VK_PRIOR
                  Case EVirtualKeyCodes.VK_DECIMAL: VKey0_out = EVirtualKeyCodes.VK_DELETE
                  End Select
              End If
    End Select
End Function

Private Sub BtnKey_Click(Index As Integer)
    IndexToCheckBoxes Index
    RaiseEvent KeyDown(Index)
    RaiseEvent KeyUp(Index)

    Dim i As Long, VKey As EVirtualKeyCodes
    If m_KeyDownList.Count > 0 Then
        For i = m_KeyDownList.Count To 1 Step -1
            VKey = m_KeyDownList.Item(i)
            RaiseEvent KeyUp(VKey)
            m_KeyDownList.Remove i
        Next
    End If
End Sub

Private Sub PnlKeyboard_KeyDown(KeyCode As Integer, Shift As Integer)
    BtnKey(KeyCode).Value = True
End Sub

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


