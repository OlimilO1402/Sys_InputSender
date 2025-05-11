VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'zuletzt gepostet von Klaus Langbein im VB.NET-Forum von ActiveVB.de am 24.02.2018 13:45:07
'
' Bespiel für SendInput in VB6/VBA
' Beispiel stammt von AllApi.net
' überarbeitet von
' https://www.mrexcel.com/forum/excel-questions/411552-sendinput-vba.html
'
' Anpassung an VB6 durch K. Langbein, ActiveVB.de Feb.2018
' Um nach dem erstmaligen Start von Notepad, in der gleichen Instanz von Notepad zu schreiben,
' muss vor Verwendung von Sendkey der Fokus auf das entsprechende Fenster gesetzt werden.

'Innerhalb dieser Anwendung geht der Input an die Textbox auf dieser Form

'Benötigt: Form1 , Text1, Command1

Const VK_H As Long = 72
Const VK_E As Long = 69
Const VK_L As Long = 76
Const VK_O As Long = 79
Const KEYEVENTF_KEYUP As Long = &H2
Const INPUT_MOUSE     As Long = 0
Const INPUT_KEYBOARD  As Long = 1
Const INPUT_HARDWARE  As Long = 2

Private Type MOUSEINPUT
    dx          As Long
    dy          As Long
    mouseData   As Long
    dwFlags     As Long
    time        As Long
    dwExtraInfo As Long
End Type

Private Type KEYBDINPUT
    wVk         As Integer
    wScan       As Integer
    dwFlags     As Long
    time        As Long
    dwExtraInfo As Long
End Type

Private Type HARDWAREINPUT
    uMsg        As Long
    wParamL     As Integer
    wParamH     As Integer
End Type

Private Type GENERALINPUT
    dwType      As Long
    xi(0 To 23) As Byte
End Type

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Sub Test()
    
    'Shell "NotePad.EXE", 1 ' hier lieber nicht
    Text1.SetFocus
    DoEvents
    'call the SendKey-function
    SendKey VK_H
    SendKey VK_E
    SendKey VK_L
    SendKey VK_L
    SendKey VK_O
    
End Sub

Private Sub SendKey(bKey As Byte)
    
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    'do the same as above, but for releasing the key
    KInput.wVk = bKey  ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))
    
End Sub

Private Sub Command1_Click()
    Call Test
End Sub
