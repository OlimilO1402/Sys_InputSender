VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput
'Private Type MOUSEINPUT
'    dwType      As EInputType
'    dX          As Long    ' 4
'    dY          As Long    ' 4
'    MouseData   As Long    ' 4
'    dwFlags     As Long    ' 4
'    Time        As Long    ' 4
'    dwExtraInfo As LongPtr ' 4/8
'End Type             ' Sum: 24
'
''https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-keybdinput
'Private Type KEYBDINPUT
'    dwType      As EInputType
'    wVk         As Integer ' 2
'    wScan       As Integer ' 2
'    dwFlags     As Long    ' 4
'    Time        As Long    ' 4
'    dwExtraInfo As LongPtr ' 4/8
'End Type             ' Sum: 16
'
''https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-hardwareinput
'Private Type HARDWAREINPUT
'    dwType      As EInputType
'    uMsg        As Long    ' 4
'    WParamL     As Integer ' 2
'    WParamH     As Integer ' 2
'End Type              ' Sum: 8
'
Public Enum EInputType
    INPUT_MOUSE
    INPUT_KEYBOARD
    INPUT_HARDWARE
End Enum
'Private Enum LongPtr
'    [Value]
'End Enum
'typedef struct tagINPUT {
'  DWORD type;
'  union {
'    MOUSEINPUT    mi;
'    KEYBDINPUT    ki;
'    HARDWAREINPUT hi;
'  } DUMMYUNIONNAME;
'} INPUT, *PINPUT, *LPINPUT;
Private Type HKMINPUT
    dwType As EInputType
    Value1 As Long
    Value2 As Long
    Value3 As Long
    Value4 As Long
    Value5 As Long
    Value6 As LongPtr
End Type
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-input
'Private Type IINPUT
'    INPUT  As HKMINPUT
'    'ki As KEYBDINPUT
'    'hi As HARDWAREINPUT
'End Type

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-lastinputinfo
'typedef struct tagLASTINPUTINFO {
'  UINT  cbSize;
'  DWORD dwTime;
'} LASTINPUTINFO, *PLASTINPUTINFO;
Private Type LASTINPUTINFO
    cbSize As Long '8 ' The size of the structure, in bytes. This member must be set to sizeof(LASTINPUTINFO)
    dwTime As Long 'The tick count when the last input event was received.
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput
Private Declare Function SendInput Lib "user32" (ByVal cInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getlastinputinfo
'BOOL GetLastInputInfo([out] PLASTINPUTINFO plii);
Private Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long
'Remarks: This function is useful for input idle detection. For more information on tick counts, see GetTickCount.

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
Private m_hWndSender   As LongPtr 'hwnd of window FMain in this app
Private m_hWndReceiver As LongPtr 'hwnd of any window to send inputs to
Private m_Date  As Date
Private m_Count As Long
Private m_Arr() As HKMINPUT 'IINPUT

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Friend Sub New_(hWndSender As LongPtr, hWndReceiver As LongPtr)
 m_hWndSender = hWndSender:          m_hWndReceiver = hWndReceiver
End Sub

'Copy to MNew:
'Public Function WndInputs(ByVal hWndSender As LongPtr, ByVal hWndReceiver As LongPtr) As WndInputs
'    Set WndInputs = New WndInputs: WndInputs.New_ hWndSender, ByVal hWndReceiver
'End Function
Public Sub Clear()
    Class_Initialize
    m_Count = 0
End Sub
Public Property Get Count() As Long
    Count = m_Count
End Property

Function GetNextCapacity(ByVal Count As Long) As Long
    Dim sq As Double: sq = VBA.Math.Sqr(Count)
    Dim isq As Long: isq = Int(sq)
    isq = isq + IIf(isq <= sq, 1, 0)
    GetNextCapacity = 2 ^ isq
End Function

Private Sub EnsureCapacity(ByVal Count As Long)
    If UBound(m_Arr) < Count Then
        Dim ucap As Long: ucap = GetNextCapacity(Count) - 1
        ReDim Preserve m_Arr(0 To ucap)
    End If
End Sub

Public Sub Add(Obj)
    EnsureCapacity m_Count + 1
    Obj.Ptr = VarPtr(m_Arr(m_Count)): Obj.SetData
    m_Count = m_Count + 1
End Sub

'Public Function AddInputMouse(Obj As WndInputMouse) As WndInputMouse
'    Set AddInputMouse = Obj
'    Add Obj
'    'EnsureCapacity
'    'Obj.Ptr = VarPtr(m_Arr(m_Count)): Obj.SetData
'    'm_Count = m_Count + 1
'End Function
'
'Public Function AddInputKeybd(Obj As WndInputKeybd) As WndInputKeybd
'    Set AddInputKeybd = Obj
'    Add Obj
'    'EnsureCapacity
'    'Obj.Ptr = VarPtr(m_Arr(m_Count)): Obj.SetData
'    'm_Count = m_Count + 1
'End Function
'
Public Property Get InputType(ByVal Index As Long) As EInputType
    If Index < 0 Or m_Count <= Index Then Exit Property
    InputType = m_Arr(Index).dwType
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = m_Count = 0
End Property

Public Sub SwapItems(ByVal Index1 As Long, ByVal Index2 As Long)
    If Index1 < 0 Or m_Count <= Index1 Then Exit Sub
    Dim tmpI1 As HKMINPUT: tmpI1 = m_Arr(Index1)
    m_Arr(Index1) = m_Arr(Index2): m_Arr(Index2) = tmpI1
End Sub

Public Property Get Item(ByVal Index As Long) As Object
    If Index < 0 Or m_Count <= Index Then Exit Property
    Dim it As EInputType: it = Me.InputType(Index)
    Select Case it
    Case EInputType.INPUT_MOUSE:    Set Item = New WndInputMouse
    Case EInputType.INPUT_KEYBOARD: Set Item = New WndInputKeybd
    Case EInputType.INPUT_HARDWARE: Set Item = New WndInputHardw
    End Select
    Item.Ptr = VarPtr(m_Arr(Index))
End Property

'Public Property Get ItemIM(ByVal Index As Long) As WndInputMouse
'    If Index < 0 Or m_Count <= Index Then Exit Property
'    If Me.InputType(Index) <> INPUT_MOUSE Then Exit Property
'    Set ItemIM = New WndInputMouse: ItemIM.Ptr = VarPtr(m_Arr(Index))
'End Property
'
'Public Property Get ItemIK(ByVal Index As Long) As WndInputKeybd
'    If Index < 0 Or m_Count <= Index Then Exit Property
'    If Me.InputType(Index) <> INPUT_KEYBOARD Then Exit Property
'    Set ItemIK = New WndInputKeybd: ItemIK.Ptr = VarPtr(m_Arr(Index))
'End Property

Public Sub Insert(ByVal Index As Long, Obj)
    If Index < 0 Or m_Count <= Index Then Exit Sub
    EnsureCapacity m_Count + 1
    Dim pDst As LongPtr: pDst = VarPtr(m_Arr(Index + 1))
    Dim pSrc As LongPtr: pSrc = VarPtr(m_Arr(Index))
    Dim blen As Long: blen = (m_Count - 1 - (Index)) * 28
    RtlMoveMemory ByVal pDst, ByVal pSrc, blen
    Obj.Ptr = VarPtr(m_Arr(Index)): Obj.SetData
    m_Count = m_Count + 1
End Sub

Public Sub Remove(ByVal Index As Long)
    If Index < 0 Or m_Count <= Index Then Exit Sub
    Dim pDst As LongPtr: pDst = VarPtr(m_Arr(Index))
    Dim pSrc As LongPtr: pSrc = VarPtr(m_Arr(Index + 1))
    Dim blen As Long: blen = (m_Count - 1 - (Index)) * 28
    RtlMoveMemory ByVal pDst, ByVal pSrc, blen
    m_Count = m_Count - 1
End Sub

Public Sub Send()
    Dim hr As Long: hr = SetForegroundWindow(m_hWndReceiver)
    hr = SendInput(m_Count, m_Arr(0), 28)
    'hr = SetForegroundWindow(m_hWndSender)
End Sub

Public Sub ToListBox(aLB As ListBox)
    aLB.Clear
    Dim i As Long, Obj 'As Object
    For i = 0 To m_Count - 1
        Set Obj = Me.Item(i)
        aLB.AddItem Obj.Key
    Next
End Sub

Public Sub Save(ByVal aPFN As String)
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open aPFN For Binary Access Write As FNr
    m_Date = Now
    Put FNr, , m_Date
    Put FNr, , m_Count
    
    'Put FNr, , m_Arr() ' Does not compile:
                        ' -----------------
                        ' (de) ' Fehler beim Kompilieren:
                               ' Objektverweisvariable oder Variable eines
                               ' benutzerdefinierten Typs, die einen Objektverweis enthält,
                               ' unterstützt Get oder Put nicht
                        ' (en) ' Error during compile:
                               ' Object instance or variable of user defined type containing object instances
                               ' is not supported by Get or Put
    ' so we have to do it this way:
    Dim i As Long
    For i = 0 To m_Count - 1
        With m_Arr(i)
            Put FNr, , .dwType
            Put FNr, , .Value1
            Put FNr, , .Value2
            Put FNr, , .Value3
            Put FNr, , .Value4
            Put FNr, , .Value5
            Put FNr, , .Value6
        End With
    Next
    GoTo Finally
Catch: MsgBox Err.Number & " " & Err.Description
Finally: Close FNr
End Sub

Public Sub Read(ByVal aPFN As String)
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open aPFN For Binary Access Read As FNr
    Get FNr, , m_Date
    Get FNr, , m_Count
    EnsureCapacity m_Count
    'ReDim m_Arr(0 To m_Count - 1)
    'Get FNr, , m_Arr
    Dim i As Long
    For i = 0 To m_Count - 1
        With m_Arr(i)
            Get FNr, , .dwType
            Get FNr, , .Value1
            Get FNr, , .Value2
            Get FNr, , .Value3
            Get FNr, , .Value4
            Get FNr, , .Value5
            Get FNr, , .Value6
        End With
    Next
    GoTo Finally
Catch: MsgBox Err.Number & " " & Err.Description
Finally: Close FNr
End Sub


