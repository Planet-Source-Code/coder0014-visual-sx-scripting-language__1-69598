VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private cInput As Long, cOut As Long, cErr As Long

Public Sub Class_Initialize()
    AllocConsole
    cInput = GetStdHandle(STD_INPUT_HANDLE)
    cOut = GetStdHandle(STD_OUTPUT_HANDLE)
    cErr = GetStdHandle(STD_ERROR_HANDLE)
    SetConsoleMode cInput, ENABLE_ECHO_INPUT
End Sub

Public Sub WriteLine(Str As String)
    WriteConsole cOut, Str, Len(Str), vbNull, vbNull
End Sub
Public Function ReadLine() As String
    On Error Resume Next
    Dim sInput As String * 128
    ReadConsole cInput, sInput, Len(sInput), vbNull, vbNull
    ReadLine = Left$(sInput, InStr(sInput, Chr$(0)) - 3)
End Function
Public Sub Init()
Class_Initialize
End Sub
Public Sub SetConTitle(Title)
    SetConsoleTitle Title
End Sub
Public Sub SetConColor(Number As Byte)
    SetConsoleTextAttribute cOut, Number
End Sub
Public Sub Terminate()
Class_Terminate
End Sub

Public Sub Class_Terminate()
    FreeConsole
End Sub


