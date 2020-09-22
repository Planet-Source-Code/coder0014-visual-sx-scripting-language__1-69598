Attribute VB_Name = "modLoops"
Function getLoopData(Data As Integer) As Variant
Dim tmpArray As Variant
tmpArray = Split(ForLoops(ForLoopsCount), ":")
Select Case Data
    Case 0: getLoopData = tmpArray(0)
    Case 1: getLoopData = tmpArray(1)
    Case 2: getLoopData = tmpArray(2)
    Case 3: getLoopData = tmpArray(3)
End Select
End Function
Function FindForEnd(StartPos As Long, ByVal Var As String) As Long
Dim x As Long
Dim CurrentVar As String
x = StartPos
search:
x = InStr2(x + 1, Code, "next ")
If x = 0 Then ShowErrorMessage "Error at " & StartPos & ": expected 'next'.", True
CurrentVar = FindTempString1(x + 5, ";")
If CurrentVar = Var Then
    FindForEnd = x + Len(Var) + 1
    Exit Function
Else
    GoTo search
End If
End Function
Function FindLoopEnd(StartPos As Long) As Long
Dim x As Long
Dim loopLocation As Long
Dim Quote As Long
For x = StartPos To Len(Code)
    ElseIf Mid(Code, x, 1) = Chr(34) Then
        Quote = InStr(x + 1, Code, Chr(34))
        If Quote <> 0 Then x = Quote
    If Mid(Code, x, 4) = "loop" Then
        If loopLocation = 0 Then
            FindLoopEnd = x + 4
            Exit Function
        Else
            loopLocation = loopLocation + 1
        End If
    ElseIf Mid(Code, x, 8) = "do until" Then
        loopLocation = loopLocation - 1
    End If
Next x
ShowErrorMessage "Error at " & StartPos & ": expected 'loop'."
End Function
Function FindLoopStart(StartPos As Long) As Long
Dim x As Long, loopLocation As Long, Quote As Boolean
loopLocation = 0
For x = StartPos To 1 vStep -1
    If Mid(Code, x, 8) = "do until" Then
        If Quote Then GoTo skip
        If loopLocation = 0 Then
            FindLoopStart = x - 1
            Exit Function
        Else
            loopLocation = loopLocation - 1
        End If
    ElseIf Mid(Code, x, 4) = "loop" Then
        If Quote Then GoTo skip
        loopLocation = loopLocation + 1
    ElseIf Mid(Code, x, 1) = Chr(34) Then
        If Quote Then Quote = False Else: Quote = True
    End If
skip:
Next x
ShowErrorMessage "Error at " & StartPos & ": loop without do"
End Function
