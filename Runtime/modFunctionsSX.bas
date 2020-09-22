Attribute VB_Name = "modFunctionsSX"
Dim sTmp As String
Dim sTmp2, sTmp3, sTmp4, sTmp5, sTmp6 As String
Function DoFunction(sFunction As String) As String
For f = 1 To Len(sFunction)
If LCase(Mid(sFunction, f, 7)) = "[lcase]" Then
f = f + 7
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 8)) = "[_lcase]"
DoFunction = SolveEquation2(LCase(sTmp)): sTmp = ""
End If
If LCase(Mid(sFunction, f, 7)) = "[ucase]" Then
f = f + 7
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 8)) = "[_ucase]"
DoFunction = SolveEquation2(UCase(sTmp)): sTmp = ""
End If
If LCase(Mid(sFunction, f, 7)) = "[right]" Then
f = f + 7
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 1)) = ",": f = f + 1
Do
sTmp2 = sTmp2 + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 8)) = "[_right]"
sTmp = SolveEquation2(sTmp)
sTmp2 = SolveEquation2(sTmp2)
DoFunction = Right(sTmp, sTmp2)
End If
If LCase(Mid(sFunction, f, 6)) = "[left]" Then
f = f + 6
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 1)) = ",": f = f + 1
Do
sTmp2 = sTmp2 + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 7)) = "[_left]"
sTmp = SolveEquation2(sTmp)
sTmp2 = SolveEquation2(sTmp2)
DoFunction = Left(sTmp, sTmp2)
End If
If LCase(Mid(sFunction, f, 5)) = "[mid]" Then
f = f + 5
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 1)) = ",": f = f + 1
Do
sTmp2 = sTmp2 + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 1)) = ",": f = f + 1
Do
sTmp3 = sTmp3 + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 6)) = "[_mid]"
sTmp = SolveEquation2(sTmp)
sTmp2 = SolveEquation2(sTmp2)
sTmp3 = SolveEquation2(sTmp3)
DoFunction = Mid(sTmp, sTmp2, sTmp3)
End If
If LCase(Mid(sFunction, f, 7)) = "[instr]" Then
f = f + 7
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 1)) = ",": f = f + 1
Do
sTmp2 = sTmp2 + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 1)) = ",": f = f + 1
Do
sTmp3 = sTmp3 + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 8)) = "[_instr]"
sTmp = SolveEquation2(sTmp)
sTmp2 = SolveEquation2(sTmp2)
sTmp3 = SolveEquation2(sTmp3)
DoFunction = InStr(sTmp, sTmp2, sTmp3)
End If
If LCase(Mid(sFunction, f, 5)) = "[hex]" Then
f = f + 5
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 6)) = "[_hex]"
DoFunction = Hex(SolveEquation2(sTmp)): sTmp = ""
End If
If LCase(Mid(sFunction, f, 5)) = "[len]" Then
f = f + 5
Do
sTmp = sTmp + Mid(sFunction, f, 1)
f = f + 1
Loop Until LCase(Mid(sFunction, f, 6)) = "[_len]"
DoFunction = Len(SolveEquation2(sTmp)): sTmp = ""
End If
Next f
End Function
