Attribute VB_Name = "modFunctions"
Public Function AddVariable(Name As String)
VarCount = VarCount + 1
ReDim Preserve Var(VarCount)
Var(VarCount).VarName = Name
Debug.Print Var(VarCount).VarName
End Function
Function Delay(Time As Long)
T = Timer
Do While Timer - T < Val(Time)
Loop
End Function
Function FindCommand2(Name As String) As Boolean
            If FindCommand(Name & " " & Chr(34)) Then
                FindCommand2 = True
                tempString = GetTempString(Chr(34) & ";")
            End If
            If FindCommand(Name & " ") Then
                FindCommand2 = True
                tempString = GetTempString(";")
                tempString = FindVarData((tempString))
            End If
            If FindCommand(Name & ":") Then
                FindCommand2 = True
                tempString = GetTempString(";")
                tempString = DoFunction((tempString))
            End If
            If FindCommand(Name & "(") Then
                FindCommand2 = True
                tempString = GetTempString(");")
                tempString = SolveEquation2(tempString)
            End If
End Function
Function GetTempString(End_ As String) As String
                Do
                GetTempString = GetTempString + Mid(C2, I, 1)
                I = I + 1
                    If I >= Len(C2) + 2 Then ShowErrorMessage "Expected '" & End_ & "'"
                Loop Until LCase(Mid(C2, I, Len(End_))) = LCase(End_)
End Function
Function FindCommand(Name As String) As Boolean
If LCase(Mid(FullProg, I, Len(Name))) = Name Then
FindCommand = True
I = I + Len(Name)
Else
FindCommand = False
End If
If LCase(Mid(C2, I, Len(Name))) = Name Then
FindCommand = True
I = I + Len(Name)
Else
FindCommand = False
End If
End Function
Function R(Text As String)
            Do
            I = I + 1
            If I >= Len(FullProg) + 2 Then ShowErrorMessage "Expected '" & Text & "'"
            Loop Until LCase(Mid(FullProg, I, Len(Text))) = LCase(Text)
End Function
Function DeCode(Text) As String
For I = 1 To Len(Text)
    DoEvents
    DeCode = DeCode & Chr(255 - Asc(Mid(Text, I, 1)))
Next I
End Function
Function ClearAll()
    tempString = ""
    tmpStr1 = ""
    tmpStr2 = ""
    tmpStr3 = ""
    I = 0
        For VCount = 0 To UBound(Var)
            Var(VCount).VarData = ""
            Var(VCount).VarName = ""
        Next VCount
        
        For lblCount = 0 To UBound(Lbl)
            Lbl(lblCount).Location = 0
            Lbl(lblCount).lblName = 0
        Next lblCount
End Function
Public Function FindVarData(Name As String)
    For VCount = 0 To UBound(Var)
        If LCase(Name) = "sx.path" Then FindVarData = App.Path
        If LCase(Name) = "sx.gettickcount" Then FindVarData = GetTickCount
        If LCase(Name) = "sx.crlf" Then FindVarData = vbCrLf
        If LCase(Name) = "sx.cr" Then FindVarData = vbCr
        If LCase(Name) = "sx.lf" Then FindVarData = vbLf
        If LCase(Name) = "true" Then FindVarData = True
        If LCase(Name) = "false" Then FindVarData = False
        If Var(VCount).VarName = Name Then
        FindVarData = Var(VCount).VarData
        Exit Function
        End If
    Next VCount
End Function
Public Function SetVarData(Name As String, Data As Variant)
    For VCount = 0 To UBound(Var)
        If Var(VCount).VarName = Name Then
        Var(VCount).VarData = Data
        Debug.Print Var(VarCount).VarData
        Exit Function
        End If
    Next VCount
End Function
Public Function VariableExist(Name) As Boolean
    For VCount = 0 To UBound(Var)
        If Var(VCount).VarName = Name Then
        VariableExist = True
        End If
    Next VCount
End Function
Public Function AddLabel(Name As String, Location As Long)
    lblCount = lblCount + 1
    ReDim Lbl(lblCount)
    Lbl(lblCount).lblName = Name
    Lbl(lblCount).Location = Location
    Debug.Print Lbl(lblCount).lblName & " " & Lbl(lblCount).Location
End Function
Public Function FindLabelLocation(Name As String)
    For lCount = 0 To UBound(Lbl)
        If Lbl(lCount).lblName = Name Then
        FindLabelLocation = Lbl(lCount).Location
        Exit Function
        End If
    Next lCount
End Function
Public Function ReadCode(Path As String) As String
    Dim FileData As String
    Open sPath For Binary As #1
    FileData = Space(LOF(1) - 2)
    Get #1, , FileData
    Close #1
    ReadCode = DeCode(Mid(FileData, InStr(1, FileData, "SX++") + 4))
    Exit Function
End Function
Public Function FileExist(ByVal FileName As String) As Boolean
On Error Resume Next
If Dir(FileName, vbSystem + vbHidden) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function
Public Function ShowErrorMessage(Text As String)
MsgBox "Error at : " & I & " - " & Text
End
End Function

Function ClearTmpStr()
tempString = ""
End Function
Function ClearTmpStr2()
tmpStr1 = ""
tmpStr2 = ""
tmpStr3 = ""
End Function
Function Inc(Number As Long)
I = I + Number
End Function

Public Function Eval(ByVal sFormula As String) As Boolean
    Dim I As Integer, iWait As Integer
    Dim LeftVal As String, RightVal As String, Operator As String
    Dim sTemp As String
    
    For I = 1 To Len(sFormula)
        sTemp = Mid(sFormula, I, 1)
        Select Case sTemp
            Case "("
                iWait = iWait + 1
            Case ")"
                iWait = iWait - 1
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, I)
                    sFormula = Trim(Mid(sFormula, I + 1))
                    Exit For
                End If
            Case Chr(34)
                I = InStr(I + 1, sFormula, Chr(34))
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, I)
                    sFormula = Trim(Mid(sFormula, I + 1))
                    Exit For
                End If
            Case Else
                If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
                    If iWait = 0 Then
                        LeftVal = Trim(Mid(sFormula, 1, I - 1))
                        sFormula = Trim(Mid(sFormula, I))
                        Exit For
                    End If
                End If
        End Select
    Next I
    sTemp = Mid(sFormula, 2, 1)
    If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
        Operator = Left(sFormula, 2)
        sFormula = Mid(sFormula, 3)
    Else
        Operator = Left(sFormula, 1)
        sFormula = Mid(sFormula, 2)
    End If
    
    RightVal = sFormula
    sFormula = ""
    
    LeftVal = SolveEquation2(LeftVal)
    RightVal = SolveEquation2(RightVal)
    
    Select Case Operator
        Case ">"
            If Val(LeftVal) > Val(RightVal) Then Eval = True
        Case "<"
            If Val(LeftVal) < Val(RightVal) Then Eval = True
        Case "<>"
            If Val(LeftVal) <> Val(RightVal) Then Eval = True
        Case ">="
            If Val(LeftVal) >= Val(RightVal) Then Eval = True
        Case "<="
            If Val(LeftVal) <= Val(RightVal) Then Eval = True
        Case "="
            If LeftVal = RightVal Then Eval = True
        Case "#"
            If LCase(LeftVal) = LCase(RightVal) Then Eval = True
    End Select
End Function

Public Function SolveEquation2(ByVal Equation_ As String) As Variant
    Dim iTemp As Integer, iTemp2 As Integer, I As Integer
    Dim sTemp As String
    
    Do
        For I = 1 To Len(Equation_)
            sTemp = Mid(Equation_, I, 1)
            If sTemp = Chr(34) Then
                I = InStr(I + 1, Equation_, Chr(34))
            ElseIf sTemp = "(" Then
                iTemp = I
            ElseIf sTemp = ")" Then
                iTemp2 = I
                Exit For
            End If
        Next I
        If iTemp <> 0 Then
            Equation_ = Mid(Equation_, 1, iTemp - 1) & " " & Chr(34) & SolveEquation1(Mid(Equation_, iTemp + 1, iTemp2 - (iTemp + 1))) & Chr(34) & " " & Mid(Equation_, iTemp2 + 1)
            iTemp = 0: iTemp2 = 0
        Else
            Exit Do
        End If
    Loop
    SolveEquation2 = SolveEquation1(Equation_)
End Function

Public Function SolveEquation1(Equation_ As String) As Variant
    On Error GoTo fErr
    Dim I As Integer, iTemp As Integer
    Dim sTemp As String, sTemp2 As String
    Dim WaitVal As Variant
    Dim WaitOp As String
    Dim WaitVar
    
    For I = 1 To Len(Equation_)
        sTemp = Mid(Equation_, I, 1)
        If sTemp = Chr(34) Then
            iTemp = InStr(I + 1, Equation_, Chr(34))
            WaitVal = Mid(Equation_, I + 1, iTemp - (I + 1))
            I = iTemp
            If WaitOp <> "" Then
                Select Case WaitOp
                    Case "+"
                        SolveEquation1 = Val(SolveEquation1) + Val(WaitVal)
                    Case "-"
                        SolveEquation1 = Val(SolveEquation1) - Val(WaitVal)
                    Case "/"
                        SolveEquation1 = Val(SolveEquation1) / Val(WaitVal)
                    Case "\"
                        SolveEquation1 = Val(SolveEquation1) \ Val(WaitVal)
                    Case "^"
                        SolveEquation1 = Val(SolveEquation1) ^ Val(WaitVal)
                    Case "*"
                        SolveEquation1 = Val(SolveEquation1) * Val(WaitVal)
                    Case "&"
                        SolveEquation1 = SolveEquation1 & WaitVal
                End Select
                WaitOp = ""
            Else
                SolveEquation1 = WaitVal
            End If
        ElseIf sTemp = " " Then
        ElseIf sTemp = "+" Or sTemp = "-" Or sTemp = "/" Or sTemp = "\" Or sTemp = "^" Or sTemp = "&" Or sTemp = "*" Then
            If WaitVar <> 0 Then
                WaitVal = FindVarData(Mid(Equation_, WaitVar, I - (WaitVar + 1)))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                        SolveEquation1 = Val(SolveEquation1) + Val(WaitVal)
                        Case "-"
                            SolveEquation1 = Val(SolveEquation1) - Val(WaitVal)
                        Case "/"
                            SolveEquation1 = Val(SolveEquation1) / Val(WaitVal)
                        Case "\"
                            SolveEquation1 = Val(SolveEquation1) \ Val(WaitVal)
                        Case "^"
                            SolveEquation1 = Val(SolveEquation1) ^ Val(WaitVal)
                        Case "*"
                            SolveEquation1 = Val(SolveEquation1) * Val(WaitVal)
                        Case "&"
                            SolveEquation1 = SolveEquation1 & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    SolveEquation1 = WaitVal
                End If
                
                WaitVar = 0
            End If
            
            WaitOp = sTemp
        Else
            If WaitVar = 0 Then WaitVar = I
            If I >= Len(Equation_) Then
                WaitVal = FindVarData(Mid(Equation_, WaitVar, I))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                            SolveEquation1 = Val(SolveEquation1) + Val(WaitVal)
                        Case "-"
                            SolveEquation1 = Val(SolveEquation1) - Val(WaitVal)
                        Case "/"
                            SolveEquation1 = Val(SolveEquation1) / Val(WaitVal)
                        Case "\"
                            SolveEquation1 = Val(SolveEquation1) \ Val(WaitVal)
                        Case "^"
                            SolveEquation1 = Val(SolveEquation1) ^ Val(WaitVal)
                        Case "*"
                            SolveEquation1 = Val(SolveEquation1) * Val(WaitVal)
                        Case "&"
                            SolveEquation1 = SolveEquation1 & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    SolveEquation1 = WaitVal
                End If
            End If
        End If
    Next I
fErr:
If Err.Description = "" Then Exit Function
End Function
Public Sub GetSubCode(SubName1 As String)
    If InStr(1, LCase(FullProg), LCase("program " & SubName1)) Then
    nameofsub = LCase("program " & SubName1)
    C2 = Mid(FullProg, InStr(1, LCase(FullProg), LCase(nameofsub)) _
        + Len(nameofsub))
    C2 = Mid(C2, 1, InStr(1, LCase(C2), "endp;") - 1)
    RunCode C2
    End If
End Sub
Public Function CopyControl(Control As Variant, Caption As String, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Text As String, cIndex As String)
On Error Resume Next
Load Control(cIndex)
With Control(cIndex)
.Tag = Tag
.Caption = Caption
.Text = Text
.Text = Text
.Visible = Visible
.Top = Top
.Left = Left
.Width = Width
.Height = Height
End With
End Function
