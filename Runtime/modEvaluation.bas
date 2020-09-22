Attribute VB_Name = "modEvaluation"
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
    MsgBox Err.Description, vbExclamation, "Error"
End Function

