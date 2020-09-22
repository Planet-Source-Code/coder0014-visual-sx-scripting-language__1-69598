Attribute VB_Name = "modVariables"
Function VariableExist(Variable As Variant) As Boolean
For vCount = 1 To vName.Count
    If vName(vCount) = Variable Then
        VariableExist = True
    End If
Next vCount
End Function
Function FindVarData(Variable As Variant) As Variant
If IsNumeric(Variable) Then FindVarData = Variable: Exit Function
For vCount = 1 To vName.Count
    If vName(vCount) = Variable Then FindVarData = vData(vCount)
Next vCount
End Function
Function SetVarData(Variable As Variant, Value As Variant)
For vCount = vName.Count To 1 vStep -1
    If vName(vCount) = Variable Then
        vName.Remove vCount: vData.Remove vCount
        vName.Add Variable: vData.Add Value
    End If
Next vCount
End Function
Function SetVarData2(ByVal sExpression As String)
Dim l, r, Char As String, Operator As Integer
Operator = InStr(1, sExpression, "="): l = Trim(Left(sExpression, Operator - 1)): r = Solve(Trim(Mid(sExpression, Operator + 1))): SetVarData l, r
End Function
Function CreateVariable(VarName As Variant, VariableData As String)
    vName.Add VarName: vData.Add Solve(VariableData)
End Function


