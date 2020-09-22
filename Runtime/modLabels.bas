Attribute VB_Name = "modLabels"
Function LabelExist(vLbl As Variant) As Boolean
For x = 1 To lblName.Count
    If lblName(x) = vLbl Then
        LabelExist = True
        Exit Function
    End If
Next x
LabelExist = False
End Function
Function FindLabel(vLbl As Variant) As Variant
If IsNumeric(vLbl) = True Then
    If vLbl < Len(Code) Then
        FindLabel = vLbl
        Exit Function
    End If
End If
If LabelExist(vLbl) = False Then ShowErrorMessage "Error at " & i & ": label not exist '" & vLbl & "'.", True
For x = 1 To lblName.Count
    If lblName(x) = vLbl Then
        FindLabel = lblData(x)
        Exit Function
    End If
Next x
End Function
Function CreateLabel(vLbl As Variant, lblLocation As Variant)
If LabelExist(vLbl) = True Then
    ShowErrorMessage "Error at " & i & ": cannot create label '" & vLbl & "'.", True
Else
    lblName.Add vLbl
    lblData.Add lblLocation
End If
End Function
