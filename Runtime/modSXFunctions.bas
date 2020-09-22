Attribute VB_Name = "modSXFunctions"

Function DoFunction(sFunction As String) As Variant
On Error Resume Next
Dim TempFunction As String
Dim Arg1 As String, Arg2 As String, Arg3 As String
Dim Args As Variant, TempString As String
Dim x As Long, Char As String, x As Long
For x = 1 To Len(sFunction)
    Char = Mid(sFunction, x, 1)
    If LCase(Mid(sFunction, x, 6)) = "lcase(" Then
        x = x + 6
        TempString = FindTempString3(x, ")", sFunction)
        If TempString <> "" Then
            Args = Split2(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                ShowErrorMessage "Error at " & iLocation & ": wrong number of arguments for lcase()."
            Else
                For x = LBound(Args) To UBound(Args)
                    TempString = Args(x)
                    Args(x) = Solve(TempString)
                Next x
                TempFunction = TempFunction & Chr(34) & LCase(Args(1)) & Chr(34)
            End If
        End If
    
    Else
        TempFunction = TempFunction & Char
    End If
Next x
DoFunction = TempFunction
End Function


