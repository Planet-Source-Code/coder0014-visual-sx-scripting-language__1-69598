Attribute VB_Name = "modGlobal"
Global I As Long
    Public C2 As String
        Public FullProg As String
            Global tmpI, tmpI2, tmpI3, tmpI4, tmpI5 As Integer
                Global tempString, tmpStr1, tmpStr2, tmpStr3 As String
                    Global tmpStr4, tmpStr5, tmpStr6 As String
                        Global v1, v2, v3, v4, v5, v6, v7 As String
                            Global c As String
                            Global tmpVal As Integer
        
        Global VarCount As Long
            Global lblCount As Integer
            
            Global Var() As Variable
                Global Lbl() As Label
                    
                    Type Variable
                        VarData As Variant
                        VarName As String
                    End Type
                    
                    Type Label
                        lblName As String
                        Location As Long
                    End Type
                    
Public Declare Function GetTickCount Lib "kernel32" () As Long
