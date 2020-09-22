Attribute VB_Name = "modMain"
Public Function RunCode(Code As String): On Error Resume Next
For I = 1 To Len(Code)
    If FindCommand2("message") Then MsgBox tempString: ClearTmpStr
     If FindCommand2("addtext") Then frmMain.Print (tempString) 'Con.WriteLine (tempString): ClearTmpStr
      If FindCommand2("outtext") Then frmMain.Print (tempString)  'Con.WriteLine tempString & Chr(10): ClearTmpStr '//CR
       If FindCommand2("exec") Then Shell tempString: ClearTmpStr
        If FindCommand2("delay") Then Delay Val(tempString): ClearTmpStr
         If FindCommand2("delete") Then Kill tempString: ClearTmpStr
          If FindCommand2("mkdir") Then MkDir tempString: ClearTmpStr
           If FindCommand2("rmdir") Then RmDir tempString: ClearTmpStr
            If FindCommand2("title") Then frmMain.Caption = tempString: ClearTmpStr
             If FindCommand2("eval") Then MsgBox Eval(tempString): ClearTmpStr
              If FindCommand2("call") Then GetSubCode (tempString): ClearTmpStr
               
                If FindCommand("global ") Then
                 tempString = GetTempString(";")
                  If VariableExist(tempString) Then
                   ShowErrorMessage "Variable already exist"
                    Else
                     AddVariable (tempString): ClearTmpStr
                      End If
                       End If
                If FindCommand("inputbox ") Then
                 tmpStr1 = GetTempString("," & Chr(34)): Inc 2
                  tmpStr2 = GetTempString(Chr(34) & ";")
                   If VariableExist(tmpStr1) Then
                    If VariableExist(tmpStr2) Then
                     SetVarData (tmpStr1), InputBox(FindVarData((tmpStr2)))
                      Else
                       SetVarData (tmpStr1), InputBox(tmpStr2)
                        End If
                         Else
                          ShowErrorMessage "Variable not defined"
                           End If
        End If
       ' If FindCommand("input ") Then
       '    tmpStr1 = GetTempString(";")
       '     If VariableExist(tmpStr1) Then
       '      SetVarData (tmpStr1), Con.ReadLine
       '       Else
       '        ShowErrorMessage "Variable not defined"
       '         End If
       ' End If
     FindCommand ("endif")
      If FindCommand("if ") Then
        tempString = GetTempString(" then")
         If Eval(tempString) = False Then R "endif": ClearTmpStr2
          End If
        If FindCommand("label ") Then
         tempString = GetTempString(";")
          AddLabel (tempString), (I)
           End If
            If FindCommand("goto ") Then
             tempString = GetTempString(";")
              I = FindLabelLocation((tempString)): ClearTmpStr
               End If
            If FindCommand("set ") Then
             tmpStr1 = Trim(GetTempString("="))
              Inc 1
               tmpStr2 = Trim(GetTempString(";"))
                Inc 1
                    If VariableExist(tmpStr1) Then
                     SetVarData (tmpStr1), SolveEquation2(tmpStr2)
                      Else
                       ShowErrorMessage "Variable not exist"
                        End If
                    End If
            If FindCommand("setf ") Then
             tmpStr1 = Trim(GetTempString("="))
              Inc 1
               tmpStr2 = Trim(GetTempString(";"))
                Inc 1
                 If VariableExist(tmpStr1) Then
                  SetVarData (tmpStr1), DoFunction((tmpStr2))
                   Else
                    ShowErrorMessage "Variable not exist"
                     End If
                      End If
                 If FindCommand("for ") Then
                  tmpI = I - 4: tmpStr1 = Trim(GetTempString("="))
                   Inc 1: tmpStr2 = Trim(GetTempString("to")): Inc 2
                    tmpStr3 = Trim(GetTempString(";")): Inc 1
                     tmpVal = Val(tmpStr2): SetVarData (tmpStr1), tmpVal
                      End If
                       If FindCommand("next;") Then
                        If FindVarData((tmpStr1)) <= tmpStr3 - 1 Then: I = tmpI: tmpVal = tmpVal + 1: SetVarData (tmpStr1), tmpVal
                         End If
        If FindCommand("do until ") Then
         tmpStr3 = GetTempString(";")
          tmpI = I - 9
           Inc 1
            End If
             If FindCommand("loop;") Then
              If Eval(tmpStr3) = False Then
               I = tmpI
                End If
                 End If
        'If FindCommand("program ") Then R "endp;"
    'If LCase(Mid(Code, I, 9)) = "doevents;" Then DoEvents
    'If LCase(Mid(Code, I, 5)) = "stop;" Then Exit Function
    'If LCase(Mid(Code, I, 9)) = "con.hide;" Then Con.Terminate
    'If LCase(Mid(Code, I, 9)) = "con.show;" Then Con.Init
    'If LCase(Mid(Code, I, 4)) = "end;" Then End
'[-----------------------------------------------------------------]'
If FindCommand("set_textfield.text[") Then
 tmpStr1 = Trim(GetTempString(",")): Inc 1: tmpStr2 = Trim(GetTempString("];")): Inc 2
  frmMain.TextField1(Val(tmpStr1)).Text = SolveEquation2((tmpStr2))
   End If
    If FindCommand("set_textfield.left[") Then
     tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
      frmMain.TextField1(Val(tmpStr1)).Left = Val(tmpStr2)
       End If
        If FindCommand("set_textfield.top[") Then
         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
          frmMain.TextField1(Val(tmpStr1)).Top = Val(tmpStr2)
           End If
            If FindCommand("set_textfield.width[") Then
             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
              frmMain.TextField1(Val(tmpStr1)).Width = Val(tmpStr2)
               End If
                If FindCommand("set_textfield.height[") Then
                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                  frmMain.TextField1(Val(tmpStr1)).Height = Val(tmpStr2)
                   End If
                    If FindCommand("set_textbox.text[") Then
                     tmpStr1 = Trim(GetTempString(",")): Inc 1: tmpStr2 = Trim(GetTempString("];")): Inc 2
                      frmMain.Text1(Val(tmpStr1)).Text = SolveEquation2((tmpStr2))
                       End If
                        If FindCommand("set_textbox.left[") Then
                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                          frmMain.Text1(Val(tmpStr1)).Left = Val(tmpStr2)
                           End If
                            If FindCommand("set_textbox.top[") Then
                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                              frmMain.Text1(Val(tmpStr1)).Top = Val(tmpStr2)
    End If
     If FindCommand("set_textbox.width[") Then
      tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
       frmMain.Text1(Val(tmpStr1)).Width = Val(tmpStr2)
        End If
         If FindCommand("set_textbox.height[") Then
          tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
           frmMain.Text1(Val(tmpStr1)).Height = Val(tmpStr2)
            End If
             If FindCommand("set_button.caption[") Then
              tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
               frmMain.Button(Val(tmpStr1)).Caption = SolveEquation2(tmpStr2)
                End If
                 If FindCommand("set_button.left[") Then
                  tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                   frmMain.Button(Val(tmpStr1)).Left = Val(tmpStr2)
                    End If
                     If FindCommand("set_button.top[") Then
                      tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                       frmMain.Button(Val(tmpStr1)).Top = Val(tmpStr2)
                        End If
                         If FindCommand("set_button.height[") Then
                          tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                           frmMain.Button(Val(tmpStr1)).Height = Val(tmpStr2)
                            End If
                             If FindCommand("set_button.width[") Then
                              tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                               frmMain.Button(Val(tmpStr1)).Width = Val(tmpStr2)
                                End If
            If FindCommand("set_label.caption[") Then
             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
              frmMain.Label1(Val(tmpStr1)).Caption = SolveEquation2(tmpStr2)
               End If
                If FindCommand("set_label.left[") Then
                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                  frmMain.Label1(Val(tmpStr1)).Left = Val(tmpStr2)
                  End If
                    If FindCommand("set_label.top[") Then
                     tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                      frmMain.Label1(Val(tmpStr1)).Top = Val(tmpStr2)
                       End If
                        If FindCommand("set_label.height[") Then
                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                          frmMain.Label1(Val(tmpStr1)).Height = Val(tmpStr2)
                           End If
                            If FindCommand("set_label.width[") Then
                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                              frmMain.Label1(Val(tmpStr1)).Width = Val(tmpStr2)
                               End If
            If FindCommand("set_picture.image[") Then
             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
              frmMain.Picture1(Val(tmpStr1)).Picture = LoadPicture(SolveEquation2(tmpStr2))
               End If
                If FindCommand("set_picture.left[") Then
                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                  frmMain.Picture1(Val(tmpStr1)).Left = Val(tmpStr2)
                   End If
                    If FindCommand("set_picture.top[") Then
                     tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                      frmMain.Picture1(Val(tmpStr1)).Top = Val(tmpStr2)
                       End If
                        If FindCommand("set_picture.height[") Then
                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                          frmMain.Picture1(Val(tmpStr1)).Height = Val(tmpStr2)
                           End If
                            If FindCommand("set_picture.width[") Then
                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                              frmMain.Picture1(Val(tmpStr1)).Width = Val(tmpStr2)
                               End If
                    If FindCommand("set_drivelistbox.drive[") Then
                     tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                      frmMain.Drive1(Val(tmpStr1)).Drive = SolveEquation2(tmpStr2)
                       End If
                        If FindCommand("set_drivelistbox.left[") Then
                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                          frmMain.Drive1(Val(tmpStr1)).Left = Val(tmpStr2)
                           End If
                            If FindCommand("set_drivelistbox.top[") Then
                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                              frmMain.Drive1(Val(tmpStr1)).Top = Val(tmpStr2)
                               End If
                               If FindCommand("set_drivelistbox.height[") Then
                                tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                 frmMain.Drive1(Val(tmpStr1)).Height = Val(tmpStr2)
                                  End If
                                   If FindCommand("set_drivelistbox.width[") Then
                                    tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                     frmMain.Drive1(Val(tmpStr1)).Width = Val(tmpStr2)
                                      End If
                        If FindCommand("set_dirlistbox.path[") Then
                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                          frmMain.Dir1(Val(tmpStr1)).Path = SolveEquation2(tmpStr2)
                           End If
                            If FindCommand("set_dirlistbox.left[") Then
                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                              frmMain.Dir1(Val(tmpStr1)).Left = Val(tmpStr2)
                               End If
                                If FindCommand("set_dirlistbox.top[") Then
                                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                  frmMain.Dir1(Val(tmpStr1)).Top = Val(tmpStr2)
                                   End If
                                    If FindCommand("set_dirlistbox.height[") Then
                                     tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                      frmMain.Dir1(Val(tmpStr1)).Height = Val(tmpStr2)
                                       End If
                                        If FindCommand("set_dirlistbox.width[") Then
                                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                          frmMain.Dir1(Val(tmpStr1)).Width = Val(tmpStr2)
                                           End If
                            If FindCommand("set_filelistbox.pattern[") Then
                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                              frmMain.File1(Val(tmpStr1)).Pattern = SolveEquation2(tmpStr2)
                               End If
                                If FindCommand("set_filelistbox.path[") Then
                                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                  frmMain.File1(Val(tmpStr1)).Path = SolveEquation2(tmpStr2)
                                   End If
                                    If FindCommand("set_filelistbox.left[") Then
                                     tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                      frmMain.File1(Val(tmpStr1)).Left = Val(tmpStr2)
                                       End If
                                        If FindCommand("set_filelistbox.top[") Then
                                         tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                          frmMain.File1(Val(tmpStr1)).Top = Val(tmpStr2)
                                           End If
                                            If FindCommand("set_filelistbox.height[") Then
                                             tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                              frmMain.File1(Val(tmpStr1)).Height = Val(tmpStr2)
                                               End If
                                                If FindCommand("set_filelistbox.width[") Then
                                                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                                  frmMain.File1(Val(tmpStr1)).Width = Val(tmpStr2)
                                                   End If
                                If FindCommand("set_timer.interval[") Then
                                 tmpStr1 = GetTempString(","): Inc 1: tmpStr2 = GetTempString("];"): Inc 2
                                  frmMain.Timer1(Val(tmpStr1)).Interval = SolveEquation2(Val(tmpStr2))
                                   End If
Next I
End Function
Function ExecEnvCode(envCode As String)
On Error Resume Next
For e = 1 To Len(envCode)
'If LCase(Mid(envCode, e, 4)) = "env{" Then
'e = e + 4
    If LCase(Mid(envCode, e, 4)) = "btn[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: c = c + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v6 = v6 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Button, (c), True, (v2), (v1), (v3), (v4), "", (v6)
        AddVariable "button" & v6 & ".caption": SetVarData "button" & v6 & ".caption", frmMain.Button(v6).Caption
        AddVariable "button" & v6 & ".left": SetVarData "button" & v6 & ".left", frmMain.Button(v6).Left
        AddVariable "button" & v6 & ".top": SetVarData "button" & v6 & ".top", frmMain.Button(v6).Top
        AddVariable "button" & v6 & ".width": SetVarData "button" & v6 & ".width", frmMain.Button(v6).Width
        AddVariable "button" & v6 & ".height": SetVarData "button" & v6 & ".height", frmMain.Button(v6).Height
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "txt[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: c = c + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v6 = v6 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Text1, "", True, (v2), (v1), (v3), (v4), c, (v6)
        AddVariable "textbox" & v6 & ".text": SetVarData "textbox" & v6 & ".text", frmMain.Text1(v6).Text
        AddVariable "textbox" & v6 & ".left": SetVarData "textbox" & v6 & ".left", frmMain.Text1(v6).Left
        AddVariable "textbox" & v6 & ".top": SetVarData "textbox" & v6 & ".top", frmMain.Text1(v6).Top
        AddVariable "textbox" & v6 & ".width": SetVarData "textbox" & v6 & ".width", frmMain.Text1(v6).Width
        AddVariable "textbox" & v6 & ".height": SetVarData "textbox" & v6 & ".height", frmMain.Text1(v6).Height
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
  End If
  If LCase(Mid(envCode, e, 4)) = "txf[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: c = c + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v6 = v6 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.TextField1, "", True, (v2), (v1), (v3), (v4), c, (v6)
        AddVariable "textfield" & v6 & ".text": SetVarData "textfield" & v6 & ".text", frmMain.TextField1(v6).Text
        AddVariable "textfield" & v6 & ".left": SetVarData "textfield" & v6 & ".left", frmMain.TextField1(v6).Left
        AddVariable "textfield" & v6 & ".top": SetVarData "textfield" & v6 & ".top", frmMain.TextField1(v6).Top
        AddVariable "textfield" & v6 & ".width": SetVarData "textfield" & v6 & ".width", frmMain.TextField1(v6).Width
        AddVariable "textfield" & v6 & ".height": SetVarData "textfield" & v6 & ".height", frmMain.TextField1(v6).Height
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
  End If
    If LCase(Mid(envCode, e, 4)) = "lbl[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: c = c + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v6 = v6 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Label1, (c), True, (v2), (v1), (v3), (v4), "", (v6)
        AddVariable "label" & v6 & ".caption": SetVarData "label" & v6 & ".caption", frmMain.Label1(v6).Caption
        AddVariable "label" & v6 & ".left": SetVarData "label" & v6 & ".left", frmMain.Label1(v6).Left
        AddVariable "label" & v6 & ".top": SetVarData "label" & v6 & ".top", frmMain.Label1(v6).Top
        AddVariable "label" & v6 & ".width": SetVarData "label" & v6 & ".width", frmMain.Label1(v6).Width
        AddVariable "label" & v6 & ".height": SetVarData "label" & v6 & ".height", frmMain.Label1(v6).Height

        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "pic[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Picture1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        AddVariable "picture" & (v5) & ".image": SetVarData "picture" & (v5) & ".image", frmMain.Picture1((v5)).Picture
        AddVariable "picture" & (v5) & ".left": SetVarData "picture" & (v5) & ".left", frmMain.Picture1((v5)).Left
        AddVariable "picture" & (v5) & ".top": SetVarData "picture" & (v5) & ".top", frmMain.Picture1((v5)).Top
        AddVariable "picture" & (v5) & ".width": SetVarData "picture" & (v5) & ".width", frmMain.Picture1((v5)).Width
        AddVariable "picture" & (v5) & ".height": SetVarData "picture" & (v5) & ".height", frmMain.Picture1((v5)).Height
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "fil[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.File1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        AddVariable "filelistbox" & (v5) & ".left": SetVarData "filelistbox" & (v5) & ".left", frmMain.File1((v5)).Left
        AddVariable "filelistbox" & (v5) & ".top": SetVarData "filelistbox" & (v5) & ".top", frmMain.File1((v5)).Top
        AddVariable "filelistbox" & (v5) & ".width": SetVarData "filelistbox" & (v5) & ".width", frmMain.File1((v5)).Width
        AddVariable "filelistbox" & (v5) & ".height": SetVarData "filelistbox" & (v5) & ".height", frmMain.File1((v5)).Height
        AddVariable "filelistbox" & (v5) & ".path": SetVarData "filelistbox" & (v5) & ".path", frmMain.File1((v5)).Path
        AddVariable "filelistbox" & (v5) & ".filename": SetVarData "filelistbox" & (v5) & ".filename", frmMain.File1((v5)).FileName
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "dir[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Dir1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        AddVariable "dirlistbox" & (v5) & ".left": SetVarData "dirlistbox" & (v5) & ".left", frmMain.Dir1((v5)).Left
        AddVariable "dirlistbox" & (v5) & ".top": SetVarData "dirlistbox" & (v5) & ".top", frmMain.Dir1((v5)).Top
        AddVariable "dirlistbox" & (v5) & ".width": SetVarData "dirlistbox" & (v5) & ".width", frmMain.Dir1((v5)).Width
        AddVariable "dirlistbox" & (v5) & ".height": SetVarData "dirlistbox" & (v5) & ".height", frmMain.Dir1((v5)).Height
        AddVariable "dirlistbox" & (v5) & ".path": SetVarData "dirlistbox" & (v5) & ".path", frmMain.Dir1((v5)).Path
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "drv[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Drive1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        AddVariable "drivelistbox" & (v5) & ".left": SetVarData "drivelistbox" & (v5) & ".left", frmMain.Drive1((v5)).Left
        AddVariable "drivelistbox" & (v5) & ".top": SetVarData "drivelistbox" & (v5) & ".top", frmMain.Drive1((v5)).Top
        AddVariable "drivelistbox" & (v5) & ".width": SetVarData "drivelistbox" & (v5) & ".width", frmMain.Drive1((v5)).Width
        AddVariable "drivelistbox" & (v5) & ".height": SetVarData "drivelistbox" & (v5) & ".height", frmMain.Drive1((v5)).Height
        AddVariable "drivelistbox" & (v5) & ".drive": SetVarData "drivelistbox" & (v5) & ".drive", frmMain.Drive1((v5)).Drive
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "wnd[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        frmMain.Left = v1: frmMain.Top = v2: frmMain.Width = v3: frmMain.Height = v4: frmMain.Caption = v5
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "tmr[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmMain.Timer1, 0, 0, 0, 0, 0, 0, 0, (v2)
        AddVariable "timer" & (v2) & ".interval": SetVarData "timer" & (v2) & ".interval", frmMain.Timer1((v5)).Interval
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
Next e
End Function
