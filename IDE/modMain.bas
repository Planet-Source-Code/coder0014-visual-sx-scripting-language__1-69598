Attribute VB_Name = "modMain"
Global DllLocation As String
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public cmdIndex, picIndex, txtIndex, lblIndex, txfIndex, dirIndex, filIndex, drvIndex     As Long
Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function
Function SetRuntimeDllLocation()
If FileExists(App.Path & "\Runtime.DLL") Then
DllLocation = App.Path & "\Runtime.DLL"
ElseIf FileExists("C:\Windows\System32\Runtime.DLL") Then
DllLocation = "C:\Windows\System32\Runtime.DLL"
ElseIf FileExists("C:\Windows\Runtime.DLL") Then
DllLocation = "C:\Windows\Runtime.DLL"
ElseIf FileExists("C:\Winnt\Runtime.DLL") Then
DllLocation = "C:\Winnt\Runtime.DLL"
ElseIf FileExists("C:\Winnt\System32\Runtime.DLL") Then
DllLocation = "C:\Winnt\System32\Runtime.DLL"
ElseIf FileExists("C:\Runtime.DLL") Then
DllLocation = "C:\Runtime.DLL"
Else
MsgBox "Fatal error : SX++ Runtime library was not found": End
End If
End Function
Function MakeEXE(FileName As String)
DoEnvCode
On Error GoTo errHandler
mdiMain.txtDebug.Text = ""
Close #1
Close #2
    If FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        fDebug "Compiling file.. "
        fDebug "Copying Runtime.DLL"
        FileCopy DllLocation, FileName
        
        Open FileName For Output As #nFile
        fDebug "Writing header .."
        Print #nFile, "->SX++:" & frmMain.txtCode.Text & frmMain.envCode.Text
        Close #nFile
        Dim sTemp As String, sTemp2 As String
        Open FileName For Output As #1
        Open DllLocation For Binary As #2
        While Not EOF(2)
            DoEvents
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
        Wend
        fDebug "Making EXE .."
        Print #1, "->SX++:" & frmMain.txtCode.Text & frmMain.envCode.Text
        Close #2
        Close #1
    End If

errHandler:

If Err.Number <> 0 Then mdiMain.txtDebug.BackColor = vbRed: fDebug " ERROR : " & Err.Number & " " & " : " & Err.Description: frmMain.Show: Exit Function
    Exit Function
End Function
Function FileExists(ByVal FileName As String) As Boolean
On Error Resume Next
If Dir(FileName, vbSystem + vbHidden) = "" Then FileExists = False Else: FileExists = True
End Function
Function ShowAbout()
MsgBox "SX++ Programing language with compiler. " & vbCrLf _
       & " Programed by Nick Chepkasov. "
End Function
Function fDebug(Text)
mdiMain.txtDebug.ZOrder 0
mdiMain.txtDebug.Text = mdiMain.txtDebug.Text + "SX++ IDE Debug    ->" & Text & vbCrLf
End Function

Function CopyControl(Control As Variant, Caption As String, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Text As String, cIndex As String)
On Error Resume Next
Load Control(cIndex)
With Control(cIndex)
.Tag = Tag
.Caption = Caption
.Text = Text
.Visible = Visible
.Top = Top
.Left = Left
.Width = Width
.Height = Height
End With
End Function


Function DoEnvCode()

frmMain.envCode.Text = ""
frmMain.envCode.Text = frmMain.envCode.Text + "env{"
frmMain.envCode.Text = frmMain.envCode.Text + "wnd["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Caption & "]"
For vBtn = frmWindow.Button.LBound + 1 To frmWindow.Button.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "btn["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Button(vBtn).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Button(vBtn).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Button(vBtn).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Button(vBtn).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Button(vBtn).Caption & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Button(vBtn).Index & "]"
Next vBtn
For vLbl = frmWindow.Label1.LBound + 1 To frmWindow.Label1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "lbl["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Label1(vLbl).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Label1(vLbl).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Label1(vLbl).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Label1(vLbl).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Label1(vLbl).Caption & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Label1(vLbl).Index & "]"
Next vLbl
For vTxt = frmWindow.Text1.LBound + 1 To frmWindow.Text1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "txt["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Text1(vTxt).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Text1(vTxt).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Text1(vTxt).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Text1(vTxt).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Text1(vTxt).Text & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Text1(vTxt).Index & "]"
Next vTxt
For vTxf = frmWindow.TextField1.LBound + 1 To frmWindow.TextField1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "txf["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.TextField1(vTxf).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.TextField1(vTxf).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.TextField1(vTxf).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.TextField1(vTxf).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.TextField1(vTxf).Text & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.TextField1(vTxf).Index & "]"
Next vTxf
For vPic = frmWindow.Picture1.LBound + 1 To frmWindow.Picture1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "pic["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Picture1(vPic).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Picture1(vPic).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Picture1(vPic).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Picture1(vPic).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Picture1(vPic).Index & "]"
Next vPic
For vFil = frmWindow.File1.LBound + 1 To frmWindow.File1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "fil["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.File1(vFil).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.File1(vFil).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.File1(vFil).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.File1(vFil).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.File1(vFil).Index & "]"
Next vFil
For vDir = frmWindow.Dir1.LBound + 1 To frmWindow.Dir1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "dir["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Dir1(vDir).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Dir1(vDir).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Dir1(vDir).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Dir1(vDir).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Dir1(vDir).Index & "]"
Next vDir
On Error Resume Next
For vDrv = frmWindow.Drive1.LBound + 1 To frmWindow.Drive1.UBound
frmMain.envCode.Text = frmMain.envCode.Text + "drv["
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Drive1(vDrv).Left & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Drive1(vDrv).Top & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Drive1(vDrv).Width & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Drive1(vDrv).Height & ","
frmMain.envCode.Text = frmMain.envCode.Text & frmWindow.Drive1(vDrv).Index & "]"
Next vDrv
End Function
Function ExecEnvCode(envCode As String)
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
        CopyControl frmWindow.Button, (c), True, (v2), (v1), (v3), (v4), "", (v6)
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
        CopyControl frmWindow.Text1, "", True, (v2), (v1), (v3), (v4), (c), (v6)
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
        CopyControl frmWindow.TextField1, "", True, (v2), (v1), (v3), (v4), (c), (v6)
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
        CopyControl frmWindow.Label1, (c), True, (v2), (v1), (v3), (v4), "", (v6)
         v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "pic[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmWindow.Picture1, "", True, (v2), (v1), (v3), (v4), "", (v5)
         v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "fil[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmWindow.File1, "", True, (v2), (v1), (v3), (v4), "", (v5)
         v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "dir[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmWindow.Dir1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 8)) = "drv[drv[" Then
    e = e + 8
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmWindow.Drive1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "drv[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        CopyControl frmWindow.Drive1, "", True, (v2), (v1), (v3), (v4), "", (v5)
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "wnd[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v3 = v3 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v4 = v4 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v5 = v5 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        frmWindow.Left = v1: frmWindow.Top = v2: frmWindow.Width = v3: frmWindow.Height = v4: frmWindow.Caption = v5
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
    If LCase(Mid(envCode, e, 4)) = "tmr[" Then
    e = e + 4
        Do: v1 = v1 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = ",": e = e + 1
        Do: v2 = v2 + Mid(envCode, e, 1): e = e + 1: Loop Until Mid(envCode, e, 1) = "]"
        'CopyControl frmWindow.Timer1, 0, 0, 0, 0, 0, 0, 0, (v2)
        v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = "": c = ""
    End If
Next e
End Function

