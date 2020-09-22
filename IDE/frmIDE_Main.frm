VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Ïëîñêà
   BackColor       =   &H00C0C0C0&
   Caption         =   "SX++"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   9465
   Begin VB.TextBox EnvCode 
      Appearance      =   0  'Ïëîñêà
      Height          =   480
      Left            =   6330
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   2970
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   8355
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16777215
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   6120
      Left            =   300
      TabIndex        =   3
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   10795
      _Version        =   393217
      BackColor       =   16777215
      BulletIndent    =   1000
      Appearance      =   0
      TextRTF         =   $"frmIDE_Main.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   255
      X2              =   255
      Y1              =   38700
      Y2              =   -30
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "'---------------------This line is being here since 00:26 / 18-10-2007------------------'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1815
      TabIndex        =   1
      Top             =   -75
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.Label Label3 
      Caption         =   $"frmIDE_Main.frx":008E
      Height          =   2160
      Left            =   1425
      TabIndex        =   0
      Top             =   -1860
      Width           =   4785
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New code"
         Shortcut        =   ^N
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenEnv 
         Caption         =   "Open environment"
      End
      Begin VB.Menu mnuFileSaveEnv 
         Caption         =   "Save environment"
      End
      Begin VB.Menu S8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu S3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu S4 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit_Time 
         Caption         =   "&Time / Date"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewCalculator 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu S6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDebug 
         Caption         =   "&Debug window"
      End
      Begin VB.Menu mnuViewCode 
         Caption         =   "&Code window"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectRun 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu S7 
         Caption         =   "-"
      End
      Begin VB.Menu munProjectMake 
         Caption         =   "&Make exe .."
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fFile As String
Dim UndoText, TextBuffer As String
Dim Q As Double

Private Sub Form_Load()
EnvCode.ZOrder 0
frmWindow.Show
If GetSetting("SX++ IDE", "REG", "COMPANY") = "" Or GetSetting("SX++ IDE", "REG", "NAME") = "" Then
Me.Hide
frmRegister.Show
Else
Me.Caption = "SX++ IDE . Registered to : " & GetSetting("SX++ IDE", "REG", "NAME") & " (" & GetSetting("SX++ IDE", "REG", "COMPANY") & ")"
End If
SetRuntimeDllLocation
End Sub

Private Sub imgMakeEXE_Click()
munProjectMake_Click
End Sub

Private Sub imgRun_Click()
mnuProjectRun_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtCode.Height = frmMain.Height - 500
txtCode.Width = frmMain.Width - 120
txtDebug.Height = frmMain.Height - 500
txtDebug.Width = frmMain.Width - 120
End Sub

Private Sub mnuCodeExamples_ViewAll_Click()
frmCodeExamples.Show
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
Clipboard.SetText txtCode.SelText
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
Clipboard.SetText txtCode.SelText
txtCode.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
txtCode.SelText = Clipboard.GetText
End Sub

Private Sub mnuEnvironment_BGColor_Click()
CM.ShowColor
txtCode.BackColor = CM.Color
SaveSetting "SX++ IDE", "ENV", "BGCOLOR", CM.Color
End Sub

Private Sub mnuEnvironment_Font_Click()
On Error Resume Next
CM.ShowFont
txtCode.Font = CM.FontName
SaveSetting "SX++ IDE", "ENV", "FONT", CM.FontName
End Sub

Private Sub mnuEnvironment_FontColor_Click()
On Error Resume Next
CM.ShowColor
txtCode.ForeColor = CM.FontName
SaveSetting "SX++ IDE", "ENV", "TEXTCOLOR", CM.Color
End Sub

Private Sub mnuFileAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuFileNew_Click()
txtCode.Text = "//SX++ IDE - New code"
End Sub

Private Sub mnuFileOpen_Click()
CM.FileName = "*.SX*"
CM.ShowOpen
fFile = CM.FileName
fFile = FreeFile
If CM.FileName <> "" And CM.FileName <> UCase("*.SX*") Then
Open CM.FileName For Input As fFile
txtCode.Text = Input(LOF(fFile), fFile)
End If
End Sub

Private Sub mnuFileOpenEnv_Click()
On Error Resume Next

            frmMain.CM.FileName = "*.VSXFF"
            frmMain.CM.ShowOpen
            fFile = frmMain.CM.FileName
            fFile = FreeFile
            If frmMain.CM.FileName <> "" And frmMain.CM.FileName <> UCase("*.SX*") Then
For b = frmWindow.Button.LBound + 1 To frmWindow.Button.UBound
Unload frmWindow.Button(b)
Next b
For l = frmWindow.Label1.LBound + 1 To frmWindow.Label1.UBound
Unload frmWindow.Label1(l)
Next l
For t = frmWindow.Text1.LBound + 1 To frmWindow.Text1.UBound
Unload frmWindow.Text1(t)
Next t
For tX = frmWindow.TextField1.LBound + 1 To frmWindow.TextField1.UBound
Unload frmWindow.Text1(tX)
Next tX
For p = frmWindow.Picture1.LBound + 1 To frmWindow.Picture1.UBound
Unload frmWindow.Picture1(p)
Next p
For d = frmWindow.Dir1.LBound + 1 To frmWindow.Dir1.UBound
Unload frmWindow.Dir1(p)
Next d
For drv = frmWindow.Drive1.LBound + 1 To frmWindow.Drive1.UBound
Unload frmWindow.Drive1(drv)
Next drv
For f = frmWindow.File1.LBound + 1 To frmWindow.File1.UBound
Unload frmWindow.File1(f)
Next f
            Open frmMain.CM.FileName For Input As fFile
            frmMain.EnvCode.Text = Input(LOF(fFile), fFile)
            End If
            ExecEnvCode frmMain.EnvCode.Text

End Sub

Private Sub mnuFileSave_Click()
CM.FileName = "*.SX*"
CM.ShowOpen
If CM.FileName <> "" Then
fFile = FreeFile
Open CM.FileName For Output As fFile
Print #fFile, txtCode.Text
Close fFile
End If
End Sub

Private Sub mnuFileSaveEnv_Click()
On Error Resume Next
            frmMain.CM.FileName = "*.VSXFF"
            frmMain.CM.ShowOpen
            If frmMain.CM.FileName <> "" Then
            fFile = FreeFile
            Open frmMain.CM.FileName For Output As fFile
            DoEnvCode
            Print #fFile, frmMain.EnvCode.Text
            Close fFile
            End If
End Sub

Private Sub mnuProjectRun_Click()
MakeEXE "C:\TMP_SX++_APP.EXE"
Shell "C:\TMP_SX++_APP.EXE", vbNormalFocus
End Sub


Private Sub mnuViewCalculator_Click()
frmCalculator.Show
End Sub

Private Sub mnuViewCode_Click()
txtCode.ZOrder 0
End Sub

Private Sub mnuViewDebug_Click()
txtDebug.ZOrder 0
End Sub

Private Sub munEditUndo_Click()
OrigSelStart = txtCode.SelStart
txtCode.Text = UndoText
txtCode.SelStart = OrigSelStart
End Sub
Private Sub munProjectMake_Click()
On Error Resume Next
CM.ShowOpen
txtDebug.Text = ""
MakeEXE CM.FileName
End Sub


Private Sub Timer1_Timer()

EnvCode.ZOrder 0
End Sub

Private Sub txtCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdiMain.txtDebug.BackColor = vbWhite
End Sub

Private Sub txtDebug_DblClick()
txtCode.ZOrder 0
End Sub

