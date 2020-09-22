VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ">>> SX++ Compiled Application <<<"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   1000
      Left            =   2445
      Top             =   75
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Ïëîñêà
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1095
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Ïëîñêà
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   1065
      TabIndex        =   6
      Top             =   855
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Ïëîñêà
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   1065
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TextField1 
      Appearance      =   0  'Ïëîñêà
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Îáà
      TabIndex        =   4
      Text            =   "frmMain.frx":0000
      Top             =   1410
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ïëîñêà
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   90
      ScaleHeight     =   270
      ScaleWidth      =   870
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Ïëîñêà
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Text            =   "TextBox"
      Top             =   420
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton Button 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Ïðîçðà÷íî
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   765
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// SX++ Programing language.
'// SX++ 1.0.01 Made by me in 16-10-2007 / 19:37.
'// ---------------------------------------------
Dim FileSize As Long, iTemp As Long
Dim FileData As String, sTemp As String



Private Sub Dir1_Change(Index As Integer)
SetVarData "dirlistbox" & Index & ".path", Dir1(Index).Path
GetSubCode "dirlistbox" & Index & "_click()"
End Sub

Private Sub Drive1_Change(Index As Integer)
SetVarData "drivelistbox" & Index & ".drive", Drive1(Index).Drive
GetSubCode "drivelistbox" & Index & "_click()"
End Sub

Private Sub File1_Click(Index As Integer)
SetVarData "filelistbox" & Index & ".filename", File1(Index).FileName
SetVarData "filelistbox" & Index & ".path", File1(Index).Path
GetSubCode "filelistbox" & Index & "_click()"
End Sub

Private Sub File1_DblClick(Index As Integer)
SetVarData "filelistbox" & Index & ".filename", File1(Index).FileName
SetVarData "filelistbox" & Index & ".path", File1(Index).Path
GetSubCode "filelistbox" & Index & "_dblclick()"
End Sub

Private Sub Form_Load()

ReDim Var(0)
ReDim Lbl(0)
    Open App.Path + "\" + App.EXEName + ".EXE" For Binary As #1
    FileSize = LOF(1)
    FileData = Space$(LOF(1))
    Get #1, , FileData
    iTemp = InStr(1, FileData, "->SX++:")
    If iTemp <> 0 Then
        iTemp = iTemp + 7
        sTemp = String(1000, 0)
        Get #1, iTemp, sTemp
        FullProg = sTemp
        ExecEnvCode FullProg
        RunCode FullProg

    End If
        Close #1
GetSubCode "window_load()"
End Sub
Private Sub Button_Click(Index As Integer)
GetSubCode "button" & Index & "_click()"
End Sub

Private Sub Label1_Click(Index As Integer)
GetSubCode "label" & Index & "_click()"
End Sub

Private Sub Picture1_Click(Index As Integer)
GetSubCode "picture" & Index & "_click()"
End Sub

Private Sub Text1_Change(Index As Integer)
SetVarData "textbox" & Index & ".text", frmMain.Text1(Index).Text
End Sub

Private Sub Text1_Click(Index As Integer)
GetSubCode "textbox" & Index & "_click()"
End Sub
Private Sub TextField1_Change(Index As Integer)
SetVarData "textfield" & Index & ".text", frmMain.TextField1(Index).Text
End Sub
Private Sub TextField1_Click(Index As Integer)
GetSubCode "textfield" & Index & "_click()"
End Sub

Private Sub Timer1_Timer(Index As Integer)
GetSubCode "timer" & Index & "_timer()"
End Sub
'------------------This line is being here since 19:40 / 16-10-2007---------------------'\

