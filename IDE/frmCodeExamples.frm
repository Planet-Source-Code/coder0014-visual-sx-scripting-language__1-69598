VERSION 5.00
Begin VB.Form frmCodeExamples 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SX++ 1.0.12 Code examples"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Ïëîñêà
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3675
      Left            =   -15
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   540
      Width           =   7905
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   285
      Width           =   945
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Simple input"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   285
      Width           =   1410
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Using DO Until loops"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   285
      Width           =   1725
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Using FOR loops"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   30
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ASCII Table"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1005
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   30
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   30
      Width           =   960
   End
   Begin VB.TextBox txtColors 
      Height          =   240
      Left            =   345
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmCodeExamples.frx":0000
      Top             =   2250
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmCodeExamples.frx":004F
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtUsingForLoops 
      Height          =   285
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmCodeExamples.frx":009E
      Top             =   2220
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtASCII_Table 
      Height          =   285
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmCodeExamples.frx":00D1
      Top             =   2040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtCalculator 
      Height          =   285
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmCodeExamples.frx":0133
      Top             =   2190
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtUsingDoUntilLoops 
      Height          =   285
      Left            =   285
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmCodeExamples.frx":051F
      Top             =   2070
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   4335
      Picture         =   "frmCodeExamples.frx":056E
      Top             =   -75
      Width           =   3525
   End
End
Attribute VB_Name = "frmCodeExamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txtCode.Text = txtCalculator.Text
End Sub

Private Sub Command2_Click()
txtCode.Text = txtASCII_Table.Text
End Sub

Private Sub Command3_Click()
txtCode.Text = txtUsingForLoops.Text
End Sub

Private Sub Command5_Click()
txtCode.Text = txtUsingDoUntilLoops.Text
End Sub

Private Sub Command6_Click()
txtCode.Text = txtInput.Text
End Sub

Private Sub Command7_Click()
txtCode.Text = txtColors.Text
End Sub

