VERSION 5.00
Begin VB.Form frmRegister 
   Appearance      =   0  'Ïëîñêà
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register your copy of Visual SX++"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Register"
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
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1440
   End
   Begin VB.TextBox txtCompany 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   1695
      Width           =   3105
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   0
      Top             =   990
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   30
      Picture         =   "frmRegister.frx":0000
      Top             =   15
      Width           =   3525
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   -15
      Picture         =   "frmRegister.frx":8D4E
      Top             =   2145
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Ïðîçðà÷íî
      Caption         =   "Company :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   165
      TabIndex        =   3
      Top             =   1485
      Width           =   3540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Ïðîçðà÷íî
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   3090
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As Boolean
Private Sub cmdRegister_Click()
If txtCompany.Text <> vbNullString And txtName.Text <> vbNullString Then
SaveSetting "SX++ IDE", "ENV", "BGCOLOR", vbWhite
SaveSetting "SX++ IDE", "ENV", "TEXTCOLOR", vbBlue
SaveSetting "SX++ IDE", "ENV", "FONT", "Courier New"
SaveSetting "SX++ IDE", "ENV", "FONTSIZE", "11"
SaveSetting "SX++ IDE", "REG", "NAME", txtName.Text
SaveSetting "SX++ IDE", "REG", "COMPANY", txtCompany.Text
e = True
frmMain.Caption = "SX++ IDE . Registered to : " & GetSetting("SX++ IDE", "REG", "NAME") & " (" & GetSetting("SX++ IDE", "REG", "COMPANY") & ")"
frmMain.Show
Unload Me
Else
End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not e Then End
End Sub
