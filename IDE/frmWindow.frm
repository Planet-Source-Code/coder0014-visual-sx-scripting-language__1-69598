VERSION 5.00
Begin VB.Form frmWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   Caption         =   "My Application"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6555
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
      Height          =   930
      Index           =   0
      Left            =   2100
      TabIndex        =   8
      Top             =   2250
      Visible         =   0   'False
      Width           =   2000
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
      Height          =   930
      Index           =   0
      Left            =   2130
      TabIndex        =   7
      Top             =   1185
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   1000
      Index           =   0
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Îáà
      TabIndex        =   6
      Text            =   "frmWindow.frx":0000
      Top             =   150
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ïëîñêà
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   0
      Left            =   90
      ScaleHeight     =   975
      ScaleWidth      =   1965
      TabIndex        =   5
      Top             =   2190
      Visible         =   0   'False
      Width           =   2000
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
      Height          =   1000
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Text            =   "TextBox"
      Top             =   1185
      Visible         =   0   'False
      Width           =   2000
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
      Height          =   1000
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   2000
   End
   Begin SXPlusPlusIDE.Resize Resize1 
      Height          =   1000
      Left            =   -2700
      TabIndex        =   0
      Top             =   -2700
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   1773
   End
   Begin VB.Image Drive1 
      Height          =   270
      Index           =   0
      Left            =   4200
      Picture         =   "frmWindow.frx":000A
      Stretch         =   -1  'True
      Top             =   1275
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Ïðîçðà÷íî
      Caption         =   "Label"
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
      Height          =   1005
      Index           =   0
      Left            =   2220
      TabIndex        =   4
      Top             =   165
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label n 
      Height          =   1725
      Left            =   4500
      TabIndex        =   2
      Top             =   -2.10060e5
      Visible         =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "frmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblClick As Byte

Private Sub Button_Click(Index As Integer)
dblClick = dblClick + 1
If dblClick = 2 Then
frmMain.Show
frmMain.WindowState = 2
dblClick = 0
If InStr(1, LCase(frmMain.txtCode.Text), "program button" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program button" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program button" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program button" & Index & "_click()")
End If
End If
End Sub

Private Sub Button_GotFocus(Index As Integer)
Set Resize1.BoundControl = Button(Index)
mdiMain.txtCaption.Text = Button(Index).Caption
mdiMain.txtLeft.Text = Button(Index).Left
mdiMain.txtTop.Text = Button(Index).Top
mdiMain.txtWidth.Text = Button(Index).Width
mdiMain.txtHeight.Text = Button(Index).Height
End Sub

Private Sub Dir1_Click(Index As Integer)
dblClick = dblClick + 1
If dblClick = 2 Then
frmMain.Show
frmMain.WindowState = 2
dblClick = 0
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program dirlistbox" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program dirlistbox" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program dirlistbox" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program dirlistbox" & Index & "_click()")
End If
End If
End Sub

Private Sub Dir1_GotFocus(Index As Integer)
Set Resize1.BoundControl = Dir1(Index)
mdiMain.txtLeft.Text = Dir1(Index).Left
mdiMain.txtTop.Text = Dir1(Index).Top
mdiMain.txtWidth.Text = Dir1(Index).Width
mdiMain.txtHeight.Text = Dir1(Index).Height
End Sub


Private Sub Drive1_Click(Index As Integer)
Set Resize1.BoundControl = Drive1(Index)
mdiMain.txtLeft.Text = Drive1(Index).Left
mdiMain.txtTop.Text = Drive1(Index).Top
mdiMain.txtWidth.Text = Drive1(Index).Width
mdiMain.txtHeight.Text = Drive1(Index).Height
End Sub

Private Sub Drive1_DblClick(Index As Integer)
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program drivelistbox" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program drivelistbox" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program drivelistbox" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program drivelistbox" & Index & "_click()")
End If
End Sub

Private Sub File1_DblClick(Index As Integer)
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program filelistbox" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program filelistbox" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program filelistbox" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program filelistbox" & Index & "_click()")
End If
End Sub

Private Sub File1_GotFocus(Index As Integer)
Set Resize1.BoundControl = File1(Index)
mdiMain.txtLeft.Text = File1(Index).Left
mdiMain.txtTop.Text = File1(Index).Top
mdiMain.txtWidth.Text = File1(Index).Width
mdiMain.txtHeight.Text = File1(Index).Height
End Sub

Private Sub Form_DblClick()
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program window_load()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program window_load()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program window_load()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program window_load()")
End If
End Sub

Private Sub Form_Load()
DrawDots
frmWindow.Height = 222 * 15
frmWindow.Width = 333 * 15
frmMain.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Resize1.BoundControl = n
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
mdiMain.txtWndHeight.Text = Me.Height
mdiMain.txtWndLeft.Text = Me.Left
mdiMain.txtWndTitle.Text = Me.Caption
mdiMain.txtWndTop.Text = Me.Top
mdiMain.txtWndWidth.Text = Me.Width
dblClick = 0
frmWindow.WindowState = 0: frmWindow.Left = 0: frmWindow.Top = 0
End Sub

Private Sub Form_Resize()
DrawDots
End Sub

Private Sub Image2_Click()
End Sub

Private Sub Label1_Click(Index As Integer)
Set Resize1.BoundControl = Label1(Index)
mdiMain.txtCaption.Text = Label1(Index).Caption
mdiMain.txtLeft.Text = Label1(Index).Left
mdiMain.txtTop.Text = Label1(Index).Top
mdiMain.txtWidth.Text = Label1(Index).Width
mdiMain.txtHeight.Text = Label1(Index).Height
End Sub

Private Sub Label1_DblClick(Index As Integer)
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program label" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program label" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program label" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program label" & Index & "_click()")
End If
End Sub

Private Sub mnuControlDelete_Click()
If Resize1.BoundControl <> n Then Unload Resize1.BoundControl
Set Resize1.BoundControl = n
End Sub

Private Sub Picture1_DblClick(Index As Integer)
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program picture" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program picture" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program picture" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program picture" & Index & "_click()")
End If
End Sub

Private Sub Picture1_GotFocus(Index As Integer)
Set Resize1.BoundControl = Picture1(Index)
mdiMain.txtCaption.Text = Picture1(Index).Picture
mdiMain.txtLeft.Text = Picture1(Index).Left
mdiMain.txtTop.Text = Picture1(Index).Top
mdiMain.txtWidth.Text = Picture1(Index).Width
mdiMain.txtHeight.Text = Picture1(Index).Height
End Sub



Private Sub Text1_DblClick(Index As Integer)
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program textbox" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program textbox" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program textbox" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program textbox" & Index & "_click()")
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Set Resize1.BoundControl = Text1(Index)
mdiMain.txtText.Text = Text1(Index).Text
mdiMain.txtLeft.Text = Text1(Index).Left
mdiMain.txtTop.Text = Text1(Index).Top
mdiMain.txtWidth.Text = Text1(Index).Width
mdiMain.txtHeight.Text = Text1(Index).Height
End Sub

Private Function CopyControl(Control As Variant, Caption As String, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Text As String)
On Error Resume Next
Dim X As Integer
X = Control.Count + 1
Load Control(X)
With Control(X)
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

Function DrawDots()
For X = 1 To frmWindow.Width Step Screen.TwipsPerPixelX * 10
For Y = 1 To frmWindow.Width Step Screen.TwipsPerPixelY * 10
PSet (X, Y), 16777125
Next Y
Next X
End Function

Private Sub TextField1_DblClick(Index As Integer)
frmMain.Show
frmMain.WindowState = 2
If InStr(1, LCase(frmMain.txtCode.Text), "program textfield" & Index & "_click()") = 0 Then
frmMain.txtCode.Text = frmMain.txtCode.Text & vbCrLf & "program textfield" & Index & "_click()" & vbCrLf & vbCrLf & "endp;"
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program textfield" & Index & "_click()")
Else
frmMain.txtCode.SelStart = InStr(1, LCase(frmMain.txtCode.Text), "program textfield" & Index & "_click()")
End If
End Sub

Private Sub TextField1_GotFocus(Index As Integer)
Set Resize1.BoundControl = TextField1(Index)
mdiMain.txtText.Text = TextField1(Index).Text
mdiMain.txtLeft.Text = TextField1(Index).Left
mdiMain.txtTop.Text = TextField1(Index).Top
mdiMain.txtWidth.Text = TextField1(Index).Width
mdiMain.txtHeight.Text = TextField1(Index).Height
End Sub
