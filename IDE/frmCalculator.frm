VERSION 5.00
Begin VB.Form frmCalculator 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SX++ Calculator"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Percent 
      BackColor       =   &H00FF8080&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2025
      Width           =   495
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2025
      Width           =   500
   End
   Begin VB.CommandButton Decimal 
      BackColor       =   &H00FF8080&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2025
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2025
      Width           =   1035
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00FF8080&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1500
      Width           =   495
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00FF8080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1500
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1500
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1500
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1500
      Width           =   500
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00FF8080&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00FF8080&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton CancelEntry 
      BackColor       =   &H00FF8080&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   435
      Width           =   495
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00FF8080&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   435
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   435
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   435
      Width           =   500
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00FF8080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   435
      Width           =   500
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Ôèêñèðîâàíî îäèí
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   -15
      TabIndex        =   19
      Top             =   30
      Width           =   2715
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Op1, Op2
Dim DecimalFlag As Integer
Dim NumOps As Integer
Dim LastInput
Dim OpFlag
Dim TempReadout

Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
End Sub

Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CCur(Op1) + CCur(Op2)
            Case "-"
                Op1 = CCur(Op1) - CCur(Op2)
            Case "X"
                Op1 = CCur(Op1) * CCur(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub
