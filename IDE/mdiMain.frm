VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "Visual SX++ 2.0.17"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8925
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Align           =   2  'Ïðèâÿçàòü âíèç
      Appearance      =   0  'Ïëîñêà
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   0
      ScaleHeight     =   1290
      ScaleWidth      =   8895
      TabIndex        =   30
      Top             =   5505
      Width           =   8925
      Begin VB.TextBox txtDebug 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'Íåò
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   300
         Width           =   8880
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Ôèêñèðîâàíî îäèí
         Caption         =   "..: Debug Window :.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   15
         Width           =   3315
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Ïðèâÿçàòü ââåðõ
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MakeEXE"
            Object.ToolTipText     =   "Make EXE"
            ImageKey        =   "Macro"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ToolBox"
            Object.ToolTipText     =   "Tool box"
            ImageKey        =   "Button"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.OptionButton Option2 
         Caption         =   "Application"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5190
         TabIndex        =   28
         Top             =   75
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4395
         TabIndex        =   27
         Top             =   75
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Ïðèâÿçàòü âïðàâî
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Íåò
      Height          =   5085
      Left            =   6525
      ScaleHeight     =   5085
      ScaleWidth      =   2400
      TabIndex        =   6
      Top             =   420
      Width           =   2400
      Begin VB.TextBox txtWndLeft 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   900
         TabIndex        =   38
         Top             =   3090
         Width           =   1485
      End
      Begin VB.TextBox txtWndTop 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   37
         Top             =   3405
         Width           =   1485
      End
      Begin VB.TextBox txtWndHeight 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   36
         Top             =   3705
         Width           =   1485
      End
      Begin VB.TextBox txtWndWidth 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   35
         Top             =   4005
         Width           =   1485
      End
      Begin VB.TextBox txtWndTitle 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   34
         Top             =   4305
         Width           =   1485
      End
      Begin VB.TextBox txtIndex 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   24
         Top             =   2550
         Width           =   1485
      End
      Begin VB.TextBox txtPicture 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   23
         Top             =   2250
         Width           =   1485
      End
      Begin VB.TextBox txtCaption 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   21
         Top             =   1950
         Width           =   1485
      End
      Begin VB.TextBox txtText 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   19
         Top             =   1650
         Width           =   1485
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   17
         Top             =   1350
         Width           =   1485
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   15
         Top             =   1050
         Width           =   1485
      End
      Begin VB.TextBox txtTop 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   13
         Top             =   750
         Width           =   1485
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Ïðàâàÿ ïðèâÿçêà
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   900
         TabIndex        =   11
         Top             =   435
         Width           =   1485
      End
      Begin VB.Line Line3 
         Index           =   11
         X1              =   15
         X2              =   2370
         Y1              =   3405
         Y2              =   3405
      End
      Begin VB.Line Line3 
         Index           =   10
         X1              =   0
         X2              =   2370
         Y1              =   3705
         Y2              =   3705
      End
      Begin VB.Line Line3 
         Index           =   9
         X1              =   0
         X2              =   2370
         Y1              =   4005
         Y2              =   4005
      End
      Begin VB.Line Line3 
         Index           =   8
         X1              =   15
         X2              =   2370
         Y1              =   4305
         Y2              =   4305
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   43
         Top             =   3150
         Width           =   330
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Top"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   42
         Top             =   3450
         Width           =   270
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   41
         Top             =   3750
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   40
         Top             =   4050
         Width           =   465
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   39
         Top             =   4350
         Width           =   345
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Ôèêñèðîâàíî îäèí
         Caption         =   "..: Window :.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   33
         Top             =   2850
         Width           =   2370
      End
      Begin VB.Line Line3 
         Index           =   7
         X1              =   15
         X2              =   2100
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Line Line3 
         Index           =   6
         X1              =   15
         X2              =   2370
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   15
         X2              =   2370
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   0
         X2              =   2370
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   15
         X2              =   2370
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   0
         X2              =   2370
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   0
         X2              =   2370
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   -15
         X2              =   2370
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Index"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   25
         Top             =   2595
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Picture"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   22
         Top             =   2295
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   20
         Top             =   1995
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   18
         Top             =   1695
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   16
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   14
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Top"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   12
         Top             =   795
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   45
         TabIndex        =   10
         Top             =   495
         Width           =   330
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2370
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1155
         TabIndex        =   9
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   15
         TabIndex        =   8
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Ôèêñèðîâàíî îäèí
         Caption         =   "..: Properties :.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   7
         Top             =   -15
         Width           =   2370
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Ïðèâÿçàòü âëåâî
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Íåò
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   420
      Width           =   1095
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   525
         Picture         =   "mdiMain.frx":3F42
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1590
         Width           =   400
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Picture         =   "mdiMain.frx":4218
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1590
         Width           =   400
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   525
         Picture         =   "mdiMain.frx":455A
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   780
         Width           =   400
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   525
         Picture         =   "mdiMain.frx":480C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1185
         Width           =   400
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Picture         =   "mdiMain.frx":4A8A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1185
         Width           =   400
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Picture         =   "mdiMain.frx":4DCC
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1995
         Width           =   405
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   525
         Picture         =   "mdiMain.frx":504A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   375
         Width           =   400
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Picture         =   "mdiMain.frx":538C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   780
         Width           =   400
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Picture         =   "mdiMain.frx":56CE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   375
         Width           =   400
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Ïëîñêà
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Ôèêñèðîâàíî îäèí
         Caption         =   "..: ToolBox :.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   -15
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5925
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5A10
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5B22
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5C34
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5D46
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5E58
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5F6A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":607C
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":618E
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":62A0
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":63B2
            Key             =   "Button"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":64C4
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewCode 
         Caption         =   "New Code"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open code"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save code"
      End
      Begin VB.Menu S4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenEnv 
         Caption         =   "Open Environment"
      End
      Begin VB.Menu mnuFileSaveEnv 
         Caption         =   "Save Environment"
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
      Begin VB.Menu S3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Project"
      Begin VB.Menu mnuProjectRun 
         Caption         =   "Run"
      End
      Begin VB.Menu S5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectMakeEXE 
         Caption         =   "Make EXE"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
picIndex = picIndex + 1
CopyControl frmWindow.Picture1, "", True, 100, 100, 700, 300, "", (picIndex)
End Sub
Private Sub Command2_Click()
cmdIndex = cmdIndex + 1
CopyControl frmWindow.Button, "New Button", True, 100, 100, 700, 300, "", (cmdIndex)
End Sub
Private Sub Command3_Click()
modMain.txtIndex = modMain.txtIndex + 1
CopyControl frmWindow.Text1, "", True, 100, 100, 700, 300, "New Text", (modMain.txtIndex)
End Sub
Private Sub Command4_Click()
lblIndex = lblIndex + 1
CopyControl frmWindow.Label1, "New Label", True, 100, 100, 700, 300, "", (lblIndex)
End Sub
Private Sub Command5_Click()
On Error Resume Next
Unload frmWindow.Resize1.BoundControl
Set frmWindow.Resize1.BoundControl = n
End Sub

Private Sub Command6_Click()
modMain.txfIndex = modMain.txfIndex + 1
CopyControl frmWindow.TextField1, "", True, 100, 100, 3000, 2000, "New TextField", (modMain.txfIndex)
End Sub

Private Sub Command7_Click()
drvIndex = drvIndex + 1
CopyControl frmWindow.Drive1, "", True, 100, 100, 1590, 270, "", (drvIndex)
End Sub

Private Sub Command8_Click()
filIndex = filIndex + 1
CopyControl frmWindow.File1, "", True, 100, 100, 2000, 1000, "", (filIndex)
End Sub

Private Sub Command9_Click()
dirIndex = dirIndex + 1
CopyControl frmWindow.Dir1, "", True, 100, 100, 2000, 1000, "", (dirIndex)
End Sub

Private Sub MDIForm_Load()
frmWindow.Show
frmMain.Hide
End Sub

Private Sub mnuFileAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileNewCode_Click()
            frmMain.txtCode.Text = "//SX++ IDE - New code"
            Unload frmWindow
            Load frmWindow
End Sub

Private Sub mnuFileOpen_Click()
            frmMain.CM.FileName = "*.SX*"
            frmMain.CM.ShowOpen
            fFile = frmMain.CM.FileName
            fFile = FreeFile
            If frmMain.CM.FileName <> "" And frmMain.CM.FileName <> UCase("*.SX*") Then
            Open frmMain.CM.FileName For Input As fFile
            frmMain.txtCode.Text = Input(LOF(fFile), fFile)
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
            frmMain.CM.FileName = "*.SX*"
            frmMain.CM.ShowOpen
            If frmMain.CM.FileName <> "" Then
            fFile = FreeFile
            Open frmMain.CM.FileName For Output As fFile
            Print #fFile, frmMain.txtCode.Text
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

Private Sub mnuProjectMakeEXE_Click()
On Error Resume Next
            frmMain.CM.ShowOpen
            txtDebug.Text = ""
            MakeEXE CM.FileName
End Sub

Private Sub mnuProjectRun_Click()
            MakeEXE "C:\TMP_SX++_APP.EXE"
            Shell "C:\TMP_SX++_APP.EXE", vbNormalFocus
End Sub

Private Sub Option1_Click()
frmWindow.Hide
frmMain.Show
frmMain.WindowState = 2
End Sub
Private Sub Option2_Click()
frmWindow.Show
frmMain.Hide
End Sub

Private Sub Text3_Change()
On Error Resume Next
frmWindow.Height = Text3.Text
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            frmMain.txtCode.Text = "//SX++ IDE - New code"
            Unload frmWindow
            Load frmWindow
        Case "Open"
            frmMain.CM.FileName = "*.SX*"
            frmMain.CM.ShowOpen
            fFile = frmMain.CM.FileName
            fFile = FreeFile
            If frmMain.CM.FileName <> "" And frmMain.CM.FileName <> UCase("*.SX*") Then
            Open frmMain.CM.FileName For Input As fFile
            frmMain.txtCode.Text = Input(LOF(fFile), fFile)
            End If
        Case "Save"
            frmMain.CM.FileName = "*.SX*"
            frmMain.CM.ShowOpen
            If frmMain.CM.FileName <> "" Then
            fFile = FreeFile
            Open frmMain.CM.FileName For Output As fFile
            Print #fFile, frmMain.txtCode.Text
            Close fFile
            End If
        Case "Copy"
            On Error Resume Next
            Clipboard.SetText frmMain.txtCode.SelText
        Case "Paste"
            On Error Resume Next
            txtCode.SelText = Clipboard.GetText
        Case "Delete"
            On Error Resume Next
            txtCode.SelText = ""
        Case "Run"
            MakeEXE "C:\TMP_SX++_APP.EXE"
            Shell "C:\TMP_SX++_APP.EXE", vbNormalFocus
        Case "MakeEXE"
            frmMain.CM.ShowOpen
            txtDebug.Text = ""
            MakeEXE CM.FileName
        Case "Properties"
            If Picture2.Visible = True Then
            Picture2.Visible = False
            Else
            Picture2.Visible = True
            End If
        Case "ToolBox"
            If Picture1.Visible = True Then
            Picture1.Visible = False
            Else
            Picture1.Visible = True
            End If
        Case "Help"
            frmAbout.Show
    End Select
End Sub
Private Sub txtCaption_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Caption = txtCaption.Text
End Sub
Private Sub txtHeight_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Height = txtHeight.Text
End Sub
Private Sub txtLeft_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Left = txtLeft.Text
End Sub
Private Sub txtPicture_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Picture = LoadPicture(txtPicture.Text)
End Sub
Private Sub txtText_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Text = txtText.Text
End Sub
Private Sub txtTop_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Top = txtTop.Text
End Sub
Private Sub txtWidth_Change()
On Error Resume Next
frmWindow.Resize1.BoundControl.Width = txtWidth.Text
End Sub

Private Sub txtWndLeft_Change()
On Error Resume Next
frmWindow.Left = txtWndLeft.Text
End Sub

Private Sub txtWndTitle_Change()
On Error Resume Next
frmWindow.Caption = txtWndTitle.Text
End Sub

Private Sub txtWndTop_Change()
On Error Resume Next
frmWindow.Top = txtWndTop.Text
End Sub

Private Sub txtWndWidth_Change()
On Error Resume Next
frmWindow.Width = txtWndWidth.Text
End Sub
