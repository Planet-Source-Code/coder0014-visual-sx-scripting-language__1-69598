VERSION 5.00
Begin VB.UserControl Resize 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Íåïðîçðà÷íî
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFFC0&
      FillStyle       =   4  'Äèàãîíàëü ââåðõ
      Height          =   3375
      Left            =   105
      Top             =   105
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Ïëîñêà
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Ôèêñèðîâàíî îäèí
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "Resize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_MoveMode As Boolean
Dim m_BoundControl As Object
Dim XX As Single
Dim YY As Single

Dim SX As Single
Dim SY As Single

Const m_def_MoveMode = 0

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Resize()
Event Moving()


Private Sub ChangeMoveState()
    If m_MoveMode = False Then
    m_MoveMode = True
    Else
    m_MoveMode = False
    End If
    Shape1.Visible = m_MoveMode
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SX = X
    YY = Y
    If Button = 2 Then ChangeMoveState
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then ResizeBoundControl Index, X, Y
End Sub

Private Sub UserControl_Initialize()

Label1(0).Move 0, 0, 100, 100

Dim i As Integer
MoveGrips 0
For i = 1 To 7
    Load Label1(i)
    MoveGrips i
    Label1(i).Visible = True
     Label1(i).ZOrder 0
Next i
SetControlOnTop
End Sub

Private Sub MoveGrips(i As Integer)
On Error Resume Next
Select Case i
Case Is = 0 'top left
    Label1(i).Move 0, 0
    Label1(i).MousePointer = 8
Case Is = 1 'top middle
    Label1(i).Move (UserControl.Width / 2) - (Label1(i).Width / 2), 0
    Label1(i).MousePointer = 7
Case Is = 2 'top right
    Label1(i).Move UserControl.Width - Label1(i).Width, 0
    Label1(i).MousePointer = 6
Case Is = 3 'middle left
    Label1(i).Move 0, (UserControl.Height / 2) - (Label1(i).Height / 2)
    Label1(i).MousePointer = 9
Case Is = 4 'middle right
    Label1(i).Move UserControl.Width - Label1(i).Width, (UserControl.Height / 2) - (Label1(i).Height / 2)
    Label1(i).MousePointer = 9
Case Is = 5 'bottom left
    Label1(i).Move 0, UserControl.Height - Label1(i).Height
    Label1(i).MousePointer = 6
Case Is = 6 'bottom middle
    Label1(i).Move (UserControl.Width / 2) - (Label1(i).Width / 2), UserControl.Height - Label1(i).Height
    Label1(i).MousePointer = 7
Case Is = 7 ' bottom right
    Label1(i).Move UserControl.Width - Label1(i).Width, UserControl.Height - Label1(i).Height
    Label1(i).MousePointer = 8
End Select
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    XX = X
    YY = Y
    If Button = 1 Then MoveBoundControl X, Y
    If Button = 2 Then ChangeMoveState
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button = 1 Then MoveBoundControl X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then MoveBoundControl X, Y
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim i As Integer
For i = 1 To 8
    MoveGrips i
Next i
Shape1.Move Label1(0).Width, Label1(0).Height, UserControl.Width - (Label1(0).Width * 2), UserControl.Height - (Label1(0).Height * 2)
End Sub

Public Property Get BoundControl() As Object
    Set BoundControl = m_BoundControl
End Property

Public Property Set BoundControl(ByVal New_BoundControl As Object)
    
    Set m_BoundControl = New_BoundControl
    PropertyChanged "BoundControl"

    If New_BoundControl Is Nothing Then
    
    Else
    SetSameContainer
    MoveUserControl
    End If
    
End Property

Private Sub SetSameContainer()
Dim i As Long
Dim OBJ As Object
On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hWnd = UserControl.ParentControls.Item(i).hWnd Then
    Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next i
If m_BoundControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub
    If m_BoundControl.Container <> OBJ.Container Then
        Set OBJ.Container = m_BoundControl.Container
    End If
End Sub
Private Sub MoveUserControl()
Dim i As Long
Dim OBJ As Object
On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hWnd = UserControl.ParentControls.Item(i).hWnd Then
    Set OBJ = UserControl.ParentControls.Item(i)
        OBJ.Move m_BoundControl.Left - offset, m_BoundControl.Top - offset, m_BoundControl.Width + (offset * 2), m_BoundControl.Height + (offset * 2)
        Exit For
    End If
Next i

End Sub

Private Sub SetControlOnTop()
Dim i As Long
Dim OBJ As Object

On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1
    If UserControl.hWnd = UserControl.ParentControls.Item(i).hWnd Then
        Set OBJ = UserControl.ParentControls.Item(i)
        OBJ.ZOrder 0
        Exit For
    End If
Next i


End Sub
Private Sub ResizeBoundControl(ID As Integer, X, Y)
Dim i As Long
Dim OBJ As Object

On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1
    If UserControl.hWnd = UserControl.ParentControls.Item(i).hWnd Then
        Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next i

If m_BoundControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub
Select Case ID
Case Is = 0
    OBJ.Top = OBJ.Top + Y
    OBJ.Height = OBJ.Height - Y
    OBJ.Left = OBJ.Left + X
    OBJ.Width = OBJ.Width - X
Case Is = 1
    OBJ.Top = OBJ.Top + Y
    OBJ.Height = OBJ.Height - Y
Case Is = 2
OBJ.Top = OBJ.Top + Y
    OBJ.Height = OBJ.Height - Y
    OBJ.Width = OBJ.Width + (X)
Case Is = 3
    OBJ.Left = OBJ.Left + X
    OBJ.Width = OBJ.Width - X
Case Is = 4
    OBJ.Width = OBJ.Width + (X)
Case Is = 5
    OBJ.Left = OBJ.Left + X
    OBJ.Width = OBJ.Width - X
    OBJ.Height = OBJ.Height + (Y)
Case Is = 6
    OBJ.Height = OBJ.Height + (Y)
Case Is = 7
    OBJ.Height = OBJ.Height + (Y)
    OBJ.Width = OBJ.Width + (X)
End Select
m_BoundControl.Move OBJ.Left + offset, OBJ.Top + offset, OBJ.Width - (offset * 2), OBJ.Height - (offset * 2)
If OBJ.Width < m_BoundControl.Width Or OBJ.Height < m_BoundControl.Height Then
    OBJ.Move m_BoundControl.Left - offset, m_BoundControl.Top - offset, m_BoundControl.Width + (offset * 2), m_BoundControl.Height + (offset * 2)
End If
End Sub

Private Sub MoveBoundControl(X As Single, Y As Single)
Dim i As Long
Dim OBJ As Object
On Error Resume Next
If m_MoveMode = False Then Exit Sub
For i = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hWnd = UserControl.ParentControls.Item(i).hWnd Then
        Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next i

If m_BoundControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub

    OBJ.Move OBJ.Left + (X - XX), OBJ.Top + (Y - YY)
    m_BoundControl.Move m_BoundControl.Left + (X - XX), m_BoundControl.Top + (Y - YY)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_BoundControl = PropBag.ReadProperty("BoundControl", Nothing)
    m_MoveMode = PropBag.ReadProperty("MoveMode", m_def_MoveMode)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BoundControl", m_BoundControl, Nothing)
    Call PropBag.WriteProperty("MoveMode", m_MoveMode, m_def_MoveMode)
End Sub

Public Property Get offset() As Single
    offset = Label1(0).Width
End Property

Public Property Get MoveMode() As Boolean
    MoveMode = m_MoveMode
End Property

Public Property Let MoveMode(ByVal New_MoveMode As Boolean)
    m_MoveMode = New_MoveMode
    PropertyChanged "MoveMode"
End Property

Private Sub UserControl_InitProperties()
    m_MoveMode = m_def_MoveMode
End Sub


