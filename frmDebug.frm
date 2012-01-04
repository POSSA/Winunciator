VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Winunciator Debug"
   ClientHeight    =   2925
   ClientLeft      =   1380
   ClientTop       =   3675
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR  LOG"
      Height          =   315
      Left            =   30
      TabIndex        =   1
      ToolTipText     =   "Clear Debug Log"
      Top             =   2610
      Width           =   6555
   End
   Begin VB.TextBox log1 
      Height          =   2595
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'resizer stuff
Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single



Private Sub Command1_Click()
log1.Text = ""
logit ("Cleared Log.")
End Sub

Private Sub Form_Load()
SaveSizes
End Sub

''
''   Below here be the resizer codes
''
' Save the form's and controls' dimensions.
Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control

    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
    
        With m_ControlPositions(i)
                        
                 If TypeOf ctl Is Line Then
                .Left = ctl.x1
                .Top = ctl.y1
                .Width = ctl.x2 - ctl.x1
                .Height = ctl.y2 - ctl.y1
             
            ElseIf TypeOf ctl Is Winsock Then
            ElseIf TypeOf ctl Is Menu Then
            Else
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl

    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight
End Sub



' Arrange the controls for the new size.
Private Sub ResizeControls()
Dim i As Integer
Dim ctl As Control
Dim x_scale As Single
Dim y_scale As Single

    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub

    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt

    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                ctl.x1 = x_scale * .Left
                ctl.y1 = y_scale * .Top
                ctl.x2 = ctl.x1 + x_scale * .Width
                ctl.y2 = ctl.y1 + y_scale * .Height
            ElseIf TypeOf ctl Is Winsock Then
            ElseIf TypeOf ctl Is Menu Then
            Else
                ctl.Left = x_scale * .Left
                ctl.Top = y_scale * .Top
                ctl.Width = x_scale * .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = y_scale * .Height
                End If
                On Error Resume Next
                ctl.Font.Size = y_scale * .FontSize
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl
End Sub



Private Sub Form_Resize()
    ResizeControls
End Sub
