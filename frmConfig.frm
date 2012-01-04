VERSION 5.00
Begin VB.Form frmConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "Winunciator Configuration"
   ClientHeight    =   4680
   ClientLeft      =   8430
   ClientTop       =   1890
   ClientWidth     =   4320
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   4320
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   60
      TabIndex        =   25
      Top             =   3390
      Width           =   4245
      Begin VB.CommandButton resetwindows 
         Caption         =   "&Reset Winunciator Window Size"
         Height          =   345
         Left            =   330
         TabIndex        =   26
         ToolTipText     =   "Reset Default Winunciator Screen Size and Shape"
         Top             =   210
         Width           =   3315
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   60
      TabIndex        =   18
      Top             =   3990
      Width           =   4245
      Begin VB.CommandButton Save 
         Caption         =   "&Save"
         Height          =   345
         Left            =   3420
         TabIndex        =   21
         ToolTipText     =   "Save COnfiguration Settings"
         Top             =   210
         Width           =   675
      End
      Begin VB.CommandButton speaksettings 
         Caption         =   "S&peak Settings"
         Height          =   345
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Speak Configuration Settings"
         Top             =   210
         Width           =   1335
      End
      Begin VB.CheckBox startminimized 
         Caption         =   "Start In Task Bar"
         Height          =   315
         Left            =   1650
         MaskColor       =   &H000000FF&
         TabIndex        =   19
         ToolTipText     =   "Enable / Disable Starting Winunciator in the System Tray"
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame speechframe 
      Caption         =   "Speach Settings"
      Height          =   2295
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   4275
      Begin VB.TextBox invalidcids 
         Height          =   285
         Left            =   990
         TabIndex        =   27
         ToolTipText     =   "List of Caller ID Names to be considered invalid.  Semi colon (;) seperator bewteen items."
         Top             =   1890
         Width           =   3195
      End
      Begin VB.ComboBox format 
         Height          =   315
         ItemData        =   "frmConfig.frx":0442
         Left            =   1200
         List            =   "frmConfig.frx":044F
         TabIndex        =   10
         Text            =   "Select"
         ToolTipText     =   "Select Number Pattern"
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox portvalue 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Port Listening On"
         Top             =   270
         Width           =   705
      End
      Begin VB.TextBox repeattext 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         ToolTipText     =   "Text Spoken between Repeats"
         Top             =   630
         Width           =   2085
      End
      Begin VB.TextBox preambletext 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         ToolTipText     =   "Text Spoken Before Announcement"
         Top             =   930
         Width           =   3345
      End
      Begin VB.ComboBox repeats 
         Height          =   315
         ItemData        =   "frmConfig.frx":0479
         Left            =   3480
         List            =   "frmConfig.frx":0489
         TabIndex        =   6
         ToolTipText     =   "Select or Key In number of times to repeat announcement."
         Top             =   600
         Width           =   765
      End
      Begin VB.ComboBox conditions 
         Height          =   315
         ItemData        =   "frmConfig.frx":0499
         Left            =   840
         List            =   "frmConfig.frx":04A3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Configure How and When Caller ID Name and Caller ID Number are Spoken."
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox noinfo 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "Text Spoken if no Caller ID Name or Number is Available. (Leave Blank for No Announcement if Data unavailable.)"
         Top             =   1230
         Width           =   3345
      End
      Begin VB.Label Label9 
         Caption         =   "Invalid CIDs"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         ToolTipText     =   "List of Caller ID Names to be considered invalid.  Semi colon (;) seperator bewteen items."
         Top             =   1950
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Number Format"
         Height          =   255
         Left            =   -150
         TabIndex        =   17
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label portlabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   255
         Left            =   3060
         TabIndex        =   16
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Preamble"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Repete"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Times"
         Height          =   255
         Left            =   2970
         TabIndex        =   13
         Top             =   660
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "No Info:"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   1290
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Conditions:"
         Height          =   225
         Left            =   60
         TabIndex        =   11
         Top             =   1590
         Width           =   735
      End
   End
   Begin VB.Frame updaterframe 
      Caption         =   "Program Update Settings"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   2310
      Width           =   4245
      Begin VB.ComboBox updateschedule 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmConfig.frx":04DC
         Left            =   870
         List            =   "frmConfig.frx":04EC
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Configure When and How to Update Winunciator"
         Top             =   630
         Width           =   3315
      End
      Begin VB.ComboBox update 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmConfig.frx":0537
         Left            =   780
         List            =   "frmConfig.frx":0547
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Configure When and How to Update Winunciator"
         Top             =   270
         Width           =   3435
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Configured for Manual Updates"
         Height          =   345
         Left            =   900
         TabIndex        =   24
         Top             =   690
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule"
         Enabled         =   0   'False
         Height          =   225
         Left            =   60
         TabIndex        =   23
         Top             =   690
         Width           =   705
      End
      Begin VB.Label updateconfig 
         Alignment       =   1  'Right Justify
         Caption         =   "Updates:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   60
         TabIndex        =   1
         Top             =   330
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents voice As SpVoice
Attribute voice.VB_VarHelpID = -1
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Dim mon As String
 
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



Private Sub Form_Load()
SaveSizes
Set voice = New SpVoice
    Dim sBuffer As String
    Dim lSize As Long


    sBuffer = Space$(255)
    lSize = Len(sBuffer)

If Winunciator.Debugcheck.Value = 1 Then logit ("Starting Configure.")
On Error GoTo 0
portvalue.Text = (GetSetting("Winunciator", "Properties", "Port"))
format.Text = (GetSetting("Winunciator", "Properties", "Format"))
repeats.Text = (GetSetting("Winunciator", "Properties", "Repeats"))
repeattext.Text = (GetSetting("Winunciator", "Properties", "RepeatText"))
preambletext.Text = (GetSetting("Winunciator", "Properties", "PreambleText"))
frmConfig.startminimized.Value = (GetSetting("Winunciator", "Properties", "StartMinimized"))
frmConfig.noinfo.Text = GetSetting("Winunciator", "Properties", "NoInfoText")
frmConfig.conditions.Text = GetSetting("Winunciator", "Properties", "Conditions")
frmConfig.Update.Text = GetSetting("Winunciator", "Properties", "UpdateText")
frmConfig.updateschedule.Text = GetSetting("Winunciator", "Properties", "UpdateSchedule")
frmConfig.invalidcids.Text = GetSetting("Winunciator", "Properties", "InvalidCIDS")

End Sub


Private Sub Form_Resize()
    ResizeControls
End Sub





Private Sub resetwindows_Click()
 If Winunciator.Debugcheck.Value = 1 Then logit ("Restoring Default Winunciator Size and Location.")
    SaveSetting "Winunciator", "Properties", "Wheight", "1725"
    SaveSetting "Winunciator", "Properties", "Wwidth", "4560"
    SaveSetting "Winunciator", "Properties", "Wtop", "1170"
    SaveSetting "Winunciator", "Properties", "Wleft", "1320"
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
   End
End Sub

Private Sub Save_Click()
If Winunciator.Debugcheck.Value = 1 Then logit ("Saving Configure.")
SaveSetting "Winunciator", "Properties", "Format", format.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved Format.")
SaveSetting "Winunciator", "Properties", "repeats", repeats.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved repeat Count.")
SaveSetting "Winunciator", "Properties", "Port", portvalue.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved Port.")
SaveSetting "Winunciator", "Properties", "repeatText", repeattext.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved repeat Text.")
SaveSetting "Winunciator", "Properties", "PreambleText", preambletext.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved Preamble Text.")
SaveSetting "Winunciator", "Properties", "NoInfoText", noinfo.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved `No Info` Text.")
SaveSetting "Winunciator", "Properties", "Conditions", conditions.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved `Conditions` Text.")
SaveSetting "Winunciator", "Properties", "InvalidCIDS", invalidcids.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved `Invalid CIDS` Text.")

'debug settings
SaveSetting "Winunciator", "Properties", "UpdateText", Update.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved Update Configuration.")
SaveSetting "Winunciator", "Properties", "UpdateSchedule", updateschedule.Text
If Winunciator.Debugcheck.Value = 1 Then logit ("Saved Update Schedule.")

Sayit = "Win unciator configuration settings have been saved."
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
voice.Speak Sayit


If Winunciator.Debugcheck.Value = 1 Then logit ("Changing listen port to: " & portvalue.Text)
frmDebug.Refresh
SaveSetting "Winunciator", "Properties", "Port", portvalue.Text
'On Error GoTo 0
Winunciator.tcpServer.Close
On Error GoTo Hell:
Winunciator.tcpServer.LocalPort = portvalue.Text
portvalue.Text = Winunciator.tcpServer.LocalPort
Winunciator.tcpServer.Listen ' set up the TCP server


Exit Sub
Hell:
MsgBox "There is a problem with the port you selected - it is already in use."
End Sub

Private Sub speaksettings_Click()
If Winunciator.Debugcheck.Value = 1 Then logit ("Speaking current config.")
On Error GoTo Hell:
voice.volume = Winunciator.volume.Value * 10

Sayit = "Win unciator version: " & App.Major & "." & App.Minor & "." & App.Revision
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
voice.Speak Sayit

Sayit = "Using Port: " & portvalue.Text
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
voice.Speak Sayit

Sayit = "Announcement Will be repeated: " & repeats.Text & " Times."
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
voice.Speak Sayit


Sayit = "Repeat Text is... " & repeattext.Text
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)


Sayit = "Pre amble text is... " & preambletext.Text
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)

Sayit = "For Calls without name or number, Win-unciator will "
If GetSetting("Winunciator", "Properties", "NoInfoText") = "" Then Sayit = Sayit & "Remain SIlent," Else Sayit = Sayit & "Say: " & GetSetting("Winunciator", "Properties", "NoInfoText")
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)

Sayit = "The Speach Rules are configured to... " & GetSetting("Winunciator", "Properties", "Conditions")
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)

Sayit = "The Following words are configured to be recognized as invalid Caller eye-dees... "
voice.Speak Sayit
DoEvents
DoEvents
invalidcids2 = GetSetting("Winunciator", "Properties", "InvalidCIDS")
invalidcids3 = Split(invalidcids2, ";")
For X = 0 To UBound(invalidcids3)
Sayit = invalidcids3(X)
voice.Speak Sayit
DoEvents: DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
Next
DoEvents: DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)

Sayit = "The Update Rules are configured to... " & GetSetting("Winunciator", "Properties", "UpdateText")
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)

Sayit = "The Update schedule is configured to... "
If Update.Text = "Manual Update Only" Then Sayit = Sayit & "Manual." Else Sayit = Sayit & GetSetting("Winunciator", "Properties", "UpdateSchedule")
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)

If updaterframe.Enabled = False Then Sayit = "The Auto Updater Subsystem is currently DISABLED."
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)


If GetSetting("Winunciator", "Properties", "StartMinimized") = "0" Then Sayit = "Start Win unciator in Task bar mode is disabled." Else Sayit = "Start Win unciator in minimized mode is enabled."
voice.Speak Sayit
DoEvents
DoEvents
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)



Exit Sub
Hell:
If Winunciator.Debugcheck.Value = 1 Then logit ("There is a problem with your sound hardware, or Windows speech synthesis.")
MsgBox "There is aproblem with your sound hardware, or Windows speech synthesis."
On Error GoTo 0
End Sub

Private Sub startminimized_Click()
SaveSetting "Winunciator", "Properties", "StartMinimized", startminimized.Value
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

Private Sub update_LostFocus()
If Update.Text = "Manual Update Only" Then
frmConfig.updateschedule.Visible = False
updateschedule.Refresh
Else
frmConfig.updateschedule.Visible = True
updateschedule.Refresh
End If
frmConfig.Refresh
End Sub

