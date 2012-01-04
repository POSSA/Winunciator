VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Winunciator 
   BackColor       =   &H80000014&
   Caption         =   "Superfecta Winunciator"
   ClientHeight    =   915
   ClientLeft      =   1380
   ClientTop       =   1920
   ClientWidth     =   4440
   ClipControls    =   0   'False
   Icon            =   "Winunciator.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   915
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Volume"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1020
      TabIndex        =   4
      Top             =   420
      Width           =   2025
      Begin MSComctlLib.Slider volume 
         Height          =   165
         Left            =   30
         TabIndex        =   5
         ToolTipText     =   "Volume"
         Top             =   210
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   291
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   10
         Value           =   10
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Minimize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Minimize to System Tray"
      Top             =   570
      Width           =   1065
   End
   Begin VB.CheckBox Debugcheck 
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Enable and Disable Debug Screen"
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox CNUM 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   " "
      ToolTipText     =   "Caller ID Number"
      Top             =   0
      Width           =   2235
   End
   Begin VB.TextBox CNAM 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   " "
      ToolTipText     =   "Caller ID Name"
      Top             =   0
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   3090
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu RepeatLastAnnouncment 
      Caption         =   "&Repeat"
   End
   Begin VB.Menu Configure 
      Caption         =   "&Configure"
   End
   Begin VB.Menu Update 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Winunciator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents voice As SpVoice
Attribute voice.VB_VarHelpID = -1
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Dim mon As String
Dim nid As NOTIFYICONDATA ' trayicon variable

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


Private Sub About_Click()
frmAbout.Show
End Sub


Private Sub Command1_Click()
minimize_to_tray

End Sub

Sub minimize_to_tray()
Me.Hide
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon ' the icon will be your Form1 project icon
'nid.szTip = CNAM.Text & " " & CNUM.Text & vbNullChar
nid.szTip = "Winunciator v " & App.Major & "." & App.Minor & "." & App.Revision & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub


Private Sub Form_Resize()
    ResizeControls
    If Winunciator.Debugcheck.Value = 1 Then logit ("Saving Winunciator Size and Location..")
    SaveSetting "Winunciator", "Properties", "Wheight", Me.Height
    SaveSetting "Winunciator", "Properties", "Wwidth", Me.Width
    SaveSetting "Winunciator", "Properties", "Wtop", Me.Top
    SaveSetting "Winunciator", "Properties", "Wleft", Me.Left
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
 
If Winunciator.Debugcheck.Value = 1 Then logit ("Saving Winunciator Size and Location..")
SaveSetting "Winunciator", "Properties", "Wheight", Me.Height
SaveSetting "Winunciator", "Properties", "Wwidth", Me.Width
SaveSetting "Winunciator", "Properties", "Wtop", Me.Top
SaveSetting "Winunciator", "Properties", "Wleft", Me.Left

  tcpServer.Close
  voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
   voice.Speak "Win-unciator is offline"
    

Shell_NotifyIcon NIM_DELETE, nid ' del tray icon

   End
End Sub
Private Sub Form_Load()
'On Error Resume Next
SaveSizes
On Error Resume Next
Me.Height = GetSetting("Winunciator", "Properties", "Wheight")
Me.Width = GetSetting("Winunciator", "Properties", "Wwidth")
Me.Top = GetSetting("Winunciator", "Properties", "Wtop")
Me.Left = GetSetting("Winunciator", "Properties", "Wleft")
Me.Show
Me.Refresh
volume.Refresh
'Winunciator.Refresh
'On Error GoTo 0
'SaveSizes

Debugcheck.Value = GetSetting("Winunciator", "Properties", "Debug")
volume.Value = GetSetting("Winunciator", "Properties", "Volume")

If Debugcheck.Value = False Then
frmDebug.Hide
Else
frmDebug.Show
End If

 
 On Error GoTo novoice:
 Set voice = New SpVoice
    Dim sBuffer As String
    Dim lSize As Long

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
On Error GoTo 0:

  If App.PrevInstance = True Then GoTo alreadyrunning
    
 On Error GoTo Hell:
 systime = time
 If Debugcheck.Value = 1 Then logit ("Starting to Listen.")
 tcpServer.LocalPort = GetSetting("Winunciator", "Properties", "Port")
 tcpServer.Listen ' set up the TCP server
 On Error GoTo 0:
    
    
      
If GetSetting("Winunciator", "Properties", "StartMinimized") = 1 Then Call minimize_to_tray
voice.volume = (GetSetting("Winunciator", "Properties", "Volume") * 10)
voice.Speak "Win-unciator is online"

    
 Exit Sub
Hell:
MsgBox "There is a problem with the port you selected. It is already in use. Is the Winunciator already running?"
Resume Next
novoice:
MsgBox "There is a problem with your sound hardware, or Windows speech synthesis."
Resume Next
alreadyrunning:
voice.Speak "There is a problem. The Win-unciator is already running. Only one copy of the win-unciator may be run at a time."
MsgBox "There is a problem. The Winunciator is already running. Only one copy of the Winunciator may be run at a time."
End
    
End Sub



'---------------------------------------------------
'-- Tray icon actions when mouse click on it, etc --
'---------------------------------------------------
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
Dim sFilter As String
msg = X / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
Me.Show ' show form
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Me.Show
Shell_NotifyIcon NIM_DELETE, nid
Case WM_RBUTTONDBLCLK
End Select
End Sub


Private Sub Configure_Click()
frmConfig.Show

End Sub

Private Sub Debugcheck_Click()
SaveSetting "Winunciator", "Properties", "Debug", Debugcheck.Value
If Debugcheck.Value = False Then
frmDebug.Hide
Else
frmDebug.Show
End If
If Debugcheck.Value = 1 Then
logit ("Started Debug Mode.")
Else
logit ("Stopped Debug Mode.")
End If
End Sub




Private Sub Speek(timestorepeat)

invalidcids = Split(frmConfig.invalidcids.Text, ";")
For X = 0 To UBound(invalidcids)
If Debugcheck.Value = 1 Then logit ("Invalid CID # " & X + 1 & " " & invalidcids(X))
If LCase(CNAM.Text) = LCase(invalidcids(X)) Then CNAM.Text = ""
Next
CNAM.Refresh
CNUM.Refresh

If (CNAM.Text = "" And (Len(Winunciator.CNUM.Text) < 1)) Then
Sayit = GetSetting("Winunciator", "Properties", "NoInfoText") & ", " & " "

ElseIf (GetSetting("Winunciator", "Properties", "Conditions") = "Speak Name Only (If Available)" And CNAM.Text <> "") Then
Sayit = GetSetting("Winunciator", "Properties", "PreambleText") & CNAM.Text & ", " & " "

Else
Sayit = GetSetting("Winunciator", "Properties", "PreambleText") & CNAM.Text & ", " & speaknum
End If

nid.szTip = CNAM.Text & " " & CNUM.Text & vbNullChar
Winunciator.Refresh
Winunciator.volume.Refresh
DoEvents

On Error GoTo Hell:
If Debugcheck.Value = 1 Then logit (Sayit)
Winunciator.Refresh
Winunciator.volume.Refresh
DoEvents
voice.volume = volume.Value * 10
voice.Speak Sayit

If val(timestorepeat) > 0 Then
    DoEvents
    timesrepeatrd = 0
    Do While (timesrepeated < val(timestorepeat))
    If Debugcheck.Value = 1 Then logit ("Saying Loop # " & (timesrepeated + 1))
    Sayit = GetSetting("Winunciator", "Properties", "repeatText")
    voice.Speak Sayit
    
    If (CNAM.Text = "" And CNUM.Text = "") Then
    Sayit = GetSetting("Winunciator", "Properties", "NoInfoText") & ", " & " "

    ElseIf (GetSetting("Winunciator", "Properties", "Conditions") = "Speak Name Only (If Available)" And CNAM.Text <> "") Then
    Sayit = GetSetting("Winunciator", "Properties", "PreambleText") & CNAM.Text & ", " & " "

    Else
    Sayit = GetSetting("Winunciator", "Properties", "PreambleText") & CNAM.Text & ", " & speaknum
    End If
    voice.Speak Sayit
    timesrepeated = timesrepeated + 1
    DoEvents
    Winunciator.Refresh
    Winunciator.volume.Refresh
    DoEvents
    Loop
End If

Exit Sub
Hell:
If Debugcheck.Value = 1 Then logit ("There is aproblem with your sound hardware, or Windows speech synthesis.")
MsgBox "There is a problem with your sound hardware, or Windows speech synthesis."

End Sub


Private Sub repeatLastAnnouncment_Click()
Call Speek(0)
End Sub


''''''''''''''''''''''''
''' SERVER FUNCTIONS '''
''''''''''''''''''''''''
Private Sub tcpServer_Close()
   If Debugcheck.Value = 1 Then logit ("socket closed/reopened")
   If tcpServer.State <> sckClosed Then
      tcpServer.Close
   End If
   tcpServer.Listen ' reenable for next connection
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
   If Debugcheck.Value = 1 Then
   logit ("socket connection request")
   End If
   If tcpServer.State <> sckClosed Then
      tcpServer.Close
   End If
   tcpServer.Accept requestID
   If Debugcheck.Value = 1 Then
   logit ("socket connected to " & tcpServer.RemoteHostIP & ":" & tcpServer.RemotePort)
   End If
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
   Dim val As String
   tcpServer.GetData val
   If Debugcheck.Value = 1 Then
'   logit ("socket: " & val)
   End If
   Call GetFile(val)
End Sub

Public Sub GetFile(Request As String)
Dim unpFile As String
Dim unpFileS() As String
authenticatedata (Request)
If datavalid Then
unpFile = Right(Request, Len(Request) - 5)
unpFileS = Split(unpFile, " HTTP")
callerdata = unpFileS(0)


Call distribute(callerdata)
End If
End Sub

Sub authenticatedata(callerinfo)
datavalid = True
'If Left$(callerinfo) <> "{123}" Then datavalid = False
If InStr(callerinfo, "|") < 1 Then datavalid = False
End Sub


Sub distribute(callerdata)

If Debugcheck.Value = 1 Then logit ("Received: " & callerdata)
splitdata = Split(callerdata, "|")

CNAM.Text = splitdata(0)
CNUM.Text = splitdata(1)
formatdata (CNUM.Text)
Winunciator.Refresh

Call Speek(GetSetting("Winunciator", "Properties", "repeats"))

End Sub


Private Sub tcpServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   If Debugcheck.Value = 1 Then
   logit ("socket error: " & Source & ":" & Description)
   End If
End Sub

Private Sub ConductUpdate()
DoEvents: DoEvents
updateprog = GetSetting("Winunciator", "Properties", "appath")
If GetSetting("Winunciator", "Properties", "updatetext") = "Auto Update with Progress Bar" Then
Shell (updateprog & "setup.exe")
ElseIf GetSetting("Winunciator", "Properties", "updatetext") = "Silent Auto Update" Then
Shell (updateprog & "setup.exe /SILENT ")
End If
DoEvents: DoEvents: DoEvents
End
End Sub

Private Sub Update_Click()
' 1. look for new version
'    a. if silent just look, if not silent, show update dialog box.
If GetSetting("Winunciator", "Properties", "updatetext") <> "Silent Auto Update" Then updater.Show Else updater.Hide
' 2. download if present
' 3. Considr Update
'   a. If
End Sub

Private Sub volume_Change()
SaveSetting "Winunciator", "Properties", "Volume", volume.Value
End Sub

''
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


