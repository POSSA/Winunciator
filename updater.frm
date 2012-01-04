VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form updater 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Winunciator Auto Updater"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "updater.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2190
      Top             =   3300
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "colsolgrp.com"
      URL             =   "http://colsolgrp.com/rss/winunciator.txt"
      Document        =   "/rss/winunciator.txt"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   975
      Left            =   180
      TabIndex        =   5
      Top             =   3210
      Width           =   1695
      ExtentX         =   2990
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox response 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1290
      Width           =   4005
   End
   Begin MSWinsockLib.Winsock tcpclient 
      Left            =   4320
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3390
      TabIndex        =   2
      Top             =   930
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   570
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label statusreport 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   900
      Width           =   3285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Checking for updates."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4185
   End
End
Attribute VB_Name = "updater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents voice As SpVoice
Attribute voice.VB_VarHelpID = -1
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Dim mon As String

''''''''''''''''''''''''
''' CLIENT FUNCTIONS '''
''''''''''''''''''''''''
Private Sub cmdClientConnect_Click()
  
End Sub
Private Sub cmdClientDisconnect_Click()
   tcpclient.Close
   AddClientHistory "Disconnected"
End Sub

Private Sub txtClientToSend_Change()
   If txtClientToSend.Text <> "" Then ' if there's something entered
      If tcpclient.State = sckConnected Then
         tcpclient.SendData txtClientToSend.Text
         AddClientHistory txtClientToSend.Text
      Else
         AddClientHistory "Not Connected"
      End If
      txtClientToSend.Text = ""
   End If
End Sub


Private Sub Cancel_Click()
'Shell ("winunciator.exe")
End

End Sub

Private Sub tcpClient_Close()
    tcpconnect = False
    If Winunciator.Debugcheck.Value = 1 Then logit ("socket closed")
End Sub

Private Sub tcpClient_Connect()
      tcpconnect = True
      If Winunciator.Debugcheck.Value = 1 Then logit ("socket connected to " & tcpclient.RemoteHostIP & ":" & tcpclient.RemotePort)
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
   Dim val As String
 '  tcpclient.GetData val
 Dim strResponse As String

    tcpclient.GetData strResponse, vbString, bytesTotal
    
    strResponse = FormatLineEndings(strResponse)
    
    ' we append this to the response box becuase data arrives
    ' in multiple packets
    response.Text = response.Text & strResponse
If Winunciator.Debugcheck.Value = 1 Then logit ("Arrival: socket: " & val)
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
     If Winunciator.Debugcheck.Value = 1 Then logit ("socket error: " & Source & ":" & Description)
End Sub


Private Function AddClientHistory(val As String)
      If Winunciator.Debugcheck.Value = 1 Then logit (val) ' limit us to 2000 lines, otherwise it grows out of control
End Function


Private Sub Form_Load()
    ProgressBar1.Value = 0
    updater.Show
   Pause (0.5)
   If Winunciator.Debugcheck.Value = 1 Then logit ("Preparing for Possible Update.")
   statusreport.Caption = "Preparing for Possible Update."
   ProgressBar1.Value = 5
   updater.Refresh
   Pause (0.5)
   If tcpclient.State <> sckClosed Then
      tcpclient.Close
   End If
   tcpclient.RemoteHost = "colsolgrp.com"
    tcpclient.RemotePort = 80 ' what to connect to
    If Winunciator.Debugcheck.Value = 1 Then logit ("Connecting to Update Server.")
   statusreport.Caption = "Connecting to Update Server."
   ProgressBar1.Value = 10
   updater.Refresh
   tcpclient.Connect
   
    
   While Not tcpclient.State = sckConnected
   DoEvents
   Wend
  ' Pause (4.1)
  ' While tcpclient.State = sckConnected
   'If tcpclient.State = sckConnected Then
   tcpclient.SendData "GET /rss/winunciator.txt HTTP/1.1" & vbCrLf & "Host:colsolgrp.com:80" & vbCrLf
   If Winunciator.Debugcheck.Value = 1 Then logit ("get http://colsolgrp.com/rss/winunciator.txt")
   'End If
 '  Pause (2.1)
   'tcpclient.SendData "GET /rss/winunciator.txt HTTP/1.1" & vbCrLf & "Host: www.colsolgrp.com" & vbCrLf
 '   Wend
   '  Next step is to oGET THE PAGE
   
End Sub



' this function converts all line endings to Windows CrLf line endings
Private Function FormatLineEndings(ByVal str As String) As String
    Dim prevChar As String
    Dim nextChar As String
    Dim curChar As String
    
    Dim strRet As String
    
    Dim X As Long
    
    prevChar = ""
    nextChar = ""
    curChar = ""
    strRet = ""
    
    For X = 1 To Len(str)
        prevChar = curChar
        curChar = Mid$(str, X, 1)
                
        If nextChar <> vbNullString And curChar <> nextChar Then
            curChar = curChar & nextChar
            nextChar = ""
        ElseIf curChar = vbLf Then
            If prevChar <> vbCr Then
                curChar = vbCrLf
            End If
            
            nextChar = ""
        ElseIf curChar = vbCr Then
            nextChar = vbLf
        End If
        
        strRet = strRet & curChar
    Next X
    
    FormatLineEndings = strRet
End Function

Private Sub geturl()
Info = Inet2.OpenURL(StrUrl)
End Sub
Info = Inet2.OpenURL(StrUrl)


