Attribute VB_Name = "Module1"
Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long, time As Byte
Global speaknum
Global datavalid, Debugcheck, tcpconnect As Boolean
' Create an Icon in System Tray Needs
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


Sub logit(logfragment)
frmDebug.log1.Text = Date & " " & TimeValue(Now) & " " & logfragment & vbCrLf & frmDebug.log1.Text
frmDebug.Refresh
End Sub

Sub formatdata(callernumber)
   speaknum = ""
Rem  winunciator set for no number formatting

    If (frmConfig.format.Text = "None" And callernumber <> "") Then
         For speakprocess = 1 To Len(callernumber)
        speaknum = speaknum & Mid(callernumber, speakprocess, 1) & " "
        Next speakprocess
Rem  winunciator set for  NAPA number formatting
  ElseIf (frmConfig.format.Text = "(123) 456-7890" And callernumber <> "") Then
        areacode = Mid(callernumber, 1, 3)
        prefix = Mid(callernumber, 4, 3)
        last4 = Mid(callernumber, 7, 4)
        Winunciator.CNUM.Text = "(" & areacode & ") " & prefix & "-" & last4
            For speakprocess = 1 To Len(callernumber)
            speaknum = speaknum & Mid(callernumber, speakprocess, 1) & " "
                If speakprocess = 3 Then
                    speaknum = speaknum & " , "
                End If
                If speakprocess = 6 Then
                    speaknum = speaknum & " , "
                End If
            Next
Rem  winunciator set for EURO PA
   ElseIf (frmConfig.format.Text = "12 34 56 78 90" And callernumber <> "") Then
    s1 = Mid(callernumber, 1, 2)
    s2 = Mid(callernumber, 3, 2)
    s3 = Mid(callernumber, 5, 2)
    s4 = Mid(callernumber, 7, 2)
    s5 = Mid(callernumber, 9, 2)
    Winunciator.CNUM.Text = s1 & " " & s2 & " " & s3 & " " & s4 & " " & s5
        For speakprocess = 1 To Len(callernumber)
            speaknum = speaknum & Mid(callernumber, speakprocess, 1) & " "
            If (speakprocess = 2 Or speakprocess = 4 Or speakprocess = 6 Or speakprocess = 8) Then
                speaknum = speaknum & " , "
            End If
        Next speakprocess
    Else
    speaknum = callernumber
   End If
    
    If Winunciator.Debugcheck.Value = 1 Then logit ("Spoken Number: " & speaknum)
    
    Winunciator.CNUM.Refresh
  
End Sub


Public Sub Pause(NbSec As Single)
 Dim Finish As Single
 Finish = Timer + NbSec
 DoEvents
 Do Until Timer >= Finish
 Loop
End Sub

