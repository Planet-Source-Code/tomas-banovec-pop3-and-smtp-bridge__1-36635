VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POP3 and SMTP local bridge"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrDisConnect 
      Left            =   2640
      Top             =   240
   End
   Begin MSWinsockLib.Winsock sckSMTPWan 
      Index           =   0
      Left            =   2040
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckSMTPLocal 
      Index           =   0
      Left            =   1560
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPOPWan 
      Index           =   0
      Left            =   1080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock sckPOPLocal 
      Index           =   0
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "© 2002 by Tomas Banovec     banovec@bdr.sk"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'Document type: $ Workfile
'Author: Tomas Banovec
'e-mail: banovec@yahoo.com
'        banovec@bdr.sk
'Created: 19th June 2002
'Last release: 26th June 2002
'Project name: POP3/SMTP local bridge
'=============================================================================

'How it works:
'- this app works like fake POP3/SMTP server. It recieves request from
'  client in local network and by defined user name and password will redirect
'  request to Internet
'  When no connection to internet is avaliable it will make it
'
Const MAX_DISCONNECT_TIME = 60
Private WithEvents fInet As WinInet ' catch events for WinInet Class
Attribute fInet.VB_VarHelpID = -1
Dim psDuns() As String ' contains list of all dial-ups
Dim strPath As String ' path of application
Dim plResult As Long
Dim TimeToDisconnect As Integer 'time in seconds to disconect active net connection

Private Sub cmdClose_Click()
Unload Me ':)
End Sub

Private Sub Form_Load()
On Error Resume Next
TimeToDisconnect = -1
tmrDisConnect.Interval = 1000 'every second will be called timer
tmrDisConnect.Enabled = True
strPath = App.Path 'to dont have problems with \
If Not Right(strPath, 1) = "\" Then strPath = strPath & "\" 'add \ ...

Set fInet = New WinInet ' create new instance of WinInet Class in memory
                        '(this class if for connecting to internet when no connection is avaliable)
sckPOPLocal(0).Close 'to be sure that nothing is wrong
sckPOPLocal(0).LocalPort = 110 'local port set to 110 / it's POP3 port
sckPOPLocal(0).Listen 'start waiting for connection

sckSMTPLocal(0).Close 'to be sure that nothing is wrong
sckSMTPLocal(0).LocalPort = 25 'local port set to 25 / it's SMTP port
sckSMTPLocal(0).Listen 'start waiting for connection


If sckPOPLocal(0).State = 2 And sckSMTPLocal(0).State = 2 Then 'let user know if app is listening (listening=2)
  lblStatus.Caption = "Status: OK (listening on port 110/25)"
Else
  If sckPOPLocal(0).State = 2 Or sckSMTPLocal(0).State = 2 Then
    If sckPOPLocal(0).State = 2 Then lblStatus.Caption = "Status: ERROR - port 25 in use"
    If sckSMTPLocal(0).State = 2 Then lblStatus.Caption = "Status: ERROR - port 110 in use"
  Else
    lblStatus.Caption = "Status: ERROR - port 110/25 in use"
  End If
End If
fInet.ListDUNs psDuns   'get names of all avaliable Dial-up connections
End Sub


Private Sub Form_Unload(Cancel As Integer)
  If vbYes = MsgBox("When you close this application, you will not be able to recieve e-mails in your private network." & vbCrLf & "Do you really want to close POP3/SMTP local bridge?", vbQuestion + vbYesNo) Then
    End ':)
  Else
    Cancel = 1 'cancel unload request, continue in app running
  End If
End Sub

Private Sub sckPOPLocal_Close(Index As Integer)
On Error Resume Next 'prevent errors
Unload sckPOPLocal(Index) 'unload socket - it is not needed
Unload sckPOPWan(Index) 'unload socket - it is not needed
End Sub

Private Sub sckPOPLocal_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'when new connection is requested
On Error Resume Next 'prevent app crash :)
Load sckPOPLocal(requestID) 'make new Winsock control
sckPOPLocal(requestID).Close 'close this control
sckPOPLocal(requestID).Accept requestID 'accept connection for this new created Winsock
sckPOPLocal(requestID).SendData "+OK 250" & vbCrLf 'send message to POP3 Client that
                                        'everything is ready for recieving mails
End Sub

Private Sub sckPOPLocal_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String
Dim strTmp As String 'variable for temporary values
sckPOPLocal(Index).GetData strData 'recieve data from Winsock buffer


'Debug.Print strData  'if you want to know what was send...

Select Case Left(strData, 4) 'POP3 message are first 4 chars from recieved string
  Case "USER"
    strTmp = Right(strData, Len(strData) - 5) 'grab username
    strTmp = Left(strTmp, Len(strTmp) - 2) 'grab username
    
    If Not IsPOP(strTmp) = "" Then 'check if we have this user in our database
      sckPOPLocal(Index).Tag = strTmp 'write user name into the .tag property
                      ' of winsock control - we will yet need it
                      ' PS: if someone knows what is real function of .tag, mail me ;))
      InetConnect ' if there is no connection to internet, connect
      Load sckPOPWan(Index)
      sckPOPWan(Index).Close
      sckPOPWan(Index).RemotePort = 110 'connecting to POP3 SERVER
      sckPOPWan(Index).RemoteHost = IsPOP(strTmp)
      sckPOPWan(Index).Connect 'connect to remote host
      
    Else
      sckPOPLocal(Index).Close 'this user is not in our database, so we are not able to
                      ' transfer this request becouse we don't know on which internet
                      ' address can user connect
    End If
  Case Else
    sckPOPWan(Index).SendData strData 'send recieved data to POP3 SERVER
End Select

End Sub

'This function will check for user name in accounts.txt .
'Result of function is "" when it isnt in database or
'string with addresss of pop3 server where to connect...
'
Function IsPOP(ByVal strUserName As String) As String
On Error Resume Next 'prevent error messages and app crash
Dim strRcv As String
IsPOP = ""

If Not Dir(strPath & "accounts.txt") = "" Then ' is there file?
  Close #1 'yes there is file, but what if...? Close file before new open
  Open strPath & "accounts.txt" For Input As #1 'open accounts.txt
  Do Until EOF(1) 'until last line is read, lines will be readed
    Line Input #1, strRcv 'recieve line
    If LCase(Left(strRcv, InStr(1, strRcv, vbTab) - 1)) = LCase(strUserName) Then 'if user name is same as that one in accounts.txt...
      strRcv = Right(strRcv, Len(strRcv) - InStr(1, strRcv, vbTab)) 'get name of pop3 server - part 1
      IsPOP = Left(strRcv, InStr(1, strRcv, vbTab) - 1) 'get name of pop3 server - part2
    Exit Do 'we found what we needed, exit reading lines
    End If ':)))
  Loop ' end of loop
  Close #1 'close file
  
End If

End Function


Function InetConnect()
TimeToDisconnect = MAX_DISCONNECT_TIME
    On Error GoTo ERR_CONN
    plResult = 0
    If UBound(psDuns) >= 0 Then 'is any dial up connection?
      plResult = fInet.StartDUN(Me.hWnd, psDuns(0)) 'try to dial first connection
                                                    'from avaliable connections
    End If
ERR_CONN:
End Function

Function InetDisconnect()
On Error Resume Next
    Dim plResult As Long
  plResult = fInet.HangUp
End Function

Private Sub sckPOPLocal_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'if error ocurred, close and unload all sockets
On Error Resume Next
Unload sckPOPLocal(Index)
Unload sckPOPWan(Index)
End Sub

Private Sub sckPOPWan_Close(Index As Integer)
'when downloading completed, close and unload sockets
On Error Resume Next
Unload sckPOPLocal(Index)
Unload sckPOPWan(Index)
End Sub

Private Sub sckPOPWan_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
TimeToDisconnect = MAX_DISCONNECT_TIME
Dim strData As String
sckPOPWan(Index).GetData strData 'recieve data from buffer
If sckPOPWan(Index).Tag = "" Then 'firs data (with UserName) we have to send manualy
      'because we have then already recieved from POP3 Client
  sckPOPWan(Index).SendData "USER " & sckPOPLocal(Index).Tag & vbCrLf
  sckPOPWan(Index).Tag = "1"
Else
  'All responses except first are send to POP3 Client
  sckPOPLocal(Index).SendData strData
End If
End Sub

Private Sub sckPOPWan_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'if error ocurred, close and unload all sockets and disconnect internet....
On Error Resume Next
Unload sckPOPLocal(Index)
Unload sckPOPWan(Index)
End Sub

Private Sub sckSMTPLocal_Close(Index As Integer)
'when sending completed, unload all sockets and disconnect net
On Error Resume Next
Unload sckSMTPWan(Index)
Unload sckSMTPLocal(Index)
End Sub

Private Sub sckSMTPLocal_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'when new connection is requested
On Error Resume Next 'prevent app crash :)
Load sckSMTPLocal(requestID) 'make new Winsock control
sckSMTPLocal(requestID).Close 'close this control
sckSMTPLocal(requestID).Accept requestID 'accept connection for this new created Winsock
sckSMTPLocal(requestID).SendData "250 " & sckSMTPLocal(requestID).LocalHostName & vbCrLf  'send message to SMTP Client that
            'everything is ready for sending mails
End Sub

'When sending e-mail via SMTP we need to know address of SMTP server
'where to connect from this bridge. Because of this, app will recieve
'first two messages -> first is HELO greeting (handshaking ;)) )
'second is MAIL FROM: which contains e-mail address of sender
'For this address app search in accounts.txt If it is found,
'app will connect to SMTP internet address from file accounts.txt
'After this, the first (HELO) message is replaced by our own
'and the second address is recieved from sckSMTPLocal(Index).Tag
'Every next message is imediatelly send to SMTP client or server...
'When authentification required it is simmilar but little different
'(check in code)

Private Sub sckSMTPLocal_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next 'prevent app crash :)
Dim strData As String
Dim tmpStr As String
sckSMTPLocal(Index).GetData strData 'recieve data from Winsock buffer
'Debug.Print strData

If Left(strData, 4) = "EHLO" Then 'Authentification REQUIRED
  sckSMTPLocal(Index).SendData "250 AUTH=LOGIN" & vbCrLf 'send message to SMTP Client that all is OK
  sckSMTPLocal(Index).Tag = ""
  Exit Sub
End If

If Left(strData, 4) = "AUTH" Then sckSMTPLocal(Index).Tag = "1"

Select Case sckSMTPLocal(Index).Tag
  Case "1"
    sckSMTPLocal(Index).SendData "334 VXNlcm5hbWU6" & vbCrLf  'Username: encoded in base64
    sckSMTPLocal(Index).Tag = "2"
    Exit Sub
  Case "2"
    If IsSMTP_ENC(strData) <> "" Then
      InetConnect ' if there is no connection to internet, connect
      sckSMTPLocal(Index).Tag = strData
      Load sckSMTPWan(Index)
      sckSMTPWan(Index).Close
      sckSMTPWan(Index).RemotePort = 25 'connecting to SMTP SERVER
      sckSMTPWan(Index).RemoteHost = IsSMTP_ENC(strData)
      sckSMTPWan(Index).Connect 'connect to remote host
    Else
      sckSMTPLocal(Index).Close 'username not found / cannot connect to remote host (unknown name)
    End If
    Exit Sub
End Select

If Left(strData, 4) = "HELO" Then
  sckSMTPLocal(Index).SendData "250 OK" & vbCrLf 'send message to SMTP Client that all is OK
  Exit Sub
End If

If Left(strData, 10) = "MAIL FROM:" And sckSMTPLocal(Index).Tag = "" Then 'get e-mail address to know where to connect (only if not required password)
  tmpStr = strData
  tmpStr = Right(tmpStr, Len(tmpStr) - 11)
  tmpStr = Left(tmpStr, Len(tmpStr) - 2)
  If Left(tmpStr, 1) = "<" Then tmpStr = Right(tmpStr, Len(tmpStr) - 1)
  If Right(tmpStr, 1) = ">" Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)
  If Not IsSMTP(tmpStr) = "" Then
      InetConnect ' if there is no connection to internet, connect
      sckSMTPLocal(Index).Tag = strData
      Load sckSMTPWan(Index)
      sckSMTPWan(Index).Close
      sckSMTPWan(Index).RemotePort = 25 'connecting to SMTP SERVER
      sckSMTPWan(Index).RemoteHost = IsSMTP(tmpStr)
      sckSMTPWan(Index).Connect 'connect to remote host
  Else
    sckSMTPLocal(Index).Close
  End If
  Exit Sub
End If

sckSMTPWan(Index).SendData strData
End Sub

Function IsSMTP_ENC(ByVal strUserName As String) As String
On Error Resume Next 'prevent error messages and app crash
Dim strRcv As String
Dim tmStr As String
IsSMTP_ENC = ""
strUserName = Left(strUserName, Len(strUserName) - 2)
If Not Dir(strPath & "accounts.txt") = "" Then
  Close #1
  Open strPath & "accounts.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, strRcv 'recieve line
    tmStr = strRcv
    tmStr = LCase(Right(tmStr, Len(tmStr) - InStr(1, tmStr, vbTab)))
    tmStr = LCase(Right(tmStr, Len(tmStr) - InStr(1, tmStr, vbTab)))
    If Base64_Encode(LCase(Left(strRcv, InStr(1, strRcv, vbTab) - 1))) = strUserName Then
      IsSMTP_ENC = Left(tmStr, InStr(1, tmStr, vbTab) - 1)
      Exit Do
    End If
  Loop
  Close #1
End If

End Function


'check wheather is e-mail address of user in database
'if yes, result is name of SMTP server for this e-mail account
' AVALIABLE ONLY IF NOT REQUIRED AUTHENTIFICATION
Function IsSMTP(ByVal strUserName As String) As String
On Error Resume Next 'prevent error messages and app crash
Dim strRcv As String
Dim tmStr As String
IsSMTP = ""

If Not Dir(strPath & "accounts.txt") = "" Then
  Close #1
  Open strPath & "accounts.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, strRcv 'recieve line
    tmStr = strRcv
    tmStr = LCase(Right(tmStr, Len(tmStr) - InStr(1, tmStr, vbTab)))
    tmStr = LCase(Right(tmStr, Len(tmStr) - InStr(1, tmStr, vbTab)))
    If LCase(Right(tmStr, Len(tmStr) - InStr(1, tmStr, vbTab))) = LCase(strUserName) Then
      IsSMTP = Left(tmStr, InStr(1, tmStr, vbTab) - 1)
    Exit Do
    End If
  Loop
  Close #1
  
End If

End Function

Private Sub sckSMTPWan_Close(Index As Integer)
On Error Resume Next
Unload sckSMTPWan(Index)
Unload sckSMTPLocal(Index)
End Sub

Private Sub sckSMTPWan_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next 'prevent app crash :)
Dim strData As String

sckSMTPWan(Index).GetData strData 'recieve data from buffer
Select Case sckSMTPWan(Index).Tag
  Case ""
    If Left(sckSMTPLocal(Index).Tag, 10) = "MAIL FROM:" Then
    'HELO is sent when authentification NOT required
    '(for require auth. must be this set in SMTP client -e.g. Outlook)
      sckSMTPWan(Index).SendData "HELO " & sckSMTPWan(Index).LocalHostName & vbCrLf
    Else
      'EHLO is sent when authentification required by client
      sckSMTPWan(Index).SendData "EHLO " & sckSMTPWan(Index).LocalHostName & vbCrLf
    End If
    sckSMTPWan(Index).Tag = "1"
  Case "1"
    If Left(sckSMTPLocal(Index).Tag, 10) = "MAIL FROM:" Then ' Require autentification
      sckSMTPWan(Index).SendData sckSMTPLocal(Index).Tag     ' not enabled in SMTP Client
    Else
      sckSMTPWan(Index).SendData "AUTH LOGIN" & vbCrLf 'most used base64_Encode request
    End If
    sckSMTPWan(Index).Tag = "2"
  Case "2"
    If Left(sckSMTPLocal(Index).Tag, 10) = "MAIL FROM:" Then 'Require autentification
      'All responses except first two are send to SMTP Client / if not Authentification
      sckSMTPLocal(Index).SendData strData
    Else
      sckSMTPWan(Index).SendData sckSMTPLocal(Index).Tag 'send username encoded in base64
    End If
    sckSMTPWan(Index).Tag = "3"
  Case "3"
    sckSMTPLocal(Index).SendData strData 'bridge messages between client and server...
End Select
End Sub

Private Sub sckSMTPWan_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'on socket error close all connections
On Error Resume Next
Unload sckSMTPWan(Index)
Unload sckSMTPLocal(Index)
End Sub

Private Sub tmrDisConnect_Timer()
On Error Resume Next
If TimeToDisconnect >= 0 Then TimeToDisconnect = TimeToDisconnect - 1
If TimeToDisconnect = 0 Then InetDisconnect
End Sub
