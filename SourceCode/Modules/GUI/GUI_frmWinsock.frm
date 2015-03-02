VERSION 5.00
Begin VB.Form GUI_frmWinsock 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "AnGeL Bot"
   ClientHeight    =   5145
   ClientLeft      =   4500
   ClientTop       =   3465
   ClientWidth     =   6465
   Icon            =   "GUI_frmWinsock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6465
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.PictureBox HideThings 
      BorderStyle     =   0  'Kein
      Height          =   5145
      Left            =   0
      ScaleHeight     =   5145
      ScaleWidth      =   6555
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   6555
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   1140
         TabIndex        =   16
         Top             =   1800
         Width           =   4275
         Begin VB.TextBox PWField 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   630
            Width           =   3795
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3555
            Picture         =   "GUI_frmWinsock.frx":5C12
            Top             =   195
            Width           =   480
         End
         Begin VB.Label Label6 
            Caption         =   "Please enter bot password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   17
            Top             =   300
            Width           =   3825
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   255
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   150
   End
   Begin VB.Timer PersonalTimer 
      Interval        =   60000
      Left            =   1125
      Top             =   1125
   End
   Begin VB.Timer ResolveTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   4560
   End
   Begin VB.Timer PingServer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1800
      Top             =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bot Shutdown"
      Height          =   315
      Left            =   4650
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1755
   End
   Begin VB.ListBox lstChannels 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   4920
      TabIndex        =   11
      Top             =   2760
      Width           =   1485
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   120
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   150
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   2760
      Width           =   4665
   End
   Begin VB.Timer ChanCheck 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2040
      Top             =   4560
   End
   Begin VB.Timer TimedEvents 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   4560
   End
   Begin VB.Timer ConnectTimeOut 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2520
      Top             =   4560
   End
   Begin VB.Timer BufferExTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2760
      Top             =   4560
   End
   Begin VB.Timer FloodTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3000
      Top             =   4560
   End
   Begin VB.Timer NTService 
      Interval        =   50
      Left            =   3240
      Top             =   4560
   End
   Begin VB.Timer TTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3480
      Top             =   4560
   End
   Begin VB.TextBox txtWinsock 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2175
      Width           =   6255
   End
   Begin VB.PictureBox imgStateHub 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   450
      Picture         =   "GUI_frmWinsock.frx":6054
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   300
   End
   Begin VB.PictureBox imgStateGreen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   450
      Picture         =   "GUI_frmWinsock.frx":6B8E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   360
      Width           =   300
   End
   Begin VB.PictureBox imgStateYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   450
      Picture         =   "GUI_frmWinsock.frx":76C8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   720
      Width           =   300
   End
   Begin VB.PictureBox imgStateRed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   435
      Picture         =   "GUI_frmWinsock.frx":8202
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   300
   End
   Begin VB.PictureBox imgStateRed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "GUI_frmWinsock.frx":8D3C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   300
   End
   Begin VB.PictureBox imgStateYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "GUI_frmWinsock.frx":9876
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   720
      Width           =   300
   End
   Begin VB.PictureBox imgStateGreen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "GUI_frmWinsock.frx":A3B0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   360
      Width           =   300
   End
   Begin VB.PictureBox imgStateHub 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "GUI_frmWinsock.frx":AEEA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label Label4 
      Caption         =   "Channels:"
      Height          =   225
      Left            =   4920
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6660
      Y1              =   4485
      Y2              =   4485
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6660
      Y1              =   4470
      Y2              =   4470
   End
   Begin VB.Label Label3 
      Caption         =   "System Messages:"
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "© 1998-2003 AnGeL-Team <http://www.angel-bot.de>"
      ForeColor       =   &H80000010&
      Height          =   225
      Left            =   2355
      TabIndex        =   8
      Top             =   105
      UseMnemonic     =   0   'False
      Width           =   4035
   End
   Begin VB.Label Label1 
      Caption         =   "Server Messages:"
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   105
      Width           =   1575
   End
   Begin VB.Menu mnuPump 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect / Disconnect"
      End
      Begin VB.Menu mnuShowMain 
         Caption         =   "Show Main Window"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "Exit Program"
      End
   End
End
Attribute VB_Name = "GUI_frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BlockResolve As Boolean

'Implements ICallBack

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'  TIMERS
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////

'Sends the next lines of the winsock2_send buffer
Private Sub BufferExTimer_Timer()
  BytesSent = 0
  BufferEx
End Sub

'Checks if the bot is in all channels it should be and tries to join if it's not
Private Sub ChanCheck_Timer()
  If (Connected = True) And (Initializing = False) Then CheckChannels
End Sub

'Checks whether the connect to an IRC server timed out
Private Sub ConnectTimeOut_Timer()
  Output "*** connect timed out. Retrying..." & vbCrLf
  'Reset all variables
  Disconnect
  'Reconnect in (i.e. 5) seconds
  ConnectServer CurrentConnectDelay, ""
End Sub

Private Sub FloodTimer_Timer()
  Dim u As Long, u2 As Long, TheFile As String
  'New day -> delete old logs, show partyline message
  If (Hour(Now) = 0) And (LastDay <> Day(Now)) Then
    SpreadMessage 0, -1, "14[" & Time & "] --- " & Format(Now, "dddd, dd.mm.yyyy")
    SpreadMessage 0, -1, MakeMsg(MSG_PLSwitchLogs1)
    LastDay = Day(Now)
    TheFile = Dir(FileAreaHome & "Logs\*.log")
    Do
      If TheFile = "" Then Exit Do
      If CDate(LogToDate(TheFile) + LogMaxAge) < Now Then Kill FileAreaHome & "Logs\" & TheFile
      TheFile = Dir
    Loop
    SpreadMessage 0, -1, MakeMsg(MSG_PLSwitchLogs2)
    
    'Check Compression
    NTFS_CheckCompress FileAreaHome & "Logs"
    NTFS_CheckCompress FileAreaHome & "Scripts"
  End If
  
  'Write userlist if necessary
  WriteUserList
  
  'Reset channel and party line flood check values
  For u = 1 To ChanCount
    For u2 = 1 To Channels(u).UserCount
      If Channels(u).User(u2).CTCPs > 0 Then Channels(u).User(u2).CTCPs = Channels(u).User(u2).CTCPs - 1
      Channels(u).User(u2).NickChanges = 0
    Next u2
  Next u
  For u = 1 To SocketCount
    If IsValidSocket(u) Then SocketItem(u).NumOfServerEvents = 0
  Next u
  If AuthJust = True Then AuthJust = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If HideThings.Visible = True Then Cancel = True: Beep
  Else
    If FormPassword <> "" Then WritePPString "Others", "Password", FormPassword, AnGeL_INI
  End If
  WritePPString "Others", "GlobalBytesSent", CStr(GlobalBytesSent), AnGeL_INI
  WritePPString "Others", "GlobalBytesReceived", CStr(GlobalBytesReceived), AnGeL_INI
End Sub

Private Sub NTService_Timer()
  If IsNTService Then If Not ContinueService Then TimedEvent "winsock2_shutdown", 0
End Sub

Private Sub PersonalTimer_Timer()
  Static PTimer As Long
  Dim i As Long
  PTimer = PTimer + 1
  If PTimer >= 2 Then
    PTimer = 1
    For i = 0 To UBound(Channels)
      SpreadChanMessage Channels(i).Name, "14[" & Mid(Time, 1, 5) & "] " & ChanStatus(Channels(i).Name)
    Next i
  End If
End Sub

Private Sub PingServer_Timer()
Dim t As Long, u As Long, u2 As Long, vsock As Long, ErrLine As Long, ThereAreOps As Boolean
Dim Rest As String, Others As Boolean, RequestGO As String
On Error GoTo PingError
ErrLine = 1
  
  'Regain ops (GO)
  For u = 1 To ChanCount
ErrLine = 102
    Channels(u).FloodEvents = 0
ErrLine = 103
    If Not Channels(u).GotOPs And Channels(u).CompletedWHO And Not Left(Channels(u).Name, 1) = "&" Then
ErrLine = 104
      ThereAreOps = False: Others = False: RequestGO = ""
      For u2 = 1 To Channels(u).UserCount
ErrLine = 105
        Select Case Channels(u).User(u2).Status
          Case "@", "@+": If InStr("-+=", Left(Mask(Channels(u).User(u2).Hostmask, 12), 1)) = 0 Then ThereAreOps = True: Exit For
ErrLine = 107
        End Select
ErrLine = 108
        If Channels(u).User(u2).Nick <> MyNick Then
          Rest = GetUserChanFlags(Channels(u).User(u2).RegNick, Channels(u).Name)
ErrLine = 109
          If InStr(Rest, "o") > 0 And InStr(Rest, "b") > 0 Then
ErrLine = 110
            If RequestGO <> "" Then RequestGO = RequestGO & " " & Channels(u).User(u2).Nick Else RequestGO = Channels(u).User(u2).Nick
ErrLine = 111
          Else
ErrLine = 112
            Others = True
ErrLine = 113
          End If
        End If
ErrLine = 114
      Next u2
ErrLine = 115
      If Not (Others Or ThereAreOps) Then
ErrLine = 116
        Rest = "": t = 0
ErrLine = 117
        For u2 = 1 To ParamCount(RequestGO)
ErrLine = 118
          t = t + 1
ErrLine = 119
          If Rest <> "" Then Rest = Rest & "," & Param(RequestGO, u2) Else Rest = Param(RequestGO, u2)
ErrLine = 120
          If t = 5 Then
ErrLine = 121
            SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLRequestedGo, Channels(u).Name, Rest)
ErrLine = 122
            SendLine "privmsg " & Rest & " :GO " & Channels(u).Name, 2: Rest = "": t = 0
          End If
        Next u2
        If Rest <> "" Then
ErrLine = 123
          SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLRequestedGo, Channels(u).Name, Rest)
ErrLine = 124
          SendLine "privmsg " & Rest & " :GO " & Channels(u).Name, 2: Rest = "": t = 0
        End If
ErrLine = 125
      End If
    End If
  Next u
ErrLine = 200
  
  'Unlink bots which are not really linked
  UnlinkDummyBots

  'Reconnect bot if there's no answer from IRC server
  If HubBot = False Then
    If GotServerPong = False Then
  ErrLine = 3
      If LastEvent + CDate("00:06:00") < Now Then
        Output "*** Server didn't answer. Reconnecting." & vbCrLf
        PutLog "Server didn't answer. Reconnecting. <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
  ErrLine = 4
        SendLine "quit :ERROR - No Server Ping Reply", 1
  ErrLine = 5
        Disconnect
        DontConnect = False
  ErrLine = 12
        ConnectServer CurrentConnectDelay, ""
        Exit Sub
      End If
  ErrLine = 14
    End If
  ErrLine = 15
  End If
  
  'Anti-Idle
  If Connected Then
    If LastSendTime + CDate("00:01:25") < Now Then
      SendLine "privmsg #anti-i-chan-" & Trim(Str(Int(Rnd * 999999))) & " :" & Choose(Int(Rnd * 4) + 1, "Much more than a simple message.", "Caress me! I like it.", "I like to chat with myself!", "I don't have bugs, I am one. =)"), 3
    End If
  End If

  'Write new seen entries to disk
  FlushExtSeenEntries

ErrLine = 17
  CountCTCPs = 0
  
  'Bringing bot back to normal state after floods
  If UnIgnoreTimes > 0 Then
    UnIgnoreTimes = UnIgnoreTimes - 1
    If UnIgnoreTimes = 0 Then SpreadMessage 0, -1, "4*** I'm not being flooded anymore. Removed ignores."
  End If
  
  For u = 1 To ChanCount
ErrLine = 24
    'GetOps: Request ops
    RequestOps u
ErrLine = 25
    'XGetOps: Offer ops
    OfferOps u
  Next u
ErrLine = 27
  
  'connect to hub bots
ErrLine = 29
  ConnectHubs
  
  CountServerPing = CountServerPing + 1
  If CountServerPing >= 3 Then
    LastWhatisOutput = ""
    LastWhoisOutput = ""
    LastSeenOutput = ""
ErrLine = 28
    PingBots
ErrLine = 30
    'If BlockSends Then SendLine "PING 1", 1
ErrLine = 31
    If HubBot = False Then
      If LastEvent + CDate("00:02:00") < Now Then Output "*** Sending Server PING" & vbCrLf: GotServerPong = False: SendLine "ping " & StripDP(ServerName), 1
    End If
    CountServerPing = 0
ErrLine = 35
  End If
Exit Sub
PingError:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  PutLog "||| ] PingServer ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
  PutLog "||| ] Der Fehler " & ErrNumber & " (" & ErrDescription & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & ErrLine
  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
'  SendNote "AnGeL PingServer ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & ErrLine)
'  SendNote "AnGeL PingServer ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & ErrLine)
  'Stop
End Sub

Private Sub PWField_KeyPress(KeyAscii As Integer)
  If Len(PWField.Text) > 50 Then Exit Sub
  If KeyAscii = 13 Then
    KeyAscii = 0
    If EncryptIt(PWField.Text) = FormPassword Then
      HideThings.Visible = False: mnuConnect.Enabled = True: mnuEnd.Enabled = True
    Else
      Beep
    End If
    PWField.Text = ""
  End If
End Sub

'Resolves hosts to IPs
Private Sub ResolveTimer_Timer()
'Dim u As Long, u2 As Long, RemovedOne As Boolean, RunningResolves As Long
'On Error GoTo RTimErr
'  If BlockResolve Then Exit Sub
'  BlockResolve = True
'  'Give up after 30 seconds
'  Do
'    RemovedOne = False
'    For u = 1 To ResolveCount
'      If Resolves(u).ASync <> 0 Then
'        If (Resolves(u).StartedAt + CDate("00:00:45")) < Now Then
'          PutLog "*R*: Cancelling Request: " & Resolves(u).ASync & " - " & Resolves(u).Host
'          winsock2_WSACancelAsyncRequest Resolves(u).ASync
'          For u2 = u To ResolveCount - 1
'            Resolves(u2) = Resolves(u2 + 1)
'          Next u2
'          ResolveCount = ResolveCount - 1
'          RemovedOne = True
'          Exit For
'        End If
'      End If
'    Next u
'  Loop While RemovedOne
'
'  'Don't allow more than 5 resolves at once
'  For u = 1 To ResolveCount
'    If Resolves(u).ASync <> 0 Then RunningResolves = RunningResolves + 1
'  Next u
'  If RunningResolves > 5 Then BlockResolve = False: Exit Sub
'
'  For u = 1 To ResolveCount
'    If Resolves(u).ASync = 0 Then
'      'Status Resolves(u).Host + vbCrLf
'      'PutLog "*R*: Resolving: " & Resolves(u).Host
'      Resolves(u).ASync = WSAAsyncGetHostByNameAlias(GUI_frmWinsock.hWnd, 1026, Resolves(u).Host)
'      Resolves(u).StartedAt = Now
'      Exit For
'    End If
'  Next u
'  BlockResolve = False
'  Exit Sub
'RTimErr:
'  Dim ErrNumber As Long, ErrDescription As String
'  ErrNumber = Err.Number
'  ErrDescription = Err.Description
'  Err.Clear
'  PutLog "||| ] ResolveTimer ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
'  PutLog "||| ] Der Fehler " & ErrNumber & " (" & ErrDescription & ") ist beim Bearbeiten folgender Zeile aufgetreten"
'  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
''  'Stop
End Sub

Private Sub TimedEvents_Timer()
Dim u As Long, ChNum As Long, UsNum As Long, RemovedOne As Boolean, ErrLine As Long, M As Long
Dim ChangMode As String, u2 As Long, vsock As Long, TempStr As String
Dim TempStr2 As String, TimeNow As Currency
On Error GoTo TevError
  ErrLine = 1
  
  TimeNow = WinTickCount
  'WinTickCount wrapped around, set all idle timers to zero (to prevent mass idlekicks etc.)
  If TimeNow < LastTick Then
    For ChNum = 1 To ChanCount
      For UsNum = 1 To Channels(ChNum).UserCount
        Channels(ChNum).User(UsNum).LastEvent = TimeNow
      Next UsNum
    Next ChNum
    Status "*** Resetted idle timers due to kernel32_GetTickCount wrap-around." & vbCrLf
    PutLog "*** Resetted idle timers due to kernel32_GetTickCount wrap-around."
  End If
  ErrLine = 2

  'Unbans
  UnbanCount = UnbanCount + 1
  ErrLine = 201
  If UnbanCount > 50 Then UnbanCount = 0: UnbanTime = FindUnbanTime
  ErrLine = 202
  If UnbanCount = 0 Then DontSeen = False: LastSeenOutput = "": LastWhoisOutput = ""
  ErrLine = 3
  
  'Ignores + Actions
  For u = 1 To EventCount
  ErrLine = 5
    If u > UBound(Events()) Then Exit For
    If (Events(u).AtTime <= TimeNow) And (Events(u).DoThis <> "") Then
  ErrLine = 6
      Select Case Param(Events(u).DoThis, 1)
        Case "UnIgnore"
  ErrLine = 8
            TempStr = Param(Events(u).DoThis, 2)
            RemIgnore TempStr
  ErrLine = 9
        Case "UnBan"
  ErrLine = 10
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 11
            If IsBanned(ChNum, Param(Events(u).DoThis, 3)) Then WaitThisLine = True: SendLine "mode " & Channels(ChNum).Name & " -b " & Param(Events(u).DoThis, 3), 2
  ErrLine = 12
        Case "Op"
  ErrLine = 13
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 14
            If ChNum > 0 Then
  ErrLine = 15
              UsNum = FindUser(Param(Events(u).DoThis, 3), ChNum)
  ErrLine = 16
              If UsNum > 0 Then If InStr(Channels(ChNum).User(UsNum).Status, "@") = 0 Then GiveOp Channels(ChNum).Name, Channels(ChNum).User(UsNum).Nick
  ErrLine = 17
            End If
  ErrLine = 18
        Case "gop"
  ErrLine = 19
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 20
            RequestOps ChNum
  ErrLine = 21
        Case "xgop"
  ErrLine = 22
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 23
            OfferOps ChNum
  ErrLine = 24
        Case "checkforflood"
            ChNum = FindChan(Param(Events(u).DoThis, 2))
            If ChNum > 0 Then
              If (Channels(ChNum).FloodEvents = 0) And (Channels(ChNum).KickCount = 0) And (Channels(ChNum).ToBanCount = 0) Then
                SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLFloodEnd, Channels(ChNum).Name)
                If InStr(Param(ChangeMode(GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), ""), 1), "i") = 0 Then
                  ChangMode = ChangeMode("-im" & GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), Channels(ChNum).Mode)
                  If ChangMode <> "" Then SendLine "mode " & Channels(ChNum).Name & " " & ChangMode, 1
                Else
                  If InStr(Param(ChangeMode(GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), ""), 1), "m") = 0 Then
                    ChangMode = ChangeMode("-m" & GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), Channels(ChNum).Mode)
                    If ChangMode <> "" Then SendLine "mode " & Channels(ChNum).Name & " " & ChangMode, 1
                  End If
                End If
                Channels(ChNum).InFlood = False
              Else
                'Channel is still being flooded / kicks are not done? Check back in 30 seconds.
                TimedEvent Events(u).DoThis, 30
              End If
            End If
        'Checks whether a specific ban matches users on a channel
        Case "CheckBans"
            ChNum = FindChan(Param(Events(u).DoThis, 2))
            If ChNum > 0 Then
              TempStr = Param(Events(u).DoThis, 3)
              If IsBanned(ChNum, TempStr) Then
                TempStr2 = Param(Events(u).DoThis, 4)
                For u2 = 1 To Channels(ChNum).UserCount
                  If (Channels(ChNum).User(u2).RegNick = "") And (Channels(ChNum).User(u2).Nick <> MyNick) And (InStr(Channels(ChNum).User(u2).Status, "@") = 0) Then
                    If MatchWM(TempStr, Channels(ChNum).User(u2).Hostmask) Or MatchWM(TempStr, Channels(ChNum).User(u2).IPmask) Then
                      AddKickUser ChNum, Channels(ChNum).User(u2).Nick, "", "Banned" & IIf(TempStr2 <> MyNick, " by " & TempStr2, "")
                    End If
                  End If
                Next u2
              End If
            End If
        'Call a script
        Case "CallScript"
            TempStr = Param(Events(u).DoThis, 2)
            For u2 = 1 To ScriptCount
              If Scripts(u2).Name = TempStr Then
                RunScript u2, GetRest(Events(u).DoThis, 3)
                Exit For
              End If
            Next u2
        'winsock2_send lines in a user's SendQ
        Case "PutSendQ"
            vsock = CLng(Param(Events(u).DoThis, 2))
            If IsValidSocket(vsock) Then
              Do
                RemovedOne = False
                If SocketItem(vsock).SendQLines > 0 Then
                  If SendTCP(vsock, SocketItem(vsock).SendQ(1)) = -1 Then
                    'Give up after 10 winsock2_send tries (20 seconds) -> erase SendQ
                    SocketItem(vsock).SendQTries = SocketItem(vsock).SendQTries + 1
                    If SocketItem(vsock).SendQTries >= 10 Then
                      SocketItem(vsock).SendQTries = 0
                      SocketItem(vsock).SendQLines = 0
                      Output "Socket [" & CStr(vsock) & "] Gave up, erased SendQ." & vbCrLf
                    Else
                      Output "Socket [" & CStr(vsock) & "] Waiting..." & vbCrLf
                      TimedEvent "PutSendQ " & CStr(vsock), 2
                    End If
                  Else
                    For u2 = 1 To SocketItem(vsock).SendQLines - 1
                      SocketItem(vsock).SendQ(u2) = SocketItem(vsock).SendQ(u2 + 1)
                    Next u2
                    SocketItem(vsock).SendQLines = SocketItem(vsock).SendQLines - 1
                    SocketItem(vsock).SendQTries = 0
                    Output "Socket [" & CStr(vsock) & "] Sent SendQ line (rest: " & CStr(SocketItem(vsock).SendQLines) & ")" & vbCrLf
                    RemovedOne = True
                  End If
                End If
              Loop Until Not RemovedOne
            End If
        Case "KickOldBot"
            KickOldBot CLng(Param(Events(u).DoThis, 2)), Param(Events(u).DoThis, 3)
        Case "FinalBotNetLogin"
            FinalBotNetLogin CLng(Param(Events(u).DoThis, 2)), Param(Events(u).DoThis, 3)
        Case "RemRepeat"
  ErrLine = 25
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 26
            If ChNum > 0 Then
  ErrLine = 27
              UsNum = FindUser(Param(Events(u).DoThis, 3), ChNum)
  ErrLine = 28
              If UsNum > 0 Then
  ErrLine = 29
                Channels(ChNum).User(UsNum).RepeatCount = Channels(ChNum).User(UsNum).RepeatCount - 1
  ErrLine = 30
                If Channels(ChNum).User(UsNum).RepeatCount < 0 Then Channels(ChNum).User(UsNum).RepeatCount = 0
  ErrLine = 31
              End If
            End If
        Case "RemLine"
  ErrLine = 25
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 26
            If ChNum > 0 Then
  ErrLine = 27
              UsNum = FindUser(Param(Events(u).DoThis, 3), ChNum)
  ErrLine = 28
              If UsNum > 0 Then
  ErrLine = 29
                Channels(ChNum).User(UsNum).LineCount = Channels(ChNum).User(UsNum).LineCount - 1
  ErrLine = 30
                If Channels(ChNum).User(UsNum).LineCount < 0 Then Channels(ChNum).User(UsNum).LineCount = 0
  ErrLine = 31
              End If
            End If
        Case "RemChars"
  ErrLine = 32
            ChNum = FindChan(Param(Events(u).DoThis, 2))
  ErrLine = 33
            If ChNum > 0 Then
  ErrLine = 34
              UsNum = FindUser(Param(Events(u).DoThis, 3), ChNum)
  ErrLine = 35
              If UsNum > 0 Then
  ErrLine = 36
                Channels(ChNum).User(UsNum).CharCount = Channels(ChNum).User(UsNum).CharCount - CLng(Param(Events(u).DoThis, 4))
  ErrLine = 37
                If Channels(ChNum).User(UsNum).CharCount < 0 Then Channels(ChNum).User(UsNum).CharCount = 0
  ErrLine = 38
              End If
            End If
        Case "RemKI"
            RemKI Param(Events(u).DoThis, 2)
        Case "KIAnswer"
            If UnIgnoreTimes = 0 Then
              TempStr = Trim(ParamX(GetRest(Events(u).DoThis, 3), "|", 1))
              TempStr2 = Trim(GetRestX(GetRest(Events(u).DoThis, 3), "|", 2))
              If TempStr2 <> "" Then TimedEvent "KIAnswer " & Param(Events(u).DoThis, 2) & " " & TempStr2, 2 + Int(Len(ParamX(TempStr2, "|", 1)) / 6) + Int(Rnd * Len(ParamX(TempStr2, "|", 1)) / 10)
                For M = 1 To ScriptCount
                  If Scripts(M).Hooks.KI Then
                    ScriptCMDs.KIAnswer = TempStr
                    RunScriptX M, "KI", Param(Events(u).DoThis, 2), TempStr
                    TempStr = ScriptCMDs.KIAnswer
                  End If
                Next M
              SendLine "PRIVMSG " & Param(Events(u).DoThis, 2) & " :" & TempStr, 3
              SpreadFlagMessage 0, "+m", "14[" & Time & "] I replied to " & Param(Events(u).DoThis, 2) & ": " & TempStr
            End If
        Case "ConnectServer"
            TempStr = GetRest(Events(u).DoThis, 2)
            ConnectServer 0, TempStr
        Case "GetIdentSocket"
            GetIdentSocket CLng(Param(Events(u).DoThis, 2)), Param(Events(u).DoThis, 3), CLng(Param(Events(u).DoThis, 4))
        Case "AddTrayIcon"
            AddTrayIcon
        Case "UPTIME"
            winsock2_send_uptime
            TimedEvent "UPTIME", 60 * 60 * 6
        Case "winsock2_shutdown"
            UnloadBot
            Exit Sub
        Case "RESTART"
            BotRestart = True
            UnloadBot
            Exit Sub
        Case Else
  ErrLine = 22
            SendLine Events(u).DoThis, 2
  ErrLine = 23
      End Select
  ErrLine = 24
      Events(u).DoThis = ""
  ErrLine = 25
    End If
  ErrLine = 26
  Next u
  ErrLine = 27
  BufferEx
  
  Do
  ErrLine = 29
    RemovedOne = False
  ErrLine = 30
    For u = 1 To EventCount
  ErrLine = 31
      If Events(u).DoThis = "" Then RemoveTimedEventFromList u: RemovedOne = True: Exit For
  ErrLine = 32
    Next u
  ErrLine = 33
    If Not RemovedOne Then Exit Do
  ErrLine = 34
  Loop
  
  'Remove expired orders
  Do
  ErrLine = 36
    RemovedOne = False
  ErrLine = 37
    For u = 1 To OrderCount
  ErrLine = 38
      If u > UBound(Orders()) Then Exit For
  ErrLine = 39
      If Orders(u).AtTime <= TimeNow Then RemoveOrder u: RemovedOne = True: Exit For
  ErrLine = 40
    Next u
  ErrLine = 41
    If Not RemovedOne Then Exit Do
  Loop
  
  'Remove expired desired bans
  Do
    RemovedOne = False
    For ChNum = 1 To ChanCount
      If ChNum > UBound(Channels()) Then Exit For
      For u = 1 To Channels(ChNum).DesiredBanCount
        If u > UBound(Channels(ChNum).DesiredBanList()) Then Exit For
        If Channels(ChNum).DesiredBanList(u).Expires <= TimeNow Then TempStr = Channels(ChNum).DesiredBanList(u).Mask: RemDesiredBan ChNum, TempStr: RemovedOne = True: Exit For
      Next u
      If RemovedOne Then Exit For
    Next ChNum
    If Not RemovedOne Then Exit Do
  Loop
  
  Do
    RemovedOne = False
    For ChNum = 1 To ChanCount
      If ChNum > UBound(Channels()) Then Exit For
      For u = 1 To Channels(ChNum).DesiredExceptCount
        If u > UBound(Channels(ChNum).DesiredExceptList()) Then Exit For
        If Channels(ChNum).DesiredExceptList(u).Expires <= TimeNow Then TempStr = Channels(ChNum).DesiredExceptList(u).Mask: RemDesiredExcept ChNum, TempStr: RemovedOne = True: Exit For
      Next u
      If RemovedOne Then Exit For
    Next ChNum
    If Not RemovedOne Then Exit Do
  Loop
  
  Do
    RemovedOne = False
    For ChNum = 1 To ChanCount
      If ChNum > UBound(Channels()) Then Exit For
      For u = 1 To Channels(ChNum).DesiredInviteCount
        If u > UBound(Channels(ChNum).DesiredInviteList()) Then Exit For
        If Channels(ChNum).DesiredInviteList(u).Expires <= TimeNow Then TempStr = Channels(ChNum).DesiredInviteList(u).Mask: RemDesiredInvite ChNum, TempStr: RemovedOne = True: Exit For
      Next u
      If RemovedOne Then Exit For
    Next ChNum
    If Not RemovedOne Then Exit Do
  Loop
  
  ErrLine = 42
  For u = 1 To SocketCount
  ErrLine = 43
    If u > UBound(SocketItem()) Then Exit For
  ErrLine = 44
    If IsValidSocket(u) Then
  ErrLine = 45
      Select Case GetSockFlag(u, SF_Status)
        Case SF_Status_File
  ErrLine = 47
            If SocketItem(u).LastEvent + CDate("00:02:00") < Now Then
  ErrLine = 48
              SpreadLevelFileAreaMessage u, MakeMsg(MSG_PLDCCGetTO, SocketItem(u).RegNick, GetFileName(SocketItem(u).FileName))
  ErrLine = 49
              Close #SocketItem(u).FileNum
              DisconnectSocket u
  ErrLine = 53
            End If
  ErrLine = 54
        Case SF_Status_SendFile
  ErrLine = 247
            If SocketItem(u).LastEvent + CDate("00:02:00") < Now Then
  ErrLine = 248
              SpreadLevelFileAreaMessage u, MakeMsg(MSG_PLDCCSendTO, SocketItem(u).RegNick, GetFileName(SocketItem(u).FileName))
  ErrLine = 249
              Close #SocketItem(u).FileNum
  ErrLine = 350
              SendNextQueuedFile SocketItem(u).RegNick
              DisconnectSocket u
  ErrLine = 253
            End If
  ErrLine = 254
        Case SF_Status_BotGetName
            If SocketItem(u).LastEvent + CDate("00:02:00") < Now Then
              SpreadFlagMessage u, "+t", MakeMsg(MSG_PLBotNetConnTO, SocketItem(u).Hostmask)
              DisconnectSocket u
            End If
        Case SF_Status_BotLinking, SF_Status_BotPreCache, SF_Status_BotGetPass
            If SocketItem(u).LastEvent + CDate("00:02:00") < Now Then
              SpreadFlagMessage u, "+t", MakeMsg(MSG_PLBotNetConnTO, SocketItem(u).RegNick)
              DisconnectSocket u
            End If
        Case SF_Status_InitBotLink
  ErrLine = 62
            If SocketItem(u).LastEvent + CDate("00:01:00") < Now Then
  ErrLine = 63
              SpreadFlagMessage u, "+t", MakeMsg(MSG_PLBotNetLinkTO, SocketItem(u).RegNick)
              DisconnectSocket u
  ErrLine = 67
            End If
  ErrLine = 68
        Case SF_Status_SendFileWaiting
  ErrLine = 369
            If SocketItem(u).LastEvent + CDate("00:01:00") < Now Then
  ErrLine = 370
              SpreadLevelFileAreaMessage u, MakeMsg(MSG_PLDCCSendTO, SocketItem(u).RegNick, GetFileName(SocketItem(u).FileName))
  ErrLine = 371
              SendNextQueuedFile SocketItem(u).RegNick
              DisconnectSocket u
  ErrLine = 375
            End If
        Case SF_Status_DCCInit, SF_Status_DCCWaiting, SF_Status_FileWaiting
  ErrLine = 69
            If SocketItem(u).LastEvent + CDate("00:01:00") < Now Then
  ErrLine = 70
              SpreadFlagMessage u, "+m", MakeMsg(MSG_PLDCCConnTO, SocketItem(u).RegNick)
              DisconnectSocket u
  ErrLine = 74
            End If
  ErrLine = 75
        Case SF_Status_UserGetName, SF_Status_BotGetName
            If SocketItem(u).LastEvent + CDate("00:02:00") < Now Then
  ErrLine = 76
              SpreadFlagMessage u, "+m", MakeMsg(IIf(GetSockFlag(u, SF_Status) = SF_Status_BotGetName, MSG_PLBotNetNickTO, MSG_PLTelNetNickTO), SocketItem(u).Hostmask)
              DisconnectSocket u
            End If
        Case SF_Status_UserGetPass, SF_Status_BotGetPass, SF_Status_UserChoosePass, SF_Status_UserChooseColors
  ErrLine = 83
            If SocketItem(u).LastEvent + CDate("00:02:30") < Now Then
  ErrLine = 84
              SpreadFlagMessage u, "+m", MakeMsg(IIf(GetSockFlag(u, SF_Status) = SF_Status_BotGetPass, MSG_PLBotNetPassTO, MSG_PLTelNetPassTO), SocketItem(u).RegNick)
              DisconnectSocket u
  ErrLine = 88
            End If
  ErrLine = 89
      End Select
  ErrLine = 90
    End If
  Next u
Exit Sub
TevError:
'  'Stop
  PutLog "||| ] TimedEvents ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
  PutLog "||| ] Der Fehler " & CStr(Err.Number) & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & CStr(ErrLine)
  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
'  SendNote "AnGeL TimedEvents ERROR", "Hippo", "", "Der Fehler " & cStr(Err.Number) & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & cStr(ErrLine)) & " - " & Events(u).DoThis
'  SendNote "AnGeL TimedEvents ERROR", "sensei", "", "Der Fehler " & cStr(Err.Number) & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & cStr(ErrLine)) & " - " & Events(u).DoThis
  Events(u).DoThis = ""
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'  FORM ACTIVITIES
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Command1_Click()
  Unload Me
End Sub

Sub cmdConnect_Click()
  Select Case GUI_frmWinsock.cmdConnect.Caption
    Case "connect"
      ConnectServer 0, ""
    Case "Disconnect"
      DontConnect = True
      SendLine "quit :Angel leaving...", 1
      PutLog "| Disconnected from server (user clicked on 'disconnect')"
      Disconnect
      DontConnect = False
  End Select
End Sub

Private Sub Form_Resize()
  If Me.WindowState = 1 Then
    Me.Hide: Me.WindowState = 0
    If FormPassword <> "" Then HideThings.Visible = True: mnuConnect.Enabled = False: mnuEnd.Enabled = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Exitting Then
    Cancel = False
  Else
    Cancel = True
    TimedEvent "winsock2_shutdown", 0
  End If
End Sub

'User wants to connect / disconnect the bot
Private Sub mnuConnect_Click()
  If ((GUI_frmWinsock.Visible = True) And (HideThings.Visible = True)) Or ((GUI_frmWinsock.Visible = False) And (FormPassword <> "")) Then Beep: Exit Sub
  cmdConnect_Click
End Sub

'User wants to shut the bot down
Private Sub mnuEnd_Click()
  If ((GUI_frmWinsock.Visible = True) And (HideThings.Visible = True)) Or ((GUI_frmWinsock.Visible = False) And (FormPassword <> "")) Then Beep: Exit Sub
  TimedEvent "winsock2_shutdown", 0
End Sub

'User wants to see the main form
Private Sub mnuShowMain_Click()
  Dim TPPX As Long, TPPY As Long
  TPPX = Screen.TwipsPerPixelX
  TPPY = Screen.TwipsPerPixelY
  Me.WindowState = 0
  Me.Visible = True
  user32_SetWindowPos Me.hwnd, -1, Me.Left \ TPPX, Me.Top \ TPPY, Me.Width \ TPPX, Me.Height \ TPPY, &H1 Or &H2 Or &H40
  user32_SetWindowPos Me.hwnd, -2, Me.Left \ TPPX, Me.Top \ TPPY, Me.Width \ TPPX, Me.Height \ TPPY, &H1 Or &H2 Or &H40
End Sub

Private Sub TTimer_Timer()
  If Me.WindowState <> 0 Then Exit Sub
  If Me.Visible = False Then Exit Sub
  Static LastIN As Currency
  Static LastOut As Currency
  If (GlobalBytesReceived > LastIN) And (GlobalBytesSent > LastOut) Then
    Picture1.BackColor = vbGreen
    Picture2.BackColor = vbRed
    Picture1.ToolTipText = "IN: " & SizeToString(GlobalBytesReceived)
    Picture2.ToolTipText = "OUT: " & SizeToString(GlobalBytesSent)
  ElseIf (Not (GlobalBytesReceived > LastIN)) And (GlobalBytesSent > LastOut) Then
    Picture1.BackColor = vbWhite
    Picture2.BackColor = vbRed
    Picture1.ToolTipText = "IN: " & SizeToString(GlobalBytesReceived)
    Picture2.ToolTipText = "OUT: " & SizeToString(GlobalBytesSent)
  ElseIf (GlobalBytesReceived > LastIN) And (Not (GlobalBytesSent > LastOut)) Then
    Picture1.BackColor = vbGreen
    Picture2.BackColor = vbWhite
    Picture1.ToolTipText = "IN: " & SizeToString(GlobalBytesReceived)
    Picture2.ToolTipText = "OUT: " & SizeToString(GlobalBytesSent)
  ElseIf (Not (GlobalBytesReceived > LastIN)) And (Not (GlobalBytesSent > LastOut)) Then
    Picture1.BackColor = vbWhite
    Picture2.BackColor = vbWhite
  End If
  LastIN = GlobalBytesReceived
  LastOut = GlobalBytesSent
End Sub

'User pressed return in the raw textbox
Private Sub txtInput_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendIt txtInput & vbCrLf
    txtInput = ""
  End If
End Sub

'User clicked on the tray icon
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Static Message As Long, RR As Boolean
Dim TPPX As Long, TPPY As Long, EnteredPassword As String
  TPPX = Screen.TwipsPerPixelX
  TPPY = Screen.TwipsPerPixelY
  
  Message = x / TPPX
  
  If RR = False Then
    RR = True
    Select Case Message
      'Left double click (brings up the AnGeL main screen...)
      Case WM_LBUTTONDBLCLK
          Me.WindowState = 0
          Me.Visible = True
          user32_SetWindowPos Me.hwnd, -1, Me.Left \ TPPX, Me.Top \ TPPY, Me.Width \ TPPX, Me.Height \ TPPY, &H1 Or &H2 Or &H40
          user32_SetWindowPos Me.hwnd, -2, Me.Left \ TPPX, Me.Top \ TPPY, Me.Width \ TPPX, Me.Height \ TPPY, &H1 Or &H2 Or &H40
      'Right button up (brings up a context menu)
      Case WM_RBUTTONUP
          Me.PopupMenu mnuPump
    End Select
    RR = False
  End If
End Sub

'Private Property Get ICallBack_MsgResponse() As CBH.eCalling
'
'End Property
'
'Private Property Let ICallBack_MsgResponse(ByVal RHS As CBH.eCalling)
'
'End Property
'
'Private Function ICallBack_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'  ICallBack_WindowProc = WindowProc(hwnd, iMsg, wParam, lParam)
'End Function


