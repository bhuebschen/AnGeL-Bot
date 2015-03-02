Attribute VB_Name = "Kernel_Main"
',-======================- ==-- -  -
'|   AnGeL - Kernel - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


' -= Bot Version spezifische Dinge
Public Const BotVersion As String = "V1.6.2"
Public Const BotVersionEx As String = "V1.6.2 BETA10"
Public Const LongBotVersion As String = "01060200"
Public Const DebugMode As Boolean = False
Public Const UpdateNote As String = "" '"Hiya! I was just updated to 3AnGeL " & BotVersionEx & ". IMPORTANT: The flags +O and +V do no longer work as AutoOp and AutoVoice. They have been changed to +a and +v due to the introduction of user-definable capital letter flags (from +A to +Z)."
Public RunOnIDE As Boolean

Function CpString() As String ' : AddStack "Routines_CpString()"
'Dim u As Long, cr As String
'cr = "®%7556$e|""U.Raee_fYdc13_MHWEIDK"
''cr = "®%7554$e|""U.Raee_fYdc"
'  For u = 1 To Len(cr)
'    Mid(cr, u, 1) = Chr(Asc(Mid(cr, u, 1)) - 4 - u * 1.78 + u ^ 1.3)
'  Next u
'  CpString = cr
  CpString = "©1998-2003 by the AnGeL Team"
End Function


Sub Main()
  Dim Index As Long, TheFile As String, Rest As String
  
  ' -= Basis Module initialisieren
  fInitialize
  FileSys_Load
  OpSys_Load
  
  Randomize Timer
  
  ' -= Ältere Instanz suchen
  If user32_FindWindowA("ThunderRT5Form", "AnGeL Bot - " & HomeDir & App.EXEName) > 0 Then
    End
  ElseIf user32_FindWindowA("ThunderRT6FormDC", "AnGeL Bot - " & HomeDir & App.EXEName) > 0 Then
    End
  End If
  
  ' -= Komandozeile auf 'serv' prüfen
  IsNTService = False
  If LCase(Trim(Command)) = "serv" Then
    NTServiceName = GetPPString("Others", "NTService", "", AnGeL_INI)
    If NTServiceName <> "" Then
      If Not StartNTService Then End
      IsNTService = True
    End If
  End If
  
  ' -= Update überprüfen
  If UCase(App.EXEName) = "UPDATE" Then
    On Local Error Resume Next
    Index = 0
    Do
      Err.Clear
      FileCopy HomeDir & "Update.exe", HomeDir & "AnGeL.exe"
      If Err.Number > 0 Then
        Err.Clear
        Index = Index + 1
        If Index = 25 Then End
      Else
        Exit Do
      End If
    Loop
    Shell HomeDir & "AnGeL.exe"
    End
  End If
  On Local Error Resume Next
  Index = 0
  Do
    If Dir(HomeDir & "Update.exe") <> "" Then Kill HomeDir & "Update.exe"
    If Err.Number > 0 Then
      Err.Clear
      Index = Index + 1
      If Index = 25 Then Exit Do
    Else
      Exit Do
    End If
  Loop

  ' -= Auf Visual Basic prüfen
  RunOnIDE = False
  Debug.Assert SetRunOnIDE = True
  
  ' -= Restliche Module initialisieren
  TimedEvents_Load
  Botnet_Load
  Partyline_Load
  UserList_Load
  Server_Load
  Sockets_Load
  Plugins_Load
  Scripts_Load
    
  'Show first-time setup dialog if AnGeL.ini was not found
  If Dir(AnGeL_INI) = "" Then
    GUI_frmBotSetup.Show vbModal
  End If
    
  'Show error message if dialog was closed
  If Dir(AnGeL_INI) = "" Then
    MsgBox "You closed the setup window. Without completing the setup this bot will not run. Bye!", vbCritical, "Setup incomplete"
    End
  End If
    
  'Remove service entry if not running on NT
  If Not WinNTOS Then
    If GetPPString("Others", "NTService", "", AnGeL_INI) <> "" Then
      DeletePPString "Others", "NTService", AnGeL_INI
    End If
  End If
  
  'Make file structure
  If Dir(HomeDir & "ExtSeen.ini") <> "" Then Kill HomeDir & "ExtSeen.ini"
  If Not DirExist(FileAreaHome) Then MkDir FileAreaHome
  If Not DirExist(FileAreaHome & "Logs") Then MkDir FileAreaHome & "Logs"
  If Not DirExist(FileAreaHome & "Incoming") Then MkDir FileAreaHome & "Incoming"
  If Not DirExist(FileAreaHome & "Scripts") Then MkDir FileAreaHome & "Scripts"
  
  'Activate NTFS Compression for Logs and Scripts
  NTFS_CheckCompress FileAreaHome & "Logs"
  NTFS_CheckCompress FileAreaHome & "Scripts"
  
  'Delete Whatis.ini
  If Dir(HomeDir & "Whatis.ini", vbArchive) <> "" Then Kill HomeDir & "Whatis.ini"
  'Delete \Logs (the old logs contain user's passwords!)
  If DirExist(HomeDir & "Logs") Then
    TheFile = Dir(HomeDir & "Logs\*.*")
    Do
      If TheFile = "" Then Exit Do
      Kill HomeDir & "Logs\" & TheFile
      TheFile = Dir
    Loop
    RmDir HomeDir & "Logs"
  End If
  'Delete \Incoming
  If DirExist(HomeDir & "Incoming") Then
    TheFile = Dir(HomeDir & "Incoming\*.*")
    Do
      If TheFile = "" Then Exit Do
      Kill HomeDir & "Incoming\" & TheFile
      TheFile = Dir
    Loop
    RmDir HomeDir & "Incoming"
  End If
  'Delete logs older than 30 days
  LastDay = Day(Now)
  TheFile = Dir(FileAreaHome & "\Logs\*.log")
  Do
    If TheFile = "" Then Exit Do
    If CDate(LogToDate(TheFile) + 30) < Now Then Kill FileAreaHome & "Logs\" & TheFile
    TheFile = Dir
  Loop

  'Initialize bot -------------------------------------------------------------
  PutLog "¸+*°´^`°*+-> AnGeL " & BotVersionEx & " Startup <-+*°´^`°*+¸"
  
  'Initialize language
  If InitLanguage Then PutLog "|  Read language file." Else PutLog "|  Using default language."
  BaseFlags = GetPPString("Others", "BaseFlags", "+fp", AnGeL_INI)
  If AnGeLFiles.UsePolicies = False Then
    PutLog "|  Creating ScriptPolicies"
    AnGeLFiles.CreatePolicies
  Else
    AnGeLFiles.LoadPolicies
    PutLog "|  Read policy file."
  End If
  
  ReadUserList
  PutLog "|  Read userlist - " & CStr(BotUserCount) & " users."
  
  'Read AnGeL.ini
  CommandPrefix = LCase(GetPPString("Others", "CMDPrefix", "!", AnGeL_INI))
  PrimaryNick = GetPPString("Identification", "PrimaryNick", "", AnGeL_INI)
  SecondaryNick = GetPPString("Identification", "SecondaryNick", "", AnGeL_INI)
  Invisible = Switch(GetPPString("Identification", "Invisible", IIf(IsNTService, "1", "0"), AnGeL_INI))
  If PrimaryNick = "" Then MsgBox "Die Datei AnGeL.ini in meinem Verzeichnis muß einen Eintrag ""PrimaryNick=..."" im Abschnitt ""[Identification]"" besitzen!", vbCritical, "Konfigurations-Fehler!": Unload GUI_frmWinsock: Exit Sub
  If SecondaryNick = "" Then MsgBox "Die Datei AnGeL.ini in meinem Verzeichnis muß einen Eintrag ""SecondaryNick=..."" im Abschnitt ""[Identification]"" besitzen!", vbCritical, "Konfigurations-Fehler!": Unload GUI_frmWinsock: Exit Sub
  App.Title = Left(GetPPString("Identification", "AppTitle", "AnGeL - " & PrimaryNick, AnGeL_INI), 39)
  BotNetNick = GetPPString("Identification", "BotNetNick", PrimaryNick, AnGeL_INI)
  If GetPPString("Identification", "KillNickChange", "xxxx", AnGeL_INI) <> "xxxx" Then DeletePPString "Identification", "KillNickChange", AnGeL_INI
  FileAreaEnabled = Switch(GetPPString("Others", "FileArea", "yes", AnGeL_INI))
  FakeIDKick = Switch(GetPPString("Others", "FakeIDKick", "yes", AnGeL_INI))
  HideBot = Switch(GetPPString("Others", "HideBot", "no", AnGeL_INI))
  ResolveIP = GetPPString("Others", "ResolveIP", "2", AnGeL_INI)
  GUI_frmWinsock.Caption = "AnGeL Bot - " & HomeDir & App.EXEName
  IdentCommand = GetPPString("Identification", "IdentCommand", "IDENT", AnGeL_INI)
  PumpDCC = CLng(GetPPString("Others", "PumpDCC", "4096", AnGeL_INI))
  If PumpDCC > 4096 Then PumpDCC = 4096
  BanLimit = CLng(GetPPString("Others", "BanLimit", "20", AnGeL_INI))
  VersionReply = GetPPString("Server", "VersionReply", "AnGeL Bot " & BotVersion + IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", "") & " - Copyright " & CpString, AnGeL_INI)
  StrictHost = Switch(GetPPString("Others", "StrictHost", "no", AnGeL_INI))
  AllowRunToS = Switch(GetPPString("Others", "AllowRunToS", "no", AnGeL_INI))
  PortRange = GetPPString("Others", "PortRange", "0", AnGeL_INI)
  LogMaxAge = IIf(IsNumeric(GetPPString("Others", "MaxLogAge", "30", AnGeL_INI)), GetPPString("Others", "MaxLogAge", "30", AnGeL_INI), "30")
  RestrictCycle = Switch(GetPPString("Others", "RestrictCycle", "yes", AnGeL_INI))
  RouterWorkAround = Switch(GetPPString("Others", "RouterWorkAround", "yes", AnGeL_INI))
  Uptime_Enabled = Switch(GetPPString("Others", "UptimeContest", "yes", AnGeL_INI))
  
  'connect delays
  DefaultConnectDelay = CLng(GetPPString("Server", "DefaultConnectDelay", "60", AnGeL_INI))
  FailureConnectDelay = CLng(GetPPString("Server", "FailureConnectDelay", "120", AnGeL_INI))
  JumpConnectDelay = CLng(GetPPString("Server", "JumpConnectDelay", "30", AnGeL_INI))
  CurrentConnectDelay = DefaultConnectDelay
  
  'GUI password
  FormPassword = GetPPString("Others", "Password", "", AnGeL_INI)
  If FormPassword <> "" Then GUI_frmWinsock.HideThings.Visible = True: GUI_frmWinsock.mnuConnect.Enabled = False: GUI_frmWinsock.mnuEnd.Enabled = False
  
  'First time superowner keyword
  FirstTimeKeyword = LCase(GetPPString("Identification", "FirstTimeKeyword", "", AnGeL_INI))
  If FirstTimeKeyword = "" Then FirstTimeKeyword = "°NOKEY"
  
  'KI Settings
  KIAge = GetPPString("KI", "Age", "", AnGeL_INI)
  KICity = GetPPString("KI", "City", "", AnGeL_INI)
  KIFName = GetPPString("KI", "FirstName", "", AnGeL_INI)
  KILName = GetPPString("KI", "LastName", "", AnGeL_INI)
  KIGender = GetPPString("KI", "Gender", "", AnGeL_INI)
  KICountry = GetPPString("KI", "Country", "", AnGeL_INI)
  
  'Network Settings
  ServerNetwork = GetPPString("NET", "NetworkName", "IRCNet", NET_INI)
  ServerMaxChannels = CByte(GetPPString("NET", "MaxChan", "10", NET_INI))
  ServerNumberOfModes = CByte(GetPPString("NET", "NumberOfModes", "3", NET_INI))
  ServerNickLen = CByte(IIf(IsNumeric(GetPPString("NET", "NickLength", "9", NET_INI)), GetPPString("NET", "NickLength", "9", NET_INI), "9"))
  ServerUseFullAdress = CBool(GetPPString("NET", "UseFullAdress", "1", NET_INI))
  ServerSplitDetection = CBool(GetPPString("NET", "SplitDetection", "1", NET_INI))
  ServerChannelPrefixes = GetPPString("NET", "ChanPrefixes", "&#", NET_INI)
  ServerUserModes = GetPPString("NET", "UserPrefixes", "(ov)@+", NET_INI)
  ServerChannelModes = GetPPString("NET", "ChanModes", "beI,k,l,imnpsaqrt", NET_INI)
  ServerTopicLen = CInt(GetPPString("NET", "TopicLen", "80", NET_INI))
  ServerInfo.SupportsMultiKicks = (GetPPString("NET", "MultiKicks", "0", NET_INI))
  ServerInfo.SupportsMultiChanWho = (GetPPString("NET", "MultiWHO", "0", NET_INI))
  AutoNetSetup = (LCase(GetPPString("Others", "AutoNETSETUP", "yes", AnGeL_INI)) = "yes")
  
  'Auth Settings
  AuthTarget = GetPPString("AUTH", "Target", "", AnGeL_INI)
  AuthCommand = GetPPString("AUTH", "Command", "IDENTIFY", AnGeL_INI)
  AuthParam1 = GetPPString("AUTH", "Username", "", AnGeL_INI)
  AuthParam2 = GetPPString("AUTH", "Password", "", AnGeL_INI)
  AuthReAuth = (GetPPString("AUTH", "ReAuth", "1", AnGeL_INI) = "1")
  
  'IdentD
  UseIDENTD = Switch(GetPPString("Others", "UseIDENTD", "yes", AnGeL_INI))
  
  'Check flood limit
  MaxBytesToServer = CLng(GetPPString("Server", "FloodProtection", "250", AnGeL_INI))
  If MaxBytesToServer < 100 Then MaxBytesToServer = 250
  
  'Delete unregistered users from Seen.txt
  ClearSeenEntries
  
  'Read help
  InitHelp
  PutLog "|  Initialized help system - " & CStr(CommandCount) & " entries."
  
  'Check for bot update and winsock2_send update flagnote if necessary
  If Val(GetPPString("Others", "LastVersion", "", AnGeL_INI)) < Val(LongBotVersion) Then
    WritePPString "Others", "LastVersion", LongBotVersion, AnGeL_INI
    If UpdateNote <> "" Then
      For Index = 1 To BotUserCount
        If MatchFlags(BotUsers(Index).Flags, "+n") Then
          SendNote BotNetNick, BotUsers(Index).Name, "+n", UpdateNote
        End If
      Next Index
      PutLog "|  Bot update detected - sent update flagnote."
    Else
      PutLog "|  Bot update detected."
    End If
  End If
  
  'Read Ignore List
  ReadIgnores
  PutLog "|  Read " & IgnoreCount & " ignore" & IIf(IgnoreCount = 1, "", "s") & "."
  
  'Read Banlist
  ReadBans
  PutLog "|  Read " & BanCount & " ban" & IIf(BanCount = 1, "", "s") & "."
  
  'Read Exceptlist
  ReadExcepts
  PutLog "|  Read " & ExceptCount & " except" & IIf(ExceptCount = 1, "", "s") & "."
  
  'Read Invitelist
  ReadInvites
  PutLog "|  Read " & InviteCount & " invite" & IIf(InviteCount = 1, "", "s") & "."
  
  'Read Invitelist
  ReadNotes
  PutLog "|  Read " & NoteCount & " note" & IIf(NoteCount = 1, "", "s") & "."
  
  HubBot = (GetPPString("Server", "Server", "", AnGeL_INI) = "")
  MSG_TaskbarCreated = RegisterWindowMessage("TaskbarCreated")
  AddTrayIcon
  If Invisible Then
    If Not WinNTOS Then HideProcess: PutLog "|  Made process invisible."
  End If
  Status "*** AnGeL Bot startup" & vbCrLf
  
  App.TaskVisible = False
  
  PutLog "|  Hooked to Winsock."
  
  'NEW! winsock2_bind to address
  Rest = GetPPString("Identification", "LocalHost", "", AnGeL_INI)
  If Rest = "" Then
    LocalAddress = ""
  Else
    If IsValidIP(Rest) Then
      LocalAddress = Rest
    Else
      LocalAddress = GetCacheIP(Rest, True)
    End If
    If Not IsValidLocalIP(LocalAddress) Then LocalAddress = ""
  End If
  If LocalAddress <> "" Then AddressDefault = WSABuildSocketAddress(LocalAddress)
  
  BotnetPort = Val(GetPPString("TelNet", "Port", "3333", AnGeL_INI))
  TelnetPort = Val(GetPPString("TelNet", "UserPort", "23", AnGeL_INI))
  GlobalBytesSent = Val(GetPPString("Others", "GlobalBytesSent", "0", AnGeL_INI))
  GlobalBytesReceived = Val(GetPPString("Others", "GlobalBytesReceived", "0", AnGeL_INI))
  
  If TelnetPort > 0 Then
    TelnetSocket = AddSocket
    If ListenTCP(TelnetSocket, TelnetPort) = 0 Then
      SocketItem(TelnetSocket).RegNick = "<TELNET>"
      SetSockFlag TelnetSocket, SF_Status, SF_Status_TelnetListen
      Output "Listening for telnet connects on vSocket " & vbTab + Trim(Str(TelnetSocket)) & ", Port " & Trim(Str(TelnetPort)) + vbCrLf
      PutLog "|  Listening for telnet connects on port " & Trim(Str(TelnetPort)) & " (vSocket: " & Trim(Str(TelnetSocket)) & ")."
    Else
      Output "Can't listen for telnet connects - Port in use." & vbCrLf
      PutLog "|  Can't listen for telnet connects - Port in use."
      RemoveSocket TelnetSocket, 0, "", True
    End If
  End If
  If BotnetPort > 0 Then
    BotnetSocket = AddSocket
    If ListenTCP(BotnetSocket, BotnetPort) = 0 Then
      SocketItem(BotnetSocket).RegNick = "<BOTNET>"
      SetSockFlag BotnetSocket, SF_Status, SF_Status_BotnetListen
      Output "Listening for Botnet connects on vSocket " & vbTab + Trim(Str(BotnetSocket)) & ", Port " & Trim(Str(BotnetPort)) + vbCrLf
      PutLog "|  Listening for Botnet connects on port " & Trim(Str(BotnetPort)) & " (vSocket: " & Trim(Str(BotnetSocket)) & ")."
    Else
      Output "Can't listen for Botnet connects - Port in use." & vbCrLf
      PutLog "|  Can't listen for Botnet connects - Port in use."
      RemoveSocket BotnetSocket, 0, "", True
    End If
  End If
  
  UnbanTime = FindUnbanTime
  
  'Initialize botnet
  AddBot BotNetNick, "", "", LongToBase64(LongBotVersion), 0
  PutLog "|  Initialized botnet."
  
  'Read channels
  ReadAutoJoinChannels
  
  'Initialize variables used by timers
  StartUpTime = Now
  LastTick = WinTickCount
  BotRestart = False
  AutoUpdate = False
  GUI_frmWinsock.BlockResolve = False
  
  'Start timers
  Load GUI_frmWinsock
  GUI_frmWinsock.Hide
  GUI_frmWinsock.NTService.Enabled = IsNTService
  GUI_frmWinsock.BufferExTimer.Enabled = True
  GUI_frmWinsock.ChanCheck.Enabled = True
  GUI_frmWinsock.FloodTimer.Enabled = True
  GUI_frmWinsock.PingServer.Enabled = True
  GUI_frmWinsock.ResolveTimer.Enabled = True
  GUI_frmWinsock.TimedEvents.Enabled = True
  If Invisible = False Then GUI_frmWinsock.TTimer.Enabled = True
  PutLog "|  Started timers."
  
  'Load scripts
  Rest = GetPPString("Scripts", "Load", "", AnGeL_INI)
  For Index = 1 To ParamCount(Rest)
    LoadScript FileAreaHome & "Scripts\" & Param(Rest, Index)
  Next Index
  PutLog "|  Loaded scripts: " & CStr(ParamCount(Rest))
  
  
  'Initialize Uptime Module
  TimedEvent "UPTIME", 60 * 60 * 6
  PutLog "|  Started Uptime Contest."
  
  PutLog "|  Initialisation complete. Bot started."
  PutLog "`·--- Startup complete. ----------------------- ----- --  -"
  
  'Connects to hub bots
  ConnectHubs
  
  'connect to server
  If HubBot = False Then ConnectServer 0, ""
End Sub

Public Sub UnloadBot() ' : AddStack "NTServiceModule_UnloadBot()"
Dim t As Single, LineNumber As Long, u As Long
On Error GoTo SendUnloadErrorNote
  Exitting = True
  DontConnect = True
  GUI_frmWinsock.Caption = "AnGeL Bot - Unloading..."
  PutLog "¸.--- Unloading Bot ---------------------------------------------------- ----- --  -"
  If Invisible Then
    If Not WinNTOS Then UnhideProcess: PutLog "|  UnHided process."
  End If
  
  kernel32_SetUnhandledExceptionFilter 0&
  
  'Write seen entries to disk
  FlushExtSeenEntries
  
  'Remove all scripts
  For u = ScriptCount To 1 Step -1
    LineNumber = 100
    RemScript u
    LineNumber = 101
  Next u
  PutLog "|  Removed scripts."
  
  'Write userlist to disk
  WriteUserList
  
  If BotRestart Then
LineNumber = 2
    ToBotNet 0, "c " & BotNetNick & " A AnGeL Bot Restart in progress. Bye! *^^*"
LineNumber = 3
    ToBotNet 0, "bye"
LineNumber = 4
    SpreadMessage 0, -1, "4*** Bot Restart initiated"
LineNumber = 5
    SpreadMessage 0, -1, ""
LineNumber = 6
    PutLog "|  Bot is restarting!"
    If ServerSocket > 0 Then
      SendLine "quit :Restarting...", 1
    End If
  Else
    If Not AutoUpdate Then
LineNumber = 8
      ToBotNet 0, "c " & BotNetNick & " A AnGeL Bot winsock2_shutdown in progress. Bye! *^^*"
LineNumber = 9
      ToBotNet 0, "bye Bot winsock2_shutdown"
LineNumber = 10
      If IsNTService Then
        SpreadMessage 0, -1, "4*** Bot winsock2_shutdown initiated (closed by service control manager)"
      Else
        SpreadMessage 0, -1, "4*** Bot winsock2_shutdown initiated (closed by host)"
      End If
LineNumber = 11
      SpreadMessage 0, -1, MSG_ThankYou
LineNumber = 12
      SpreadMessage 0, -1, ""
LineNumber = 13
      If IsNTService Then
        PutLog "|  Bot is being closed by service control manager!"
      Else
        PutLog "|  Bot is being closed by host!"
      End If
      If ServerSocket > 0 Then
        SendLine "quit :" & GetPPString("Server", "QuitMessage", "AnGeL leaving...", AnGeL_INI), 1
      End If
    Else
LineNumber = 14
      ToBotNet 0, "c " & BotNetNick & " A AnGeL AutoUpdate in progress. Bye! *^^*"
LineNumber = 15
      ToBotNet 0, "bye AnGeL AutoUpdate in progress"
LineNumber = 16
      PutLog "|  Bot is AutoUpdating!"
      SendLine "quit :AutoUpdate in progress...", 1
    End If
  End If
  If ServerSocket > 0 Then
LineNumber = 17
    t = Timer: Do: DoEvents: Loop Until Timer - t > 1
  End If
LineNumber = 19
  Disconnect
  RemTrayIcon
LineNumber = 21
  Sockets_Unload
'  RemCB GUI_frmWinsock, GUI_frmWinsock.hwnd, Socket_UDP_EVENT
'  RemCB GUI_frmWinsock, GUI_frmWinsock.hwnd, 1026
'  RemCB GUI_frmWinsock, GUI_frmWinsock.hwnd, 1025
  PutLog "|  Unhooked from Winsock."
LineNumber = 22
LineNumber = 23
  'Restart Bot if needed
  If IsNTService Then
    PutLog "`·--- Good bye. ---------------------------------------------------------- ----- --  -"
    Unload GUI_frmWinsock
    StopService
    If BotRestart Then Shell "net start " & GetPPString("Others", "NTService", "", AnGeL_INI), vbHide
  ElseIf Not IsNTService And BotRestart Then
    PutLog "|  Executing AnGeL Binary...": Shell HomeDir & App.EXEName & ".exe"
  End If
  PutLog "`·--- Good bye. ---------------------------------------------------------- ----- --  -"
  Unload GUI_frmWinsock
  fTerminate
  End
  'Kill App.hInstance
  Exit Sub
SendUnloadErrorNote:
'  PutLog "||| ] Unload ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
'  PutLog "||| ] Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & cStr(LineNumber))
'  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
'  SendNote "AnGeL Unload ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & cStr(LineNumber))
'  SendNote "AnGeL Unload ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & cStr(LineNumber))
'  user32_SetWindowLongA GUI_frmWinsock.hwnd, GWL_WNDPROC, OldWndProc
  Sockets_Unload
  If BotRestart Then PutLog "|  ERROR in UnloadBot! Trying to restart before it's too late...": Shell HomeDir & "Angel.exe"
  PutLog "`·--- Good bye. ---------------------------------------------------------- ----- --  -"
  Unload GUI_frmWinsock
End Sub

Function SetRunOnIDE() As Boolean
  RunOnIDE = True
  SetRunOnIDE = RunOnIDE
End Function

