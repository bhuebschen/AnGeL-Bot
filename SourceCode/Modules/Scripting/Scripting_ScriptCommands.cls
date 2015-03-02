VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scripting_ScriptCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
',-——————————————————- -- -  -
'|   AnGeL Script Module
'| © 2000 by Tibor Schiemann
'| © 2003 by Benedikt Hübschen
'|-—————————- -- -  -
'| 12.02.2003 - Benedikt Hübschen
'|     Scripting Policies!
'|
'| 12.02.2003 - Benedikt Hübschen
'|     LoadDSN, SQLExecute, CloseDSN, Format, DoEvent,
'|     CurrentTime, CurrentDate, TickCount, SendNote,
'|     AddUser, RemUser
'|
'| 14.02.2003 - Benedikt Hübschen
'|     MyVersion, MyVersionEx, MatchPassword, NTFSCompress,
'|     NTFSEncrypt, PlatformType, WindowsVersion, NTFSDecrypt,
'|     NTFSDeCompress, MakeSessionInValid, MakeSessionValid - POLICIES ADDED! -
'|
'| 16.02.2003 - Benedikt Hübschen
'|     WMI, LoadAddIn
'|
'| 19.02.2003 - Benedikt Hübschen
'|     SockListenPortInUse
'`-—————————- -- -  -
Option Explicit

Public CurrentScript As Long
Public MyAnswer As String
Public CurrentCommand As String
Dim BlowFish As clsBlowfish

Public Function AddUser(UserName) As Byte: CurrentCommand = "AddUser(" & UserName & ")"
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    AddUser = Userlist_BotUserlist.AddUser(CStr(UserName), BaseFlags)
  Else
    AddUser = AU_DENIED
  End If
End Function

Public Function LoadAddIn(ModuleName) As Object: CurrentCommand = "LoadAddIn(" & ModuleName & ")"
  If AnGeLFiles.CommandAllowed("AddinModules") Then
    On Local Error Resume Next
    Dim dll As String, F As Integer, dllContent$
    Dim signatur As String * 144
    dll = GetRegString(HKEY_CLASSES_ROOT, "CLSID\" & GetRegString(HKEY_CLASSES_ROOT, "AnGeL." & ModuleName & "\CLSID", "") & "\InprocServer32", "")
    If Dir(dll) = "" Then
      Partyline_Functions.SpreadFlagMessage 0, "+m", "*** LoadAddin failed: File not found"
    Else
      Set BlowFish = New clsBlowfish
      F = FreeFile: Open dll For Binary As #F
        dllContent = String(LOF(F) - 144, Chr(0))
        Get #F, , dllContent
        Get #F, , signatur
        InitHashtable
      Close #F
      If SHA256(Mid(signatur, 1, 112)) = Mid(signatur, 113) Then
        dll = BlowFish.DecryptString(Mid(signatur, 1, 112), DecryptString("1C2660C2D6", "AnGeL"))
        If Mid(dll, 1, 32) = SHA256(dllContent) Then
          Set LoadAddIn = Interaction.CreateObject("AnGeL." & ModuleName)
          LoadAddIn.Init New Scripting_ScriptCommands, CurrentScript
          If Err.Number <> 0 Then
            Partyline_Functions.SpreadFlagMessage 0, "+m", "*** LoadAddin failed: (" & Err.Number & ") " & Err.Description
            Set LoadAddIn = Nothing
            Err.Clear
          End If
        Else
          Partyline_Functions.SpreadFlagMessage 0, "+m", "*** LoadAddin failed: Module has been modified or is corrupt"
        End If
      Else
        Partyline_Functions.SpreadFlagMessage 0, "+m", "*** LoadAddin failed: Module has been modified or is corrupt"
      End If
      Set BlowFish = Nothing
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function RemUser(UserName) As Byte: CurrentCommand = "RemUser(" & UserName & ")"
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    RemUser = Userlist_BotUserlist.RemUser(CStr(UserName))
  Else
    RemUser = RU_DENIED
  End If
End Function

Public Function LoadWMI(Optional ComputerName = ".") As Object: CurrentCommand = "LoadWMI(" & ComputerName & ")"
  If AnGeLFiles.CommandAllowed("WMI") Then
    Dim TempObj As Object
    Dim FileName As String
    Set TempObj = Interaction.GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & ComputerName & "\root\cimv2")
    Set LoadWMI = TempObj
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function LoadDSN(Name) As Object: CurrentCommand = "LoadDSN(" & Name & ")"
  If AnGeLFiles.CommandAllowed("Database") Then
    Dim TempObj As Object
    Dim FileName As String
    Set TempObj = Interaction.CreateObject("ADODB.Connection")
    FileName = Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("\", (Name), "\")
    On Local Error Resume Next
    TempObj.Open "FILEDSN=" & FileName
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** AnGeL Database System: Error (" & Hex(Err.Number) & ") " & Err.Description
      Set LoadDSN = Nothing
      Exit Function
    End If
    Set LoadDSN = TempObj
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function SQLExecute(Connection, Query) As Object: CurrentCommand = "SQLExecute(Connection, " & Query & ")"
  If AnGeLFiles.CommandAllowed("Database") Then
    On Local Error Resume Next
    Set SQLExecute = Connection.Execute(Query)
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** AnGeL Database System: Error (" & Hex(Err.Number) & ") " & Err.Description
      Set SQLExecute = Nothing
      Exit Function
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub CloseDSN(Connection): CurrentCommand = "CloseDSN(Connection)"
  If AnGeLFiles.CommandAllowed("Database") Then
    Connection.Close
    Set Connection = Nothing
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Function AFormat(Expression, Optional Format, Optional FirstDayOfWeek = vbSunday, Optional FirstWeekOfYear = vbFirstJan1) As String: CurrentCommand = "AFormat(" & Expression & ", " & Format & ", " & FirstDayOfWeek & ", " & FirstWeekOfYear & ")"
  On Local Error Resume Next
  AFormat = Strings.Format(Expression, Format, FirstDayOfWeek, FirstWeekOfYear)
  If Err.Number <> 0 Then
    If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** AFormat: Error (" & Err.Number & ") " & Err.Description
  End If
End Function

Public Sub DoEvent(): CurrentCommand = "DoEvent()"
  DoEvents
End Sub

Public Function CurrentDate() As Variant: CurrentCommand = "CurrentDate()"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    CurrentDate = Date
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function CurrentTime() As Variant: CurrentCommand = "CurrentTime()"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    CurrentTime = Time
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function CurrentTimer() As Single: CurrentCommand = "CurrentTimer()"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    CurrentTimer = Timer
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function TickCount() As Currency: CurrentCommand = "TickCount()"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    TickCount = WinTickCount
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function SendNote(RegUser, NoteText) As Boolean: CurrentCommand = "SendNote(" & RegUser & ", " & NoteText & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    SendNote = Plugin_Notes.SendNote(Scripts(CurrentScript).Name, CStr(RegUser), "", CStr(NoteText))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub DontLoadScript(Reason As String)
  Scripting_Scripts.DontLoadScript = Reason
End Sub

Public Function CommandPrefix() As String
  CommandPrefix = Server_Connection.CommandPrefix
End Function

Public Function GetObject(PathName As String, Class As String) As Object: CurrentCommand = "GetObject(" & PathName & ", " & Class & ")"
  If AnGeLFiles.CommandAllowed("Objects") Then
    GetObject = Interaction.GetObject(PathName, Class)
  Else
    GetObject = vbNullString
  End If
End Function

Public Function CreateObject(Class As String, Optional ServerName As String = "") As Object: CurrentCommand = "CreateObject(" & Class & ", " & ServerName & ")"
  If AnGeLFiles.CommandAllowed("Objects") Then
    CreateObject = Interaction.CreateObject(Class, ServerName)
  Else
    CreateObject = vbNullString
  End If
End Function

Public Function IIf(Statement As Boolean, TruePart As Variant, FalsePart As Variant) As Variant
  If Statement = True Then IIf = TruePart Else IIf = FalsePart
End Function

Public Function LongBotVersion() As Long: CurrentCommand = "LongBotVersion()"
  On Local Error Resume Next
  LongBotVersion = CLng(Kernel_Main.LongBotVersion)
End Function

Public Function MyVersion() As String
  On Local Error Resume Next
  MyVersion = (Kernel_Main.BotVersion)
End Function

Public Function MyVersionEx() As String
  On Local Error Resume Next
  MyVersionEx = (Kernel_Main.BotVersionEx)
End Function

Public Function Strip(Line) As String: CurrentCommand = "Strip(" & Line & ")"
  On Local Error Resume Next
  Strip = Server_Functions.Strip((Line))
End Function

Public Function AddHost(RegUser, Host) As Byte: CurrentCommand = "AddHost(" & RegUser & ", " & Host & ")"
  On Local Error Resume Next
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    AddHost = Userlist_BotUserlist.AddHost(0, (RegUser), (Host))
  Else
    AddHost = AH_DENIED
  End If
End Function

Public Function RemHost(RegUser, Host) As Byte: CurrentCommand = "RemHost(" & RegUser & ", " & Host & ")"
  On Local Error Resume Next
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    RemHost = Userlist_BotUserlist.RemHost((RegUser), (Host))
  Else
    RemHost = RH_DENIED
  End If
End Function

Public Function Chattr(RegUser, ChangeLine) As Byte: CurrentCommand = "Chattr(" & RegUser & ", " & ChangeLine & ")"
  On Local Error Resume Next
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    Chattr = Userlist_BotUserlist.Chattr((RegUser), (ChangeLine))
  Else
    Chattr = CH_DENIED
  End If
End Function

Public Sub UpdateRegUsers(Action): CurrentCommand = "UpdateRegUsers(" & Action & ")"
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    UpdateRegUsers (Action)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub AddResolve(Host): CurrentCommand = "AddResolve(" & Host & ")"
    On Local Error Resume Next
  If AnGeLFiles.CommandAllowed("Resolves") Then
    Dim GetIPResult As String
    If Scripts(CurrentScript).Hooks.Resolves = False Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** AddResolve: Hook 'Resolves' is missing! (Script: " & Trim(Str(CurrentScript)) & ")"
      Exit Sub
    End If
    GetIPResult = Sockets_Resolve.GetCacheIP((Host), False)
    If GetIPResult <> "" Then
      RunScriptX CurrentScript, "Resolves", Host, GetIPResult
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SendFile(Nick, Name): CurrentCommand = "SendFile(" & Nick & ", " & Name & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    Dim FileName As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    If (InitiateDCCSend((Nick), (Nick), (FileName)) = False) And (Scripts(CurrentScript).Silent = False) Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** SendFile: Couldn't initiate DCC send - file '" & Name & "' not found!"
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function IsValidIP(Host) As Boolean: CurrentCommand = "IsValidIP(" & Host & ")"
  On Local Error Resume Next
  IsValidIP = Sockets_Winsock2.IsValidIP((Host))
End Function

Public Function SockListenPortInUse(Port) As Boolean: CurrentCommand = "SockListenPortInUse(" & Port & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    SockListenPortInUse = Sockets_Winsock2.IsListenPortInUse(CLng(Port))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function SockListen(Port, ReplyToSub) As Long: CurrentCommand = "SockListen(" & Port & ", " & ReplyToSub & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    Dim NewSock As Long
    On Local Error Resume Next
    If Not IsNumeric(Port) Then
      SockListen = 0
      Exit Function
    End If
    NewSock = AddSocket
    If ListenTCP(NewSock, ((Port))) = 0 Then
      SetSockFlag NewSock, SF_Status, SF_Status_ScriptSocket
      SocketItem(NewSock).CurrentQuestion = (ReplyToSub)
      SocketItem(NewSock).SetupChan = Scripts(CurrentScript).Name
      SocketItem(NewSock).UserNum = 0
      SocketItem(NewSock).OnBot = ""
      SocketItem(NewSock).IRCNick = ""
      SocketItem(NewSock).RegNick = "<SCRIPT>"
      SocketItem(NewSock).IsInternalSocket = True
      SockListen = NewSock
    Else
      SockListen = 0
      RemoveSocket NewSock, 0, "", True
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function SockOpenUDP(Port, ReplyToSub) As Long: CurrentCommand = "SockListenUDP(" & Port & ", " & ReplyToSub & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    Dim NewSock As Long
    Dim u As Long
    On Local Error Resume Next
    If Not IsNumeric((Port)) Then
      SockOpenUDP = 0
      Exit Function
    End If
    NewSock = AddSocket
    If ListenUDP(NewSock, (Port)) = 0 Then
      SetSockFlag NewSock, SF_Status, SF_Status_ScriptSocket
      SocketItem(NewSock).CurrentQuestion = (ReplyToSub)
      SocketItem(NewSock).SetupChan = Scripts(CurrentScript).Name
      SocketItem(NewSock).UserNum = 0
      SocketItem(NewSock).OnBot = ""
      SocketItem(NewSock).IRCNick = ""
      SocketItem(NewSock).RegNick = "<SCRIPT>"
      SocketItem(NewSock).Used = True
      SocketItem(NewSock).IsInternalSocket = True
      SockOpenUDP = NewSock
    Else
      SockOpenUDP = 0
      RemoveSocket NewSock, 0, "", True
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function SockConnect(Host, Port, ReplyToSub) As Long: CurrentCommand = "SockConnect(" & Host & ", " & Port & ", " & ReplyToSub & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    Dim NewSock As Long, strRemoteIP As String
    On Local Error Resume Next
    NewSock = AddSocket
    If Not IsNumeric((Port)) Then
      SockConnect = 0
      Exit Function
    End If
    If ConnectTCP(NewSock, (Host), (Port)) = 0 Then
      SetSockFlag NewSock, SF_Status, SF_Status_ScriptSocket
      SocketItem(NewSock).CurrentQuestion = (ReplyToSub)
      SocketItem(NewSock).SetupChan = Scripts(CurrentScript).Name
      SocketItem(NewSock).UserNum = 0
      SocketItem(NewSock).OnBot = ""
      SocketItem(NewSock).IRCNick = ""
      SocketItem(NewSock).RegNick = "<SCRIPT>"
      SocketItem(NewSock).IsInternalSocket = True
      SockConnect = NewSock
    Else
      SockConnect = 0
      RemoveSocket NewSock, 0, "", True
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub SockWrite(vsock, What): CurrentCommand = "SockWrite(" & vsock & ", " & What & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      SendTCP (vsock), (What)
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** SockWrite failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SockWriteUDP(vsock, IP, Port, What): CurrentCommand = "SockWriteUDP(" & vsock & ", " & What & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      SendUDP (vsock), (IP), (Port), (What)
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** SockWrite failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SockClose(vsock): CurrentCommand = "SockClose(" & vsock & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      If SocketItem((vsock)).SetupChan = Scripts(CurrentScript).Name Then
        SocketItem((vsock)).SetupChan = ""
          DisconnectSocket (vsock)
      End If
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** SockClose failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function GetSockTag(vsock) As String: CurrentCommand = "GetSockTag(" & vsock & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      GetSockTag = SocketItem((vsock)).SockTag
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** GetSockTag failed: No such vsock - " & CStr(vsock) & ".": Exit Function
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub SetSockTag(vsock, TagValue): CurrentCommand = "SetSockTag(" & vsock & ", " & TagValue & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      SocketItem((vsock)).SockTag = (TagValue)
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** SetSockTag failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub ExecuteCommand(vsock, Command, Optional ReplyToSub = "", Optional UseNick = ""): CurrentCommand = "ExecuteCommand(" & vsock & ", " & Command & ", " & CStr(ReplyToSub) & ", " & CStr(UseNick) & ")"
  If AnGeLFiles.CommandAllowed("ExecuteCommand") Then
    Dim NewSock As Long
    On Local Error Resume Next

    If vsock > 0 Then
      'Script just wants a certain user to execute a command
      If IsValidSocket((vsock)) Then
        Party vsock, SocketItem(vsock).SocketNumber, (Command)
      Else
        If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** ExecuteCommand failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
      End If
    Else
      'The script itself wants to execute a command
      NewSock = AddSocket
      If UseNick = "" Then UseNick = Scripts(CurrentScript).Name
      SocketItem(NewSock).RegNick = (UseNick)
      SocketItem(NewSock).Flags = "fijmnoprstvwx"
      SocketItem(NewSock).IRCNick = "²*SCRIPT*²"
      SocketItem(NewSock).CurrentQuestion = (ReplyToSub)
      SocketItem(NewSock).SetupChan = Scripts(CurrentScript).Name
      SocketItem(NewSock).SocketNumber = 0
      SocketItem(NewSock).UserNum = 0
      SocketItem(NewSock).OnBot = ""
      SetSockFlag NewSock, SF_Status, SF_Status_Party
      Party NewSock, 0, (Command)
      SocketItem(NewSock).IsInternalSocket = False
      RemoveSocket NewSock, 0, "", True
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub GrabUser(vsock, Reason, ReplyToSub): CurrentCommand = "GrabUser(" & vsock & ", " & Reason & ", " & ReplyToSub & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      If ReplyToSub = "" Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** GrabUser failed: ReplyToSub parameter must not be empty.": Exit Sub
      If GetSockFlag((vsock), SF_Status) = SF_Status_ScriptUser Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** GrabUser failed: Vsock " & CStr(vsock) & " is already owned by the script.": Exit Sub
      'Store old SockFlag in FileName (well, it's unused ;))
      SocketItem((vsock)).FileName = GetSockFlag((vsock), SF_Status)
      'Set away and give user to the script
      Partyline_Functions.SetAway ((vsock)), (Reason)
      SetSockFlag (vsock), SF_Status, SF_Status_ScriptUser
      SocketItem((vsock)).CurrentQuestion = (ReplyToSub)
      SocketItem((vsock)).SetupChan = Scripts(CurrentScript).Name
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** GrabUser failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub ReleaseUser(vsock): CurrentCommand = "ReleaseUser(" & vsock & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) Then
      If GetSockFlag((vsock), SF_Status) <> SF_Status_ScriptUser Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** ReleaseUser failed: Vsock " & CStr(vsock) & " is not owned by the script.": Exit Sub
      SetSockFlag (vsock), SF_Status, SocketItem(vsock).FileName
      Partyline_Functions.SetAway (vsock), ""
      SocketItem(vsock).CurrentQuestion = ""
    Else
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** ReleaseUser failed: No such vsock - " & CStr(vsock) & ".": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function Hostmask(Chan, Nick) As String: CurrentCommand = "Hostmask(" & Chan & ", " & Nick & ")"
  If AnGeLFiles.CommandAllowed("UserManagement") Then
    Dim ChNum As Long, UsNum As Long
    On Local Error Resume Next
    ChNum = FindChan((Chan))
    If ChNum = 0 Then Hostmask = "": Exit Function
    UsNum = FindUser((Nick), ChNum)
    If UsNum = 0 Then Hostmask = "": Exit Function
    Hostmask = Channels(ChNum).User(UsNum).Hostmask
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function Mask(Hostmask, Nr) As String: CurrentCommand = "Mask(" & Hostmask & ", " & Nr & ")"
  On Local Error Resume Next
  Mask = Userlist_Functions.Mask((Hostmask), (Nr))
End Function

Public Function VSockInfo(vsock, GetWhat As Byte): CurrentCommand = "VSockInfo(" & vsock & ", " & GetWhat & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) = False Then
      Select Case GetWhat
        Case 5, 7: VSockInfo = 0
        Case Else: VSockInfo = ""
      End Select
      Exit Function
    End If
    Select Case GetWhat
      Case 1 'VS_RegNick
        VSockInfo = SocketItem((vsock)).RegNick
      Case 2 'VS_Flags
        VSockInfo = SocketItem((vsock)).Flags
      Case 3 'VS_IRCNick
        VSockInfo = SocketItem((vsock)).IRCNick
      Case 4 'VS_Hostmask
        VSockInfo = SocketItem((vsock)).Hostmask
      Case 5 'VS_UserNum
        VSockInfo = SocketItem((vsock)).UserNum
      Case 6 'VS_AwayMessage
        VSockInfo = SocketItem((vsock)).AwayMessage
      Case 7 'VS_IdleTime
        VSockInfo = GetSecondsTillNow(SocketItem((vsock)).LastEvent)
      Case Else
        VSockInfo = ""
    End Select
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function ChanUser(ChNum, Nr, GetWhat As Byte): CurrentCommand = "ChanUser(" & ChNum & ", " & Nr & ", " & GetWhat & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    On Local Error Resume Next
    If CLng(ChNum) > Server_Channels.ChanCount Then ChanUser = "": Exit Function
    If CLng(Nr) > Server_Channels.Channels((ChNum)).UserCount Then ChanUser = "": Exit Function
    Select Case GetWhat
      Case 1 'CU_Nick
        ChanUser = Server_Channels.Channels((ChNum)).User((Nr)).Nick
      Case 2 'CU_RegUser
        ChanUser = Server_Channels.Channels((ChNum)).User((Nr)).RegNick
      Case 3 'CU_Hostmask
        ChanUser = Server_Channels.Channels((ChNum)).User((Nr)).Hostmask
      Case 4 'CU_IPMask
        ChanUser = Server_Channels.Channels((ChNum)).User((Nr)).IPmask
      Case 5 'CU_Status
        ChanUser = Server_Channels.Channels((ChNum)).User((Nr)).Status
      Case 6 'CU_IdleTime
        ChanUser = GetSecondsTillNow2(Server_Channels.Channels((ChNum)).User((Nr)).LastEvent)
      Case Else
        ChanUser = ""
    End Select
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function ChanBan(ChNum, Nr): CurrentCommand = "ChanBan(" & ChNum & ", " & Nr & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    On Local Error Resume Next
    If CLng(ChNum) > Server_Channels.ChanCount Then ChanBan = "": Exit Function
    If CLng(Nr) > Server_Channels.Channels((ChNum)).BanCount Then ChanBan = "": Exit Function
    ChanBan = Server_Channels.Channels((ChNum)).BanList((Nr)).Mask
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function BotUser(Nr, GetWhat As Byte, Optional OptParam): CurrentCommand = "BotUser(" & Nr & ", " & GetWhat & ", " & CStr(OptParam) & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    Dim u As Long
    On Local Error Resume Next
    If CLng(Nr) > BotUserCount Then BotUser = "": Exit Function
    Select Case GetWhat
      Case 1 'BU_RegNick
        BotUser = BotUsers((Nr)).Name
      Case 2 'BU_Flags
        BotUser = BotUsers((Nr)).Flags
      Case 3 'BU_Password
        BotUser = BotUsers((Nr)).Password
      Case 4 'BU_ChanFlags
        If CStr(OptParam) = "" Then BotUser = "": Exit Function
        For u = 1 To BotUsers((Nr)).ChannelFlagCount
          If CStr(OptParam) = BotUsers((Nr)).ChannelFlags(u).Channel Then
            BotUser = BotUsers((Nr)).ChannelFlags(u).Flags
            Exit Function
          End If
        Next u
      Case 5 'BU_Hostmasks
        If CLng(OptParam) > BotUsers((Nr)).HostMaskCount Then BotUser = "": Exit Function
        If CLng(OptParam) = 0 Then BotUser = BotUsers((Nr)).HostMaskCount: Exit Function
        BotUser = BotUsers((Nr)).HostMasks((OptParam))
      Case 6 'BU_ValidSession
        BotUser = CBool(BotUsers((Nr)).ValidSession)
      Case Else
        BotUser = ""
    End Select
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function ChanNum(Name) As Long: CurrentCommand = "ChanNum(" & Name & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    On Local Error Resume Next
    ChanNum = Server_Functions.FindChan((Name))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function ChanUserNum(Chan, Nick) As Long: CurrentCommand = "ChanUserNum(" & Chan & ", " & Nick & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    Dim ChNum As Long, u As Long
    On Local Error Resume Next
    ChNum = FindChan((Chan))
    If ChNum = 0 Then ChanUserNum = 0: Exit Function
    ChanUserNum = FindUser((Nick), ChNum)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function IsLocalUser(vsock) As Boolean: CurrentCommand = "IsLocalUser(" & vsock & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) = False Then IsLocalUser = False: Exit Function
    IsLocalUser = (GetSockFlag((vsock), SF_LocalVisibleUser) = SF_YES)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub SetAway(vsock, Reason): CurrentCommand = "SetAway(" & vsock & ", " & Reason & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    If IsValidSocket((vsock)) = False Then Exit Sub
    Partyline_Functions.SetAway ((vsock)), (Reason)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub WaitWithSend(): CurrentCommand = "WaitWithSend()"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    Server_Connection.WaitThisLine = True
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SendNow(): CurrentCommand = "SendNow()"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    Server_Connection.WaitThisLine = False
    Server_Functions.BufferEx
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SetTime(NewTime): CurrentCommand = "SetTime(" & NewTime & ")"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    On Local Error Resume Next
    Date = CDate(NewTime)
    Time = CDate(NewTime)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub TimedEvent(DoThis, AtTime): CurrentCommand = "TimedEvent(" & DoThis & ", " & AtTime & ")"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    On Local Error Resume Next
    If LCase(Param(DoThis, 1)) = "callscript" Then SpreadFlagMessage 0, "+n", "5*** TimedEvent (" & DoThis & ") failed: Please use TimedCommand.": Exit Sub
    If IsDate(AtTime) = True Then
      TimedEvents_Events.TimedEvent (DoThis), GetSecondsTill(CDate(AtTime))
    Else
      TimedEvents_Events.TimedEvent (DoThis), CLng(AtTime)
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub TimedCommand(DoThis, AtTime): CurrentCommand = "TimedCommand(" & DoThis & ", " & AtTime & ")"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    On Local Error Resume Next
    If AnGeLFiles.CommandAllowed("Objects") = False Then
      If InStr(LCase(DoThis), "createobject") > 0 Then SpreadFlagMessage 0, "+n", "5*** TimedCommand failed (CreateObject may not be used in scripts)": Exit Sub
      If InStr(LCase(DoThis), "getobject") > 0 Then SpreadFlagMessage 0, "+n", "5*** TimedCommand failed (GetObject may not be used in scripts)": Exit Sub
    End If
    If IsDate(AtTime) = True Then
      TimedEvents_Events.TimedEvent "CallScript " & Scripts(CurrentScript).Name & " " & (DoThis), GetSecondsTill(CDate(AtTime))
    Else
      TimedEvents_Events.TimedEvent "CallScript " & Scripts(CurrentScript).Name & " " & (DoThis), CLng(AtTime)
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function MatchGrade(Str1, Str2) As Long: CurrentCommand = "MatchGrade(" & Str1 & ", " & Str2 & ")"
  On Local Error Resume Next
  MatchGrade = Server_Functions.MatchGrade((Str1), (Str2))
End Function

Public Function TimeSpan(LastTime) As String: CurrentCommand = "TimeSpan(" & LastTime & ")"
  If AnGeLFiles.CommandAllowed("TimeOperations") Then
    On Local Error Resume Next
    TimeSpan = TimedEvents_Functions.TimeSpan((LastTime))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function GetUserFlags(RegUser) As String: CurrentCommand = "GetUserFlags(" & RegUser & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    GetUserFlags = Userlist_BotUserlist.GetUserFlags((RegUser))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function GetUserNum(RegUser) As Long: CurrentCommand = "GetUserNum(" & RegUser & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    GetUserNum = Userlist_BotUserlist.GetUserNum((RegUser))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function GetUserChanFlags(RegUser, Chan) As String: CurrentCommand = "GetUserChanFlags(" & RegUser & ", " & Chan & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
  On Local Error Resume Next
  GetUserChanFlags = Userlist_BotUserlist.GetUserChanFlags((RegUser), (Chan))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function SearchUserFromHostmask(Hostmask) As String: CurrentCommand = "SearchUserFromHostmask(" & Hostmask & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
  On Local Error Resume Next
  SearchUserFromHostmask = Userlist_Functions.SearchUserFromHostmask((Hostmask))
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function MatchWM(MatchString, Hostmask) As Boolean: CurrentCommand = "MatchWM(" & MatchString & ", " & Hostmask & ")"
  On Local Error Resume Next
  MatchWM = Userlist_Functions.MatchWM((MatchString), (Hostmask))
End Function

Public Function MatchHost(MatchString, Hostmask) As Boolean: CurrentCommand = "MatchHost(" & MatchString & ", " & Hostmask & ")"
  On Local Error Resume Next
  MatchHost = Userlist_Functions.MatchHost((MatchString), (Hostmask))
End Function

Public Function MatchFlags(Flags, MatchString) As Boolean: CurrentCommand = "MatchFlags(" & (Flags) & ", " & (MatchString) & ")"
  On Local Error Resume Next
  MatchFlags = Userlist_Functions.MatchFlags((Flags), (MatchString))
End Function

Public Function MatchPassword(UserNum, Password) As Boolean
  On Local Error Resume Next
  MatchPassword = (Kernel_Functions.EncryptIt(CStr(Password)) = BotUser(UserNum, BU_Password))
End Function

Public Function IsOp(Nick, Chan) As Boolean: CurrentCommand = "IsOp(" & Nick & ", " & Chan & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    Dim ChNum As Long, UsNum As Long
    On Local Error Resume Next
    ChNum = Server_Functions.FindChan((Chan))
    If ChNum = 0 Then IsOp = False: Exit Function
    UsNum = Server_Channels.FindUser((Nick), ChNum)
    If UsNum = 0 Then IsOp = False: Exit Function
    If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 Then IsOp = True: Exit Function
    IsOp = False
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function IsHalfOp(Nick, Chan) As Boolean: CurrentCommand = "IsHalfOp(" & Nick & ", " & Chan & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    Dim ChNum As Long, UsNum As Long
    On Local Error Resume Next
    ChNum = Server_Functions.FindChan((Chan))
    If ChNum = 0 Then IsHalfOp = False: Exit Function
    UsNum = Server_Channels.FindUser((Nick), ChNum)
    If UsNum = 0 Then IsHalfOp = False: Exit Function
    If InStr(Channels(ChNum).User(UsNum).Status, "%") > 0 Then IsHalfOp = True: Exit Function
    IsHalfOp = False
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function IsVoice(Nick, Chan) As Boolean: CurrentCommand = "IsVoice(" & Nick & ", " & Chan & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    Dim ChNum As Long, UsNum As Long
    On Local Error Resume Next
    ChNum = Server_Functions.FindChan((Chan))
    If ChNum = 0 Then IsVoice = False: Exit Function
    UsNum = Server_Channels.FindUser((Nick), ChNum)
    If UsNum = 0 Then IsVoice = False: Exit Function
    If InStr(Channels(ChNum).User(UsNum).Status, "+") > 0 Then IsVoice = True: Exit Function
    IsVoice = False
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function IsOn(Nick, Chan) As Boolean: CurrentCommand = "IsOn(" & Nick & ", " & Chan & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    Dim ChNum As Long, UsNum As Long
    On Local Error Resume Next
    ChNum = Server_Functions.FindChan((Chan))
    If ChNum = 0 Then IsOn = False: Exit Function
    UsNum = Server_Channels.FindUser((Nick), ChNum)
    If UsNum = 0 Then IsOn = False: Exit Function
    IsOn = True
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub Script(Description): CurrentCommand = "Script(" & Description & ")"
  On Local Error Resume Next
  Scripts(CurrentScript).Description = (Description)
End Sub

Public Property Get KIAnswer()
  KIAnswer = MyAnswer
End Property

Public Property Let KIAnswer(NewAnswer)
  MyAnswer = NewAnswer
End Property

Public Sub HaltDefault(): CurrentCommand = "HaltDefault()"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    On Local Error Resume Next
    Scripting_Scripts.HaltDefault = True
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub Hook(HookTo): CurrentCommand = "Hook(" & HookTo & ")"
  If AnGeLFiles.CommandAllowed("NotificationHooks") Then
    On Local Error Resume Next
    If Not IsProcedure(CurrentScript, Param((HookTo), 1)) Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Script Error (" & Scripts(CurrentScript).Name & "): Hook couldn't find sub '" & (HookTo) & "'.": Exit Sub
    Select Case LCase(Param((HookTo), 1))
      'Entsprechend auch in Scripting.AddScript ändern!
      Case "chan_msg": Scripts(CurrentScript).Hooks.Chan_msg = True
      Case "priv_msg": Scripts(CurrentScript).Hooks.Priv_msg = True
      Case "chan_act": Scripts(CurrentScript).Hooks.Chan_act = True
      Case "priv_act": Scripts(CurrentScript).Hooks.Priv_act = True
      Case "priv_ctcp": Scripts(CurrentScript).Hooks.Priv_ctcp = True
      Case "priv_ctcpreply": Scripts(CurrentScript).Hooks.Priv_ctcpreply = True
      Case "chan_ctcp": Scripts(CurrentScript).Hooks.Chan_ctcp = True
      Case "chan_notice": Scripts(CurrentScript).Hooks.Chan_notice = True
      Case "priv_notice": Scripts(CurrentScript).Hooks.Priv_notice = True
      Case "server_notice": Scripts(CurrentScript).Hooks.Server_notice = True
      Case "join": Scripts(CurrentScript).Hooks.Join = True
      Case "part": Scripts(CurrentScript).Hooks.Part = True
      Case "quit": Scripts(CurrentScript).Hooks.Quit = True
      Case "kick": Scripts(CurrentScript).Hooks.Kick = True
      Case "nick": Scripts(CurrentScript).Hooks.Nick = True
      Case "commands": Scripts(CurrentScript).Hooks.Commands = True
      Case "op": Scripts(CurrentScript).Hooks.Op = True
      Case "deop": Scripts(CurrentScript).Hooks.Deop = True
      Case "mode": Scripts(CurrentScript).Hooks.Mode = True
      Case "modeend": Scripts(CurrentScript).Hooks.ModeEnd = True
      Case "topic": Scripts(CurrentScript).Hooks.Topic = True
      Case "numerics": Scripts(CurrentScript).Hooks.Numerics = True
      Case "bn": Scripts(CurrentScript).Hooks.Botnet = True
      Case "bn_msg": Scripts(CurrentScript).Hooks.BN_Msg = True
      Case "whois": Scripts(CurrentScript).Hooks.Whois = True
      Case "resolves": Scripts(CurrentScript).Hooks.Resolves = True
      Case "party_out": Scripts(CurrentScript).Hooks.Party_out = True
      Case "addeduser": Scripts(CurrentScript).Hooks.AddedUser = True
      Case "removeduser": Scripts(CurrentScript).Hooks.RemovedUser = True
      Case "addedhost": Scripts(CurrentScript).Hooks.AddedHost = True
      Case "removedhost": Scripts(CurrentScript).Hooks.RemovedHost = True
      Case "changednick": Scripts(CurrentScript).Hooks.ChangedNick = True
      Case "pljoin": Scripts(CurrentScript).Hooks.PLJoin = True
      Case "unload": Scripts(CurrentScript).Hooks.Unload = True
      Case "srv_connect": Scripts(CurrentScript).Hooks.Srv_Connect = True
      Case "ban": Scripts(CurrentScript).Hooks.Ban = True
      Case "unban": Scripts(CurrentScript).Hooks.UnBan = True

      Case "fa_uploadbegin": Scripts(CurrentScript).Hooks.fa_uploadbegin = True
      Case "fa_uploadcomplete": Scripts(CurrentScript).Hooks.fa_uploadcomplete = True
      Case "fa_downloadbegin": Scripts(CurrentScript).Hooks.fa_downloadbegin = True
      Case "fa_downloadcomplete": Scripts(CurrentScript).Hooks.fa_downloadcomplete = True
      Case "fa_command": Scripts(CurrentScript).Hooks.fa_command = True
      Case "fa_userjoin": Scripts(CurrentScript).Hooks.fa_userjoin = True
      Case "fa_userleft": Scripts(CurrentScript).Hooks.fa_userleft = True
      Case "seen": Scripts(CurrentScript).Hooks.seen = True
      Case "ki": Scripts(CurrentScript).Hooks.KI = True

      Case "raw"
        Scripts(CurrentScript).Hooks.Raw = True
        If GetRest(HookTo, 2) <> "" Then Scripts(CurrentScript).Hooks.RawFilter = Scripts(CurrentScript).Hooks.RawFilter & GetRest(HookTo, 2)

      Case Else
        Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Unknown hook target: " & (HookTo)
    End Select
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub Silent(MakeSilent As Boolean): CurrentCommand = "Silent(" & MakeSilent & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    Scripts(CurrentScript).Silent = MakeSilent
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function FileOpen(Name, Access As Byte) As Long: CurrentCommand = "FileOpen(" & Name & ", " & Access & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    Dim FNum As Long, FileName As String
    Dim FAccess As String
    FNum = FreeFile
    FileOpen = 0
    FileName = Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("\", (Name), "\")
    On Local Error Resume Next
    If (LCase(FileName) = HomeDir & "update.exe") And (AutoUpdate = True) Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - File is locked during AutoUpdate!": Exit Function
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - Access Denied!": Exit Function
    End If
    Select Case Access
      Case 1 'FA_Read
        If Dir(FileName) = "" Then Exit Function
        Open FileName For Input As #FNum
        If Err.Number = 0 Then
          FileOpen = FNum
        ElseIf Err.Number > 0 Then
          Close FNum
          FileOpen = 0
          If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - " & Err.Description
          Exit Function
        End If
      Case 2 'FA_Write
        If LCase(FileName) = LCase(AnGeL_INI) Then FileOpen = 0: Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - you must use WriteINIString to change this file.": Exit Function
        If Dir(OneDirBack(FileName), vbDirectory) = "" Then MkDir OneDirBack(FileName)
        Open FileName For Output As #FNum
        If Err.Number = 0 Then
          FileOpen = FNum
        ElseIf Err.Number > 0 Then
          Close #FNum
          FileOpen = 0
          If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - " & Err.Description
          Exit Function
        End If
        If GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini") = "" Then
          WritePPString Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), Scripts(CurrentScript).Name, HomeDir & "files.ini"
        End If
      Case 3 'FA_Append
        If LCase(FileName) = LCase(AnGeL_INI) Then FileOpen = 0: Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - you must use WriteINIString to change this file.": Exit Function
        If Dir(OneDirBack(FileName), vbDirectory) = "" Then MkDir OneDirBack(FileName)
        Open FileName For Append As #FNum
        If Err.Number = 0 Then
          FileOpen = FNum
        ElseIf Err.Number > 0 Then
          Close #FNum
          FileOpen = 0
          If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - " & Err.Description
          Exit Function
        End If
        If GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini") = "" Then
          WritePPString Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), Scripts(CurrentScript).Name, HomeDir & "files.ini"
        End If
      Case 4 'FA_Binary
        If LCase(FileName) = LCase(AnGeL_INI) Then FileOpen = 0: Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - you must use WriteINIString to change this file.": Exit Function
        If Dir(OneDirBack(FileName), vbDirectory) = "" Then MkDir OneDirBack(FileName)
        Open FileName For Binary As #FNum
        If Err.Number = 0 Then
          FileOpen = FNum
        ElseIf Err.Number > 0 Then
          Close #FNum
          FileOpen = 0
          If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't open '" & (Name) & "' - " & Err.Description
          Exit Function
        End If
        If GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini") = "" Then
          WritePPString Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), Scripts(CurrentScript).Name, HomeDir & "files.ini"
        End If
      Case Else
        FileOpen = 0
        If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Unknown file access mode, couldn't open '" & (Name) & "'."
    End Select
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub FileClose(FNum): CurrentCommand = "FileClose(" & FNum & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    If FNum = 0 Then Exit Sub
    Close #FNum
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileClose: Couldn't close file number " & CStr(FNum) & " - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function FileGet(FNum, MaxSize, Line) As Byte: CurrentCommand = "FileGet(" & FNum & ", " & MaxSize & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    Dim DataBlock As String, TheSize As Long
    On Local Error Resume Next
    If FNum = 0 Then Exit Function
    If LOF(FNum) - Loc(FNum) < CLng(MaxSize) Then
      TheSize = LOF(FNum) - Loc(FNum)
      FileGet = FG_LastData
    Else
      TheSize = CLng(MaxSize)
      FileGet = FG_Success
    End If
    DataBlock = Space(TheSize)
    Get #FNum, , DataBlock
    Line = DataBlock
    If Err.Number <> 0 Then
      FileGet = FG_Error
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileGet: Couldn't get data from file number " & CStr(FNum) & " - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub FilePut(FNum, Line): CurrentCommand = "FilePut(" & FNum & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    If FNum = 0 Then Exit Sub
    Put #FNum, , CStr(Line)
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FilePut: Couldn't write data to file number " & CStr(FNum) & " - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function FileReadln(FNum) As String: CurrentCommand = "FileReadln(" & FNum & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    If FNum = 0 Then Exit Function
    Line Input #FNum, FileReadln
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileReadln: Couldn't read from file number " & CStr(FNum) & " - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Sub FileWrite(FNum, Line): CurrentCommand = "FileWrite(" & FNum & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    If FNum = 0 Then Exit Sub
    Print #FNum, (Line);
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileWrite: Couldn't write to file number " & CStr(FNum) & " - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub FileWriteln(FNum, Line): CurrentCommand = "FileWriteln(" & FNum & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    If FNum = 0 Then Exit Sub
    Print #FNum, (Line)
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileWriteln: Couldn't write to file number " & CStr(FNum) & " - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub FileCopy(File, File2): CurrentCommand = "FileCopy(" & File & ", " & File2 & ")"
 If AnGeLFiles.CommandAllowed("FileOperations") Then
   On Local Error Resume Next
    Dim FName1 As String
    Dim FName2 As String
    Dim FAccess As String
    If InStr(LCase(File2), "angel.exe") > 0 Or InStr(LCase(File2), "update.exe") > 0 Or InStr(LCase(File2), "angel.ini") > 0 Or Dir(Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("", File, "")) = "" Then Exit Sub
    FName1 = Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("", File, "")
    FName2 = Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("", File2, "")
    FAccess = GetPPString(Replace(FName1, FileAreaHome, ""), Mid(FName1, InStrRev(FName1, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileCopy: Couldn't open '" & (File) & "' - Access Denied!": Exit Sub
    End If
    FileCopy FName1, FName2
    If Err.Number > 0 Then
      Err.Clear
    Else
      If GetPPString(Replace(FName2, FileAreaHome, ""), Mid(FName2, InStrRev(FName2, "\") + 1), "", HomeDir & "files.ini") = "" Then
        WritePPString Replace(FName2, FileAreaHome, ""), Mid(FName2, InStrRev(FName2, "\") + 1), Scripts(CurrentScript).Name, HomeDir & "files.ini"
      End If
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub FileDelete(File): CurrentCommand = "FileDelete(" & File & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    Dim FileName As String
    Dim FAccess As String
    If InStr(LCase(File), "angel.exe") > 0 Or InStr(LCase(File), "update.exe") > 0 Or InStr(LCase(File), "angel.ini") > 0 Or Dir(Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("", File, "")) = "" Then Exit Sub
    FileName = Left(HomeDir, Len(HomeDir) - 1) & GetNewDir("", File, "")
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileDelete: Couldn't open '" & (File) & "' - Access Denied!": Exit Sub
    End If
    Kill FileName
    If GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini") = "" Then
      DeletePPString Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), HomeDir & "files.ini"
    End If

    If Err.Number > 0 Then Err.Clear
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function FileEOF(FNum) As Boolean: CurrentCommand = "FileEOF(" & FNum & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    FileEOF = EOF(FNum)
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileEOF: Couldn't get EOF for file number " & CStr(FNum) & " - " & Err.Description: FileEOF = True
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function FileLen(FileName) As Long: CurrentCommand = "FileLen(" & FileName & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    Dim FAccess As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (FileName), "\")
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileLen: Couldn't open '" & (FileName) & "' - Access Denied!": Exit Function
    End If
    FileLen = FileSystem.FileLen(FileName)
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileOpen: Couldn't get Filelength '" & (FileName) & "' - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function FileLOF(FileNum) As Long: CurrentCommand = "FileLOF(" & FileNum & ""
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    FileLOF = LOF(CInt(FileNum))
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileLOF: Couldn't get Filelength '" & (FileNum) & "' - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function FileDate(FileName) As Date: CurrentCommand = "FileDate(" & FileName & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") And AnGeLFiles.CommandAllowed("TimeOperations") Then
    On Local Error Resume Next
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (FileName), "\")
    FileDate = FileDateTime(FileName)
    If Err.Number <> 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileDate: Couldn't get FileDate '" & (FileName) & "' - " & Err.Description
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function FileOwner(FileName) As String: CurrentCommand = "FileOwner(" & FileName & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (FileName), "\")
    FileOwner = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function
Public Sub SetFilePublic(FileName): CurrentCommand = "SetFilePublic(" & FileName & ")"
  If AnGeLFiles.CommandAllowed("FileOperations") Then
    On Local Error Resume Next
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (FileName), "\")
    If GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini") = Scripts(CurrentScript).Name Then
      WritePPString Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "PUBLIC", HomeDir & "files.ini"
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "*** FileLen: Couldn't open '" & (FileName) & "' - Access Denied!": Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function MyServer() As String: CurrentCommand = "MyServer()"
  On Local Error Resume Next
  MyServer = ServerName
End Function

Public Function MyHostmask() As String: CurrentCommand = "MyHostmask()"
  On Local Error Resume Next
  MyHostmask = Server_Connection.MyHostmask
End Function

Public Property Get MyIP() As String
  On Local Error Resume Next
  MyIP = Server_Connection.MyIPmask
End Property

Public Function BotNetNick() As String: CurrentCommand = "BotNetNick()"
  On Local Error Resume Next
  BotNetNick = Server_Connection.BotNetNick
End Function

Public Function MyNick() As String: CurrentCommand = "MyNick()"
  On Local Error Resume Next
  MyNick = Server_Connection.MyNick
End Function

Public Function Param(Text, Num) As String ': CurrentCommand = "Param(" & (Text) & ", " & (Num) & ")"
  On Local Error Resume Next
  If VarType(Text) >= vbArray Then Text = Join(Text, "")
  Param = Kernel_Functions.Param(CStr(Text), CLng(Num))
End Function

Public Function ParamCount(Text) As Long: CurrentCommand = "ParamCount(" & Text & ")"
  On Local Error Resume Next
  ParamCount = Kernel_Functions.ParamCount((Text))
End Function

Public Function ParamX(Text, SepChar, Num) As String: CurrentCommand = "ParamX(" & Text & ", " & SepChar & ", " & Num & ")"
  On Local Error Resume Next
  ParamX = Kernel_Functions.ParamX((Text), (SepChar), (Num))
End Function

Public Function ParamXCount(Text, SepChar) As Long: CurrentCommand = "ParamXCount(" & Text & ", " & SepChar & ")"
  On Local Error Resume Next
  ParamXCount = Kernel_Functions.ParamXCount((Text), (SepChar))
End Function

Public Function WildMatch(Wild, Full) As Boolean: CurrentCommand = "WildMatch(" & Wild & ", " & Full & ")"
  On Local Error Resume Next
  WildMatch = Userlist_Functions.SimpleMatch((Wild), (Full))
End Function

Public Function GetRest(Text, Num) As String: CurrentCommand = "GetRest(" & Text & ", " & Num & ")"
  On Local Error Resume Next
  GetRest = Kernel_Functions.GetRest((Text), (Num))
End Function

Public Function GetRestX(Text, SepChar, Num) As String: CurrentCommand = "GetRestX(" & Text & ", " & SepChar & ", " & Num & ")"
  On Local Error Resume Next
  GetRestX = Kernel_Functions.GetRestX((Text), (SepChar), (Num))
End Function

Public Sub TU(vsock, Line): CurrentCommand = "TU(" & vsock & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    CalledByScript = True
    Sockets_SocketList.TU (vsock), (Line)
    CalledByScript = False
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SpreadMessage(vsock, PLChannel, Line): CurrentCommand = "SpreadMessage(" & vsock & ", " & PLChannel & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    Partyline_Functions.SpreadMessage (vsock), (PLChannel), (Line)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SpreadFlagMessage(vsock, NeededFlags, Line): CurrentCommand = "SpreadFlagMessage(" & vsock & ", " & NeededFlags & ", " & Line & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    Partyline_Functions.SpreadFlagMessage (vsock), (NeededFlags), (Line)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SendToBot(Nick, What): CurrentCommand = "SendToBot(" & Nick & ", " & What & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    Botnet_BotTree.SendToBot (Nick), "z " & BotNetNick & " " & (Nick) & " " & (What)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SendRawToBot(Nick, What): CurrentCommand = "SendRawToBot(" & Nick & ", " & What & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    Botnet_BotTree.SendToBot (Nick), (What)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub ToBotNet(vsock, What): CurrentCommand = "ToBotNet(" & vsock & ", " & What & ")"
  If AnGeLFiles.CommandAllowed("SocketOperations") Then
    On Local Error Resume Next
    Botnet_BotTree.ToBotNet CLng(vsock), (What)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub AddKick(Chan, Nick, Reason): CurrentCommand = "AddKick(" & Chan & ", " & Nick & ", " & Reason & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    Dim ChNum As Long, UsNum As Long
    On Local Error Resume Next
    ChNum = FindChan((Chan))
    If ChNum = 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: AddKick couldn't find channel '" & (Chan) & "'."
      Exit Sub
    End If
    UsNum = FindUser((Nick), ChNum)
    If UsNum = 0 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: AddKick couldn't find user '" & (Nick) & "' on " & (Chan) & "."
      Exit Sub
    End If
    Server_Channels.AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, (Reason)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub AddCommand(Name, Class As Integer, SpreadTo, Description): CurrentCommand = "AddCommand(" & Name & ", " & Class & ", " & SpreadTo & ", " & Description & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    Partyline_CommandList.AddCommand (Name), Class, "", Scripts(CurrentScript).Name, (SpreadTo), (Description)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub RemCommand(Name): CurrentCommand = "RemCommand(" & Name & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    On Local Error Resume Next
    Partyline_CommandList.RemCommand (Name)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub SendLine(ByVal What, BufNum): CurrentCommand = "SendLine(" & What & ", " & BufNum & ")"
  If AnGeLFiles.CommandAllowed("ChannelInteractions") Then
    On Local Error Resume Next
    If CLng(BufNum) < 1 Or CLng(BufNum) > 3 Then
      If Not Scripts(CurrentScript).Silent Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: 2nd SendLine parameter has to be 1-3."
      Exit Sub
    End If
    Server_Functions.SendLine (What), CLng(BufNum)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub WriteINIString(Section, Entry, Value, Name): CurrentCommand = "WriteINIString(" & Section & ", " & Entry & ", " & Value & ", " & Name & ")"
  If AnGeLFiles.CommandAllowed("INIoperations") Then
    Dim FileName As String
    If (LCase(Section) = "others") And (LCase(Entry) = "allowruntos") Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: Scripts are not allowed to change the AllowRunToS setting.": Exit Sub
    If (LCase(Section) = "others") And (LCase(Value) = "allowruntos") Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: Scripts are not allowed to change the AllowRunToS setting.": Exit Sub
    If CheckForAllow(Section, Entry, Value) = False Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: Scripts are not allowed to change the AllowCreateObject setting.": Exit Sub
    If CheckForAllow2(Section, Entry, Value) = False Then Partyline_Functions.SpreadFlagMessage 0, "+m", "*** Scripting Error: Scripts are not allowed to change the AllowRunToS setting.": Exit Sub
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FileSys_Functions.WritePPString (Section), (Entry), (Value), FileName
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub DeleteINISection(Section, Name): CurrentCommand = "DeleteINISection(" & Section & ", " & Name & ")"
  If AnGeLFiles.CommandAllowed("INIoperations") Then
    Dim FileName As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FileSys_Functions.DeletePPString (Section), "", FileName
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub DeleteINIString(Section, Entry, Name): CurrentCommand = "DeleteINIString(" & Section & ", " & Entry & ", " & Name & ")"
  If AnGeLFiles.CommandAllowed("INIoperations") Then
    Dim FileName As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FileSys_Functions.DeletePPString (Section), (Entry), FileName
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function ReadINIString(Section, Entry, Default, Name) As String: CurrentCommand = "ReadINIString(" & Section & ", " & Entry & ", " & Default & ", " & Name & ")"
  If AnGeLFiles.CommandAllowed("INIoperations") Then
    Dim FileName As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    ReadINIString = FileSys_Functions.GetPPString((Section), (Entry), (Default), FileName)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function
Public Function seen(seentext, MatchedOne As Boolean): CurrentCommand = "seen(" & seentext & ", " & MatchedOne & ")"
  If AnGeLFiles.CommandAllowed("BotInteractions") Then
    seen = Plugin_Seen.LastSeen(CStr(seentext), "", "", "", MatchedOne)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Function

Public Function WindowsVersion() As String
  WindowsVersion = WinVersionName
End Function

Public Sub NTFSCompress(Name): CurrentCommand = "NTFSCompress(" & Name & ")"
  If AnGeLFiles.CommandAllowed("NativeFS") = True Then
    Dim FileName As String
    Dim ReturnVal As Long
    Dim FAccess As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSCompress: Couldn't open '" & (Name) & "' - Access Denied!": Exit Sub
    End If
    ReturnVal = FileSys_NTFS.NTFS_SetCompression(True, (CStr(FileName)))
    If ReturnVal <> 0 Then
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSCompress: " & GetErrMSG(ReturnVal): Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub NTFSDeCompress(Name): CurrentCommand = "NTFSDeCompress(" & Name & ")"
  If AnGeLFiles.CommandAllowed("NativeFS") = True Then
    Dim FileName As String
    Dim ReturnVal As Long
    Dim FAccess As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSDeCompress: Couldn't open '" & (Name) & "' - Access Denied!": Exit Sub
    End If
    ReturnVal = FileSys_NTFS.NTFS_SetCompression(True, (CStr(FileName)))
    If ReturnVal <> 0 Then
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSDeCompress: " & GetErrMSG(ReturnVal): Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub NTFSEncrypt(Name): CurrentCommand = "NTFSEncrypt(" & Name & ")"
  If AnGeLFiles.CommandAllowed("NativeFS") = True Then
    Dim FileName As String
    Dim ReturnVal As Long
    Dim FAccess As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSEncrypt: Couldn't open '" & (Name) & "' - Access Denied!": Exit Sub
    End If
    ReturnVal = FileSys_NTFS.NTFS_Encrypt(CStr(FileName))
    If ReturnVal <> 0 Then
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSEncrypt: " & GetErrMSG(ReturnVal): Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Sub NTFSDecrypt(Name): CurrentCommand = "NTFSDecrypt(" & Name & ")"
  If AnGeLFiles.CommandAllowed("NativeFS") = True Then
    Dim FileName As String
    Dim ReturnVal As Long
    Dim FAccess As String
    FileName = Left(HomeDir, Len(HomeDir) - 1) + GetNewDir("\", (Name), "\")
    FAccess = GetPPString(Replace(FileName, FileAreaHome, ""), Mid(FileName, InStrRev(FileName, "\") + 1), "", HomeDir & "files.ini")
    If FAccess = "" Or FAccess = "PUBLIC" Or FAccess = Scripts(CurrentScript).Name Then
      FAccess = FAccess
    Else
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSDecrypt: Couldn't open '" & (Name) & "' - Access Denied!": Exit Sub
    End If
    If InStr(LCase(Name), "policies.dat") > 0 Then
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSDecrypt: Scripts are NOT allowd to decrypt the policies.dat": Exit Sub
      Exit Sub
    ElseIf InStr(LCase(Name), "files.ini") > 0 Then
      Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSDecrypt: Scripts are NOT allowd to decrypt the files.ini": Exit Sub
      Exit Sub
    End If
    ReturnVal = FileSys_NTFS.NTFS_Decrypt(CStr(FileName))
    If ReturnVal <> 0 Then
     Partyline_Functions.SpreadFlagMessage 0, "+m", "5*** NTFSDecrypt: " & GetErrMSG(ReturnVal): Exit Sub
    End If
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Sub MakeSessionValid(UserNum): CurrentCommand = "MakeSessionValid(" & UserNum & ")"
  If AnGeLFiles.CommandAllowed("SessionChange") = True Then
    MakeSessionValid CLng(UserNum)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Sub MakeSessionInValid(UserNum): CurrentCommand = "MakeSessionInValid(" & UserNum & ")"
  If AnGeLFiles.CommandAllowed("SessionChange") = True Then
    MakeSessionInValid CLng(UserNum)
  Else
    PolError CurrentScript, CurrentCommand
  End If
End Sub

Public Function ValidSock(vsock) As Boolean: On Local Error Resume Next: ValidSock = IsValidSocket((vsock)): End Function

Public Function VSockCount() As Long: On Local Error Resume Next: VSockCount = SocketCount: End Function

Public Function MSS(TheString) As String: On Local Error Resume Next: MSS = Scripting_Scripts.MSS((TheString)): End Function

Public Function ChanName(Nr) As String: On Local Error Resume Next: ChanName = Server_Channels.Channels((Nr)).Name: End Function

Public Function ChanTopic(Nr) As String: On Local Error Resume Next: ChanTopic = Server_Channels.Channels((Nr)).Topic: End Function

Public Function ChanMode(Nr) As String: On Local Error Resume Next: ChanMode = Server_Channels.Channels((Nr)).Mode: End Function

Public Function ChanLimit(Nr) As Long: On Local Error Resume Next: ChanLimit = Server_Channels.GetChannelLimit((Nr)): End Function

Public Function ChanCount() As Long: On Local Error Resume Next: ChanCount = Server_Channels.ChanCount: End Function

Public Function ChanUserCount(Nr) As Long: On Local Error Resume Next: ChanUserCount = Server_Channels.Channels((Nr)).UserCount: End Function

Public Function ChanBanCount(Nr) As Long: On Local Error Resume Next: ChanBanCount = Server_Channels.Channels((Nr)).BanCount: End Function

Public Function BotUserCount() As Long: On Local Error Resume Next: BotUserCount = Userlist_BotUserlist.BotUserCount: End Function

Public Function BotCount() As Long: On Local Error Resume Next: BotCount = Botnet_BotTree.BotCount: End Function

Public Function BotNick(Nr) As String: On Local Error Resume Next: BotNick = Bots((Nr)).Nick: End Function

Public Function BotSubBotOf(Nr) As String: On Local Error Resume Next: BotSubBotOf = Bots((Nr)).SubBotOf: End Function

Public Function BotVersion(Nr) As String: On Local Error Resume Next: BotVersion = Bots((Nr)).Version: End Function

'-- CONSTANTS ---------------------------------- (Yes, I know, lame implementation....)

Public Function FA_Read() As Byte: FA_Read = 1: End Function
Public Function FA_Write() As Byte: FA_Write = 2: End Function
Public Function FA_Append() As Byte: FA_Append = 3: End Function
Public Function FA_Binary() As Byte: FA_Binary = 4: End Function

Public Function FG_Success() As Byte: FG_Success = 1: End Function
Public Function FG_LastData() As Byte: FG_LastData = 2: End Function
Public Function FG_Error() As Byte: FG_Error = 3: End Function

Public Function Cl_User() As Integer: Cl_User = 1: End Function
Public Function Cl_What() As Integer: Cl_What = 2: End Function
Public Function Cl_Op() As Integer: Cl_Op = 4: End Function
Public Function Cl_Net() As Integer: Cl_Net = 8: End Function
Public Function Cl_CMas() As Integer: Cl_CMas = 16: End Function
Public Function Cl_Mas() As Integer: Cl_Mas = 32: End Function
Public Function Cl_COwn() As Integer: Cl_COwn = 64: End Function
Public Function Cl_Own() As Integer: Cl_Own = 128: End Function
Public Function Cl_SOwn() As Integer: Cl_SOwn = 256: End Function

'VSockInfo function
Public Function VS_RegNick() As Byte: VS_RegNick = 1: End Function
Public Function VS_Flags() As Byte: VS_Flags = 2: End Function
Public Function VS_IRCNick() As Byte: VS_IRCNick = 3: End Function
Public Function VS_Hostmask() As Byte: VS_Hostmask = 4: End Function
Public Function VS_UserNum() As Byte: VS_UserNum = 5: End Function
Public Function VS_AwayMessage() As Byte: VS_AwayMessage = 6: End Function
Public Function VS_IdleTime() As Byte: VS_IdleTime = 7: End Function

'ChanUser function
Public Function CU_Nick() As Byte: CU_Nick = 1: End Function
Public Function CU_RegUser() As Byte: CU_RegUser = 2: End Function
Public Function CU_Hostmask() As Byte: CU_Hostmask = 3: End Function
Public Function CU_IPMask() As Byte: CU_IPMask = 4: End Function
Public Function CU_Status() As Byte: CU_Status = 5: End Function
Public Function CU_IdleTime() As Byte: CU_IdleTime = 6: End Function

'BotUser function
Public Function BU_RegNick() As Byte: BU_RegNick = 1: End Function
Public Function BU_Flags() As Byte: BU_Flags = 2: End Function
Public Function BU_Password() As Byte: BU_Password = 3: End Function
Public Function BU_ChanFlags() As Byte: BU_ChanFlags = 4: End Function
Public Function BU_Hostmasks() As Byte: BU_Hostmasks = 5: End Function
Public Function BU_ValidSession() As Byte: BU_ValidSession = 6: End Function

'Socket events
Public Function SE_ConnectFailed() As Byte: SE_ConnectFailed = 0: End Function
Public Function SE_Connected() As Byte: SE_Connected = 1: End Function
Public Function SE_Read() As Byte: SE_Read = 2: End Function
Public Function SE_Closed() As Byte: SE_Closed = 3: End Function
Public Function SE_IncomingFailed() As Byte: SE_IncomingFailed = 4: End Function
Public Function SE_Incoming() As Byte: SE_Incoming = 5: End Function

'AddHost replies
Public Function AH_Success() As Byte: AH_Success = 0: End Function
Public Function AH_UserNotFound() As Byte: AH_UserNotFound = 1: End Function
Public Function AH_InvalidHost() As Byte: AH_InvalidHost = 2: End Function
Public Function AH_AlreadyThere() As Byte: AH_AlreadyThere = 3: End Function
Public Function AH_MatchingUser() As Byte: AH_MatchingUser = 4: End Function
Public Function AH_TooManyHosts() As Byte: AH_TooManyHosts = 5: End Function
Public Function AH_DENIED() As Byte: AH_DENIED = 6: End Function

'RemHost replies
Public Function RH_Success() As Byte: RH_Success = 0: End Function
Public Function RH_UserNotFound() As Byte: RH_UserNotFound = 1: End Function
Public Function RH_HostNotFound() As Byte: RH_HostNotFound = 2: End Function
Public Function RH_DENIED() As Byte: RH_DENIED = 3: End Function

'AddUser replies
Public Function AU_Success() As Byte: AU_Success = 0: End Function
Public Function AU_TooLong() As Byte: AU_TooLong = 1: End Function
Public Function AU_UserExists() As Byte: AU_UserExists = 2: End Function
Public Function AU_InvalidNick() As Byte: AU_InvalidNick = 3: End Function
Public Function AU_DENIED() As Byte: AU_DENIED = 4: End Function

'RemUser replies
Public Function RU_Success() As Byte: RU_Success = 0: End Function
Public Function RU_UserNotFound() As Byte: RU_UserNotFound = 1: End Function
Public Function RU_DENIED() As Byte: RU_DENIED = 2: End Function

'Chattr replies
Public Function CH_Success() As Byte: CH_Success = 0: End Function
Public Function CH_NoChanges() As Byte: CH_NoChanges = 1: End Function
Public Function CH_NoChanFlag() As Byte: CH_NoChanFlag = 2: End Function
Public Function CH_DENIED() As Byte: CH_DENIED = 3: End Function

