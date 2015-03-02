Attribute VB_Name = "Sockets_SocketEvents"
',-======================- ==-- -  -
'|   AnGeL - Sockets - Ereignisse
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Sub Socket_Closed(ByVal TheSock As Long, ByVal ErrorCode As Long, ByVal Reason As String)
  Dim u As Long
  If GetSockFlag(TheSock, SF_LocalVisibleUser) = SF_YES Then
    ToBotNet 0, "pt " & BotNetNick & " " & SocketItem(TheSock).RegNick & " " & SocketItem(TheSock).OrderSign
    SpreadMessageEx TheSock, SocketItem(TheSock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLLeave, SocketItem(TheSock).RegNick)
    WriteSeenEntry SocketItem(TheSock).RegNick, "", Now, "*mine*", "*partyline*", Mask(SocketItem(TheSock).Hostmask, 10)
  Else
    'No visible bot user - handle other vsock types
    Select Case GetSockFlag(TheSock, SF_Status)
      Case SF_Status_Server
        ServerSocket = -1
        GUI_frmWinsock.cmdConnect.Caption = "connect"
        Output vbCrLf
        Output "*** Socket Closed (Server)" & vbCrLf
        If (Connected = True) Or (ProxyToConnect <> "") Then
          If (ProxyToConnect <> "") And (Connected = False) Then
            SpreadFlagMessage 0, "+m", MakeMsg(ERR_ServerFailed, GetAddr(ServerToConnect) & " (via " & GetAddr(ProxyToConnect) & ")", "Proxy connection closed." & IIf(Reason = "", "", " " & Reason))
            PutLog "| Proxy connection closed"
          Else
            Status "*** Connection to server lost" & vbCrLf
            SpreadFlagMessage 0, "+m", MakeMsg(ERR_ServerLost, IIf(Reason = "", "", Reason))
          End If
        Else
          SpreadFlagMessage 0, "+m", MakeMsg(ERR_ServerFailed, GetAddr(ServerToConnect), Reason)
        End If
        'Reset all variables
        Disconnect
        'connect to next server
        AutoConnect
      
      Case SF_Status_RelaySrv
        DisconnectSocket CLng(SocketItem(TheSock).CurrentQuestion)
      
      Case SF_Status_RelayCli
        TU CLng(SocketItem(TheSock).CurrentQuestion), "5*** Disconnected from " & SocketItem(TheSock).SetupChan & "."
        SetSockFlag CLng(SocketItem(TheSock).CurrentQuestion), SF_Status, SF_Status_Party
        SetSockFlag CLng(SocketItem(TheSock).CurrentQuestion), SF_LocalVisibleUser, SF_YES
      
      Case SF_Status_ScriptSocket
        For u = 1 To ScriptCount
          If Scripts(u).Name = SocketItem(TheSock).SetupChan Then
            RunScriptX u, SocketItem(TheSock).CurrentQuestion, TheSock, 3, "Connection closed."
          End If
        Next u
      
      Case SF_Status_File
        If SocketItem(TheSock).FileSize = LOF(SocketItem(TheSock).FileNum) Then
          Close #SocketItem(TheSock).FileNum
          If Dir(HomeDir & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(HomeDir & SocketItem(TheSock).FileName) <> "" Then Kill HomeDir & SocketItem(TheSock).FileName
            Name HomeDir & SocketItem(TheSock).FileName & ".temp" As HomeDir & SocketItem(TheSock).FileName
          ElseIf Dir(FileAreaHome & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(FileAreaHome & SocketItem(TheSock).FileName) <> "" Then Kill FileAreaHome & SocketItem(TheSock).FileName
            Name FileAreaHome & SocketItem(TheSock).FileName & ".temp" As FileAreaHome & SocketItem(TheSock).FileName
          ElseIf Dir(FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName) <> "" Then Kill FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName
            Name FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName & ".temp" As FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName
          ElseIf Dir(FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName) <> "" Then Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
            Name FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName & ".temp" As FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
          End If
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC Get from " & SocketItem(TheSock).RegNick & " completed (" & GetFileName(SocketItem(TheSock).FileName) & ")"
          For u = 1 To ScriptCount
            If Scripts(u).Hooks.fa_uploadcomplete Then
              RunScriptX u, "fa_uploadcomplete", SocketItem(TheSock).IRCNick, SocketItem(TheSock).RegNick, SocketItem(TheSock).FileName, SocketItem(TheSock).FileSize
            End If
          Next u
          AfterFileCompletion TheSock
        Else
          Close #SocketItem(TheSock).FileNum
          If Dir(HomeDir & SocketItem(TheSock).FileName & ".temp") <> "" Then
            Kill HomeDir & SocketItem(TheSock).FileName & ".temp"
          ElseIf Dir(FileAreaHome & SocketItem(TheSock).FileName & ".temp") <> "" Then
            Kill FileAreaHome & SocketItem(TheSock).FileName & ".temp"
          ElseIf Dir(FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName & ".temp") <> "" Then
            Kill FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName & ".temp"
          ElseIf Dir(FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName & ".temp") <> "" Then
            Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName & ".temp"
          End If
          SpreadFlagMessage TheSock, "+m", "14[" & Time & "] *** DCC Get from " & SocketItem(TheSock).RegNick & " incomplete!"
        End If
      
      Case SF_Status_SendFile
        If SocketItem(TheSock).BytesReceived = SocketItem(TheSock).FileSize Then
          Close #SocketItem(TheSock).FileNum
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " completed (" & GetFileName(SocketItem(TheSock).FileName) & ")"
            For u = 1 To ScriptCount
              If Scripts(u).Hooks.fa_downloadcomplete Then
                RunScriptX u, "fa_downloadcomplete", SocketItem(TheSock).RegNick, GetFileName(SocketItem(TheSock).FileName)
              End If
            Next u
          SendNextQueuedFile SocketItem(TheSock).RegNick
        Else
          Close #SocketItem(TheSock).FileNum
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " aborted!  (" & GetFileName(SocketItem(TheSock).FileName) & ")"
          SendNextQueuedFile SocketItem(TheSock).RegNick
        End If
      
      Case SF_Status_BotGetName, SF_Status_BotGetPass, SF_Status_BotLinking, SF_Status_BotPreCache
        SpreadFlagMessageEx TheSock, "+t", SF_Local_Bot, MakeMsg(MSG_PLBotNetLost, IIf(SocketItem(TheSock).RegNick <> "", SocketItem(TheSock).RegNick, SocketItem(TheSock).Hostmask))
      
      Case SF_Status_UserGetName, SF_Status_UserGetPass
        If GetSockFlag(TheSock, SF_DCC) = SF_NO Then
          SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLTelNetClosed, SocketItem(TheSock).Hostmask)
        Else
          SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCClosed, SocketItem(TheSock).RegNick)
        End If
      
      Case SF_Status_InitBotLink
        SpreadFlagMessageEx TheSock, "+t", SF_Local_Bot, MakeMsg(MSG_PLBotNetNoLink, SocketItem(TheSock).RegNick, "Connection closed.")
        'Couldn't connect hub bot - link to alternate bots
        If MatchFlags(BotUsers(SocketItem(TheSock).UserNum).BotFlags, "+h") Then
          ConnectAlternates
        End If
      
      Case SF_Status_Bot
        'Reconnect to hub bot
        If Exitting = False Then
          If MatchFlags(BotUsers(SocketItem(TheSock).UserNum).BotFlags, "+h") Then
            PutLog "Reconnecting to hub bot at " & GetUserData(SocketItem(TheSock).UserNum, UD_LinkAddr, "")
            InitiateBotChat SocketItem(TheSock).UserNum, True
          End If
        End If
      
      Case SF_Status_Ident
      
      Case Else
        Trace "hihi", SocketItem(TheSock).SocketFlag(SF_Status)
        If GetSockFlag(TheSock, SF_DCC) = SF_NO Then
          SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLTelNetClosed, SocketItem(TheSock).RegNick)
        Else
          SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCClosed, SocketItem(TheSock).RegNick)
        End If
        
    End Select
  End If
End Sub

Public Sub Socket_ConnectError(ByVal TheSock As Long, ByVal ErrorCode As Long, ByVal Reason As String)
  Dim u As Long
  Select Case GetSockFlag(TheSock, SF_Status)
    Case SF_Status_Server
      If ProxyToConnect = "" Then
        SpreadFlagMessage 0, "+m", MakeMsg(ERR_ServerFailed, GetAddr(ServerToConnect), Reason)
      Else
        SpreadFlagMessage 0, "+m", MakeMsg(ERR_ServerFailed, GetAddr(ServerToConnect) & " (via " & GetAddr(ProxyToConnect) & ")", Reason)
      End If
      'Reset all variables
      Disconnect
      'connect to next server
      AutoConnect
      
    Case SF_Status_RelayCli
      TU CLng(SocketItem(TheSock).CurrentQuestion), "5*** Couldn't connect to " & SocketItem(TheSock).SetupChan & "."
      SetSockFlag CLng(SocketItem(TheSock).CurrentQuestion), SF_Status, SF_Status_Party
      SetSockFlag CLng(SocketItem(TheSock).CurrentQuestion), SF_LocalVisibleUser, SF_YES
      RemoveSocket TheSock, ErrorCode, Reason, True
      
    Case SF_Status_ScriptSocket
      For u = 1 To ScriptCount
        If Scripts(u).Name = SocketItem(TheSock).SetupChan Then
          RunScriptX u, SocketItem(TheSock).CurrentQuestion, TheSock, 0, "Couldn't connect to host."
        End If
      Next u
      RemoveSocket TheSock, ErrorCode, Reason, True
      
    Case SF_Status_FileWaiting
      SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC Get from " & SocketItem(TheSock).RegNick & " aborted - Couldn't connect!"
      Close #SocketItem(TheSock).FileNum
      RemoveSocket TheSock, ErrorCode, Reason, True
      
    Case SF_Status_InitBotLink
      If SocketItem(TheSock).LinkStatus <> "AutoLink" Then
        SpreadFlagMessageEx TheSock, "+t", SF_Local_Bot, MakeMsg(MSG_PLBotNetNoLink, SocketItem(TheSock).RegNick, Reason)
      End If
      'Couldn't connect hub bot - link to alternate bots
      If MatchFlags(BotUsers(SocketItem(TheSock).UserNum).BotFlags, "+h") Then
        ConnectAlternates
      End If
      RemoveSocket TheSock, ErrorCode, Reason, True
      
    Case Else
      SpreadFlagMessageEx 0, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCFailed, SocketItem(TheSock).RegNick)
      RemoveSocket TheSock, ErrorCode, Reason, True
    
  End Select
End Sub

Public Sub Socket_Connected(ByVal TheSock As Long)
  Dim TempIdent As String, Identd As String, Data As String, u As Long, u2 As Long
  Select Case GetSockFlag(TheSock, SF_Status)
    Case SF_Status_Server
      GUI_frmWinsock.cmdConnect.Caption = "Disconnect"
      TempIdent = GetPPString("Identification", "admin", "", AnGeL_INI)
      If TempIdent = "" Then TempIdent = "angel"
      Identd = Param(GetSetting("NT-SHELL", "identd", TempIdent, GetPPString("Identification", "Identd", "", AnGeL_INI)), 1)
      If Identd = "" Then MsgBox "Die Datei AnGeL.ini in meinem Verzeichnis muﬂ einen Eintrag ""Identd=..."" im Abschnitt ""[Identification]"" besitzen!", vbCritical, "Konfigurations-Fehler!": Unload GUI_frmWinsock: Exit Sub
      RealName = GetPPString("Identification", "RealName", "Frag doch! :)", AnGeL_INI)
      BlockSends = False: BytesSent = 0
      Data = GetPPString("Identification", "LocalHost", "", AnGeL_INI)
      If Data = "" Then Data = WSAGetLocalHostName
      If ProxyToConnect = "" Then
        If ParamX(Param(ServerToConnect, 1), ":", 3) <> "" Then SendIt "PASS " & ParamX(Param(ServerToConnect, 1), ":", 3) + vbCrLf: SpreadFlagMessage 0, "+m", "3*** Sent server password / login"
        If InStr(1, Data, ":", vbBinaryCompare) Then
          Data = "localhost"
        End If
        SendIt "USER " & Identd & " " & Replace(Data, " ", "_") & " " & GetCacheIP(Data, True) & " :" & RealName + vbCrLf
        SendIt "NICK " & PrimaryNick & vbCrLf
      Else
        SpreadFlagMessage 0, "+m", "3*** Got connection to proxy - forwarding to " & ServerToConnect & "..."
        Select Case GetPort(ProxyToConnect, 1080)
          Case 80, 3128, 8080
            SendIt "connect " & ServerToConnect & " HTTP/1.0" & vbCrLf + vbCrLf
          Case Else
            SendIt "" & Mid(MakeHEXNum(GetPort(ServerToConnect, 6667)), 3, 2) & MakeString(GetCacheIP(GetAddr(ServerToConnect), True)) + ParamX(ProxyToConnect, ":", 3) + Chr(0)
        End Select
      End If
      GUI_frmWinsock.ConnectTimeOut.Enabled = False
      ConnectTime = Now
    
    Case SF_Status_RelayCli
      If CLng(SocketItem(CLng(SocketItem(TheSock).CurrentQuestion)).CurrentQuestion) = TheSock Then
        TU CLng(SocketItem(TheSock).CurrentQuestion), "3*** Connected to " & SocketItem(TheSock).SetupChan & "."
      End If
      SocketItem(TheSock).RegNick = "<RELAY>"
      SocketItem(TheSock).AwayMessage = ""
    
    Case SF_Status_ScriptSocket
      For u = 1 To ScriptCount
        If Scripts(u).Name = SocketItem(TheSock).SetupChan Then
          RunScriptX u, SocketItem(TheSock).CurrentQuestion, TheSock, 1, "Connection established."
        End If
      Next u
    
    Case SF_Status_FileWaiting
      Data = IIf(LCase(Right(SocketItem(TheSock).FileName, 4)) <> ".asc", FileAreaHome & "Incoming", FileAreaHome & "Scripts")
      If Not DirExist(FileAreaHome) Then MkDir FileAreaHome
      If Not DirExist(Data) Then MkDir Data
      SetSockFlag TheSock, SF_Status, SF_Status_File
      SocketItem(TheSock).FileNum = FreeFile: Open Data & "\" & SocketItem(TheSock).FileName & ".temp" For Binary As #SocketItem(TheSock).FileNum
      Data = IIf(LCase(Right(SocketItem(TheSock).FileName, 9)) <> ".asc.temp", "\Incoming", "\Scripts")
      WritePPString Data, SocketItem(TheSock).FileName, SocketItem(TheSock).RegNick, HomeDir & "Files.ini"
      If Err.Number > 0 Then
        Err.Clear
        SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC Get from " & SocketItem(TheSock).RegNick & " aborted - Error!"
        RemoveSocket TheSock, 0, "", True
      End If
    
    Case SF_Status_InitBotLink
      SetSockFlag TheSock, SF_Status, SF_Status_BotLinking
      If SocketItem(TheSock).LinkStatus = "AutoLink" Then
        If IsBotConnected(SocketItem(TheSock).RegNick) Then
          u2 = GetNextBot(SocketItem(TheSock).RegNick)
          If u2 > 1 Then
            If MatchFlags(BotUsers(GetUserNum(Bots(u2).Nick)).BotFlags, "-h") Then
              u = Bots(GetBotPos(Bots(u2).Nick)).vSocket
              If u > 0 Then
                RTU u, "bye Restructure"
                RemoveSocket u, 0, MakeMsg(MSG_BNRestructure), True
              End If
            End If
          End If
        End If
      End If
      SocketItem(TheSock).LinkStatus = ""
    
    Case Else
      If GetSockFlag(TheSock, SF_Status) = SF_Status_DCCWaiting Then
        For u = 1 To SocketCount
          If IsValidSocket(u) And u <> TheSock Then
            If SocketItem(u).RegNick = SocketItem(TheSock).RegNick Then
              If GetSockFlag(u, SF_Status) = SF_Status_DCCWaiting Then RemoveSocket u, 0, "", True
            End If
          End If
        Next u
      End If
      If BotUsers(SocketItem(TheSock).UserNum).Password <> "" Then
        SetSockFlag TheSock, SF_Status, SF_Status_UserGetPass
        SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCOpened, SocketItem(TheSock).RegNick)
        TU TheSock, MakeMsg(MSG_EnterPWD, SocketItem(TheSock).RegNick)
      Else
        SetSockFlag TheSock, SF_Status, SF_Status_UserChoosePass
        SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCFirst, SocketItem(TheSock).RegNick)
        TU TheSock, MakeMsg(MSG_ChoosePWD, SocketItem(TheSock).RegNick)
      End If
  End Select
End Sub

Public Sub Socket_DataTCP(ByVal TheSock As Long, ByVal Data As String)
  Dim SNum As Long, TempIdent As String, Identd As String, u As Long
  AddBytesIn Len(Data)
  If GetSockFlag(TheSock, SF_LocalVisibleUser) = SF_YES Then
    SNum = SocketItem(TheSock).SocketNumber
    PartySort TheSock, SNum, Data
  Else
    'Nope. Handle other vsock types.
    Select Case GetSockFlag(TheSock, SF_Status)
      Case SF_Status_Server
        If (ProxyToConnect <> "") And (SentLogin = False) Then
          SpreadFlagMessage 0, "+m", "3*** Sent login string"
          TempIdent = GetPPString("Identification", "admin", "", AnGeL_INI)
          TempIdent = Mid(TempIdent, 1, InStr(TempIdent, " ") - 1): If TempIdent = "" Then TempIdent = "angel"
          Identd = GetSetting("NT-SHELL", "identd", TempIdent, GetPPString("Identification", "Identd", "", AnGeL_INI))
          RealName = GetPPString("Identification", "RealName", "Frag doch! :)", AnGeL_INI)
          If IsValidIP(GetAddr(ProxyToConnect)) Then
            SendIt "USER " & Identd & " " & GetAddr(ProxyToConnect) & " " & GetCacheHost(GetAddr(ProxyToConnect), True) & " :" & RealName & vbCrLf
          Else
            SendIt "USER " & Identd & " " & GetCacheIP(GetAddr(ProxyToConnect), True) & " " & GetAddr(ProxyToConnect) & " :" & RealName & vbCrLf
          End If
          SendIt "NICK " & PrimaryNick + vbCrLf
          SentLogin = True
        Else
          Handle Data
        End If
      
      Case SF_Status_RelaySrv
        SendTCP CLng(SocketItem(TheSock).CurrentQuestion), Data
          
      Case SF_Status_RelayCli
        If SocketItem(TheSock).AwayMessage = "" And InStr(1, Data, "  |   AnGeL Telnet Server V", vbBinaryCompare) Then
          u = InStr(1, Data, "Server V", vbBinaryCompare)
          If CByte(Mid(Data, u + 8, 1)) = 1 And CByte(Mid(Data, u + 10, 1)) >= 2 Then
            TU TheSock, "100 relay"
            SocketItem(TheSock).AwayMessage = "done"
          ElseIf CByte(Mid(Data, u + 8, 1)) >= 2 Then
            TU TheSock, "100 relay"
            SocketItem(TheSock).AwayMessage = "done"
          End If
        ElseIf SocketItem(TheSock).AwayMessage = "done" And InStr(1, Data, "101 ", vbBinaryCompare) <> 0 Then
          u = InStr(1, Data, "101 ", vbBinaryCompare)
          Data = Left(Data, 0) & Mid(Data, InStr(u, Data, vbLf, vbBinaryCompare) + 1)
          SocketItem(TheSock).AwayMessage = "relay"
        End If
        SendTCP CLng(SocketItem(TheSock).CurrentQuestion), IIf(GetSockFlag(CLng(SocketItem(TheSock).CurrentQuestion), SF_DCC) = SF_YES, MakeDCCColor(Data, TheSock), Data)
      
      Case SF_Status_ScriptSocket
        For u = 1 To ScriptCount
          If u > UBound(Scripts) Then Exit For
          If Scripts(u).Name = SocketItem(TheSock).SetupChan Then
            RunScriptX u, SocketItem(TheSock).CurrentQuestion, TheSock, 2, Data
          End If
        Next u
      
      Case SF_Status_Ident
        If Param(Left(Data, Len(Data) - 2), 1) = Trim(Str(ServerPort)) Then
          TempIdent = GetPPString("Identification", "admin", "", AnGeL_INI)
          TempIdent = Mid(TempIdent, 1, InStr(TempIdent, " ") - 1): If TempIdent = "" Then TempIdent = "angel"
          Identd = GetSetting("NT-SHELL", "identd", TempIdent, GetPPString("Identification", "Identd", "", AnGeL_INI))
          PutLog "|  *** Ident replied: " & Left(Data, Len(Data) - 2) & " : USERID : " & UCase(WinIdentifier) & " : " & Identd
          Output "*** Ident replied: " & Left(Data, Len(Data) - 2) & " : USERID : " & UCase(WinIdentifier) & " : " & Identd + vbCrLf
          SS TheSock, Left(Data, Len(Data) - 2) & " : USERID : " & UCase(WinIdentifier) & " : " & Identd
        Else
          PutLog "|  *** No ident replied - wrong local port: " & Param(Left(Data, Len(Data) - 2), 1) & " (mine: " & Trim(Str(ServerPort)) & ")"
          Output "*** No ident replied - wrong local port: " & Param(Left(Data, Len(Data) - 2), 1) & " (mine: " & Trim(Str(ServerPort)) & ")" & vbCrLf
          SS TheSock, Left(Data, Len(Data) - 2) & " : USERID : OTHER : " & Identd
        End If
        RemoveSocket TheSock, 0, "", True
      
      Case SF_Status_File
        On Local Error Resume Next
        Put #SocketItem(TheSock).FileNum, , Data
        SocketItem(TheSock).BytesReceived = SocketItem(TheSock).BytesReceived + Len(Data)
        SendTCP TheSock, MakeHEXNum(SocketItem(TheSock).BytesReceived)
        If SocketItem(TheSock).FileSize = LOF(SocketItem(TheSock).FileNum) Then
          If Dir(HomeDir & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(HomeDir & SocketItem(TheSock).FileName) <> "" Then Kill HomeDir & SocketItem(TheSock).FileName
            Name HomeDir & SocketItem(TheSock).FileName & ".temp" As HomeDir & SocketItem(TheSock).FileName
          ElseIf Dir(FileAreaHome & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(FileAreaHome & SocketItem(TheSock).FileName) <> "" Then Kill FileAreaHome & SocketItem(TheSock).FileName
            Name FileAreaHome & SocketItem(TheSock).FileName & ".temp" As FileAreaHome & SocketItem(TheSock).FileName
          ElseIf Dir(FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName) <> "" Then Kill FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName
            Name FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName & ".temp" As FileAreaHome & "Scripts\" & SocketItem(TheSock).FileName
          ElseIf Dir(FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName & ".temp") <> "" Then
            If Dir(FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName) <> "" Then Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
            Name FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName & ".temp" As FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
          End If
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC Get from " & SocketItem(TheSock).RegNick & " completed (" & GetFileName(SocketItem(TheSock).FileName) & ")"
          For u = 1 To ScriptCount
            If Scripts(u).Hooks.fa_uploadcomplete Then
              RunScriptX u, "fa_uploadcomplete", SocketItem(TheSock).IRCNick, SocketItem(TheSock).RegNick, Left(SocketItem(TheSock).FileName, Len(SocketItem(TheSock).FileName) - 5), SocketItem(TheSock).FileSize
            End If
          Next u
          AfterFileCompletion TheSock
          DisconnectSocket TheSock
        End If
        If Err Then Err.Clear
      
      Case SF_Status_SendFile
        On Local Error Resume Next
        u = MakeLongNum(Data)
        If u > SocketItem(TheSock).BytesReceived Then
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " aborted!  (sent < received: " & GetFileName(SocketItem(TheSock).FileName) & ")"
          SendNextQueuedFile SocketItem(TheSock).RegNick
          Close #SocketItem(TheSock).FileNum
          RemoveSocket TheSock, 0, "", True
        Else
          If u = SocketItem(TheSock).BytesReceived Then
            If SocketItem(TheSock).BytesReceived < SocketItem(TheSock).FileSize Then
              u = PumpDCC: If SocketItem(TheSock).BytesReceived + u > SocketItem(TheSock).FileSize Then u = SocketItem(TheSock).FileSize - SocketItem(TheSock).BytesReceived
              Data = Space(u): Get SocketItem(TheSock).FileNum, , Data
              If Err.Number > 0 Then
                SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " aborted!  (" & Err.Description & ": " & GetFileName(SocketItem(TheSock).FileName) & ")"
                Err.Clear
                SendNextQueuedFile SocketItem(TheSock).RegNick
                Close #SocketItem(TheSock).FileNum
                RemoveSocket TheSock, 0, "", True
              End If
              SocketItem(TheSock).BytesReceived = SocketItem(TheSock).BytesReceived + u
              SendTCP TheSock, Data
            Else
              If SocketItem(TheSock).BytesReceived = SocketItem(TheSock).FileSize Then
                SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " completed (" & GetFileName(SocketItem(TheSock).FileName) & ")"
                RemoveSocket TheSock, 0, "", True
              End If
            End If
          End If
        End If
      
      Case SF_Status_Bot, SF_Status_BotLinking, SF_Status_BotPreCache, SF_Status_BotGetName, SF_Status_BotGetPass
        BotNetSort TheSock, Data
      
      Case Else
        SNum = SocketItem(TheSock).SocketNumber
        PartySort TheSock, SNum, Data
    End Select
  End If
End Sub

Public Sub Socket_DataUDP(ByVal TheSock As Long, ByVal Data As String, ByVal Host As String, ByVal Port As Long)
  Dim u As Long
  Select Case GetSockFlag(TheSock, SF_Status)
    Case SF_Status_ScriptSocket
      For u = 1 To ScriptCount
        If Scripts(u).Name = SocketItem(TheSock).SetupChan Then
          RunScriptX u, SocketItem(TheSock).CurrentQuestion, TheSock, 2, Host, Port, Data
        End If
      Next u
    
    Case Else
      RemoveSocket TheSock, 0, "", True
  End Select
End Sub

Public Sub Socket_Error(ByVal vSocket As Long, ByVal Critical As Boolean, ByVal Number As Long, ByVal Description As String)
End Sub

Public Sub Socket_Incoming(ByVal TheSock As Long, ByVal NewSock As Long, ByVal Host As String, ByVal Port As Long)
  Dim HostAddr As String
  Dim u As Long
  Select Case GetSockFlag(TheSock, SF_Status)
    Case SF_Status_TelnetListen
      HostAddr = "*!telnet@" & GetCacheHost(Host, True)
      If (HostAddr = "*!telnet@Unknown") Or (InStr(HostAddr, ".") = 0) Then HostAddr = "*!telnet@" & Host
      If Not IsIgnoredTHost(HostAddr) Then
        SocketItem(NewSock).Hostmask = HostAddr
        SetSockFlag NewSock, SF_Status, SF_Status_UserGetName
        SetSockFlag NewSock, SF_Echo, SF_YES
        SetSockFlag NewSock, SF_DCC, SF_NO
        SetSockFlag NewSock, SF_LF_ONLY, SF_NO
        SetSockFlag NewSock, SF_Telnet, SF_YES
        SocketItem(NewSock).OnBot = BotNetNick
        SocketItem(NewSock).RegNick = ""
        SocketItem(NewSock).IsInternalSocket = True
        Wait 250
        If SocketItem(NewSock).SocketFlag(SF_Telnet) = SF_YES Then
          SendTCP NewSock, T_IAC & T_WILL & T_SPGA & T_IAC & T_WILL & T_ECHO
          SendTCP NewSock, "[7l[1;37;40m[2j"
        End If
        TU NewSock, "   ______________________________   "
        TU NewSock, "  |                              || "
        TU NewSock, "  |   AnGeL Telnet Server V1.2   || "
        TU NewSock, "  |   (C) 2002 by T. Schiemann   || "
        TU NewSock, "  |               B. Huebschen   || "
        TU NewSock, "  |______________________________|| "
        TU NewSock, "   -------------------------------' "
        TU NewSock, " "
        TU NewSock, "Please enter your user name:"
        SpreadFlagMessageEx 0, "+m", SF_Local_JP, MakeMsg(MSG_PLTelNetIncoming, SocketItem(NewSock).Hostmask)
      Else
        RemoveSocket NewSock, 0, "", True
      End If

    Case SF_Status_BotnetListen
      HostAddr = "*!botnet@" & GetCacheHost(Host, True)
      If (HostAddr = "*!*@Unknown") Or (InStr(HostAddr, ".") = 0) Then HostAddr = "*!botnet@" & Host
      If Not IsIgnoredTHost(HostAddr) Then
        SocketItem(NewSock).Hostmask = HostAddr
        SetSockFlag NewSock, SF_Status, SF_Status_BotGetName
        SetSockFlag NewSock, SF_Echo, SF_NO
        SetSockFlag NewSock, SF_DCC, SF_YES
        SetSockFlag NewSock, SF_LF_ONLY, SF_YES
        SocketItem(NewSock).OnBot = BotNetNick
        SocketItem(NewSock).RegNick = ""
        SocketItem(NewSock).OrderSign = CStr(Timer)
        SocketItem(NewSock).IsInternalSocket = True
        TU NewSock, "AnGeL Telnet Server V1.0"
        TU NewSock, " "
        TU NewSock, "Please enter your user name."
        SpreadFlagMessageEx 0, "+t", SF_Local_Bot, MakeMsg(MSG_PLBotNetIncoming, SocketItem(NewSock).Hostmask)
      Else
        RemoveSocket NewSock, 0, "", True
      End If
    
    Case SF_Status_IdentListen
      SocketItem(NewSock).IsInternalSocket = True
      SocketItem(NewSock).RegNick = "<IDENT>"
      SetSockFlag NewSock, SF_Status, SF_Status_Ident
      SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLIdentRequest, Host)

    Case SF_Status_ScriptSocket
      For u = 1 To ScriptCount
        If Scripts(u).Name = SocketItem(TheSock).SetupChan Then
          SetSockFlag NewSock, SF_Status, SF_Status_ScriptSocket
          SocketItem(NewSock).CurrentQuestion = SocketItem(TheSock).CurrentQuestion
          SocketItem(NewSock).SetupChan = SocketItem(TheSock).SetupChan
          SocketItem(NewSock).UserNum = 0
          SocketItem(NewSock).OnBot = ""
          SocketItem(NewSock).IRCNick = ""
          SocketItem(NewSock).RegNick = Scripts(u).Name
          SocketItem(NewSock).IsInternalSocket = True
          HostAddr = GetCacheHost(Host, True)
          If (HostAddr = "Unknown") Or (InStr(HostAddr, ".") = 0) Then HostAddr = Host
          RunScriptX u, SocketItem(TheSock).CurrentQuestion, NewSock, 5, HostAddr
        End If
      Next u

    Case SF_Status_DCCInit
      u = SocketItem(TheSock).SocketNumber
      SocketItem(TheSock).SocketNumber = SocketItem(NewSock).SocketNumber
      SocketItem(TheSock).LocalAddress = SocketItem(NewSock).LocalAddress
      SocketItem(TheSock).RemoteAddress = SocketItem(NewSock).RemoteAddress
      SocketItem(NewSock).SocketNumber = u
      RemoveSocket NewSock, 0, "", True
      If BotUsers(SocketItem(TheSock).UserNum).Password <> "" Then
        SetSockFlag TheSock, SF_Status, SF_Status_UserGetPass
        SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCOpened, SocketItem(TheSock).RegNick)
        TU TheSock, MakeMsg(MSG_EnterPWD, SocketItem(TheSock).RegNick)
      Else
        SetSockFlag TheSock, SF_Status, SF_Status_UserChoosePass
        SpreadFlagMessageEx TheSock, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCFirst, SocketItem(TheSock).RegNick)
        TU TheSock, MakeMsg(MSG_ChoosePWD, SocketItem(TheSock).RegNick)
      End If

    Case SF_Status_SendFileWaiting
      u = SocketItem(TheSock).SocketNumber
      SocketItem(TheSock).SocketNumber = SocketItem(NewSock).SocketNumber
      SocketItem(TheSock).LocalAddress = SocketItem(NewSock).LocalAddress
      SocketItem(TheSock).RemoteAddress = SocketItem(NewSock).RemoteAddress
      SocketItem(NewSock).SocketNumber = u
      RemoveSocket NewSock, 0, "", True
      If Dir(SocketItem(TheSock).FileName) <> "" Then
        SetSockFlag TheSock, SF_Status, SF_Status_SendFile
        SocketItem(TheSock).FileNum = FreeFile: Open SocketItem(TheSock).FileName For Binary Shared As #SocketItem(TheSock).FileNum
        If Err.Number > 0 Then
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " failed (not found: " & GetFileName(SocketItem(TheSock).FileName) & ")!"
          SendNextQueuedFile SocketItem(TheSock).RegNick
          Close #SocketItem(TheSock).FileNum
          RemoveSocket TheSock, 0, "", True
        Else
          SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " starting  (" & GetFileName(SocketItem(TheSock).FileName) & ")"
          u = PumpDCC: If SocketItem(TheSock).BytesReceived + u > SocketItem(TheSock).FileSize Then u = SocketItem(TheSock).FileSize - SocketItem(TheSock).BytesReceived
          If u > 0 Then
            If SocketItem(TheSock).BytesReceived > 0 Then Seek SocketItem(TheSock).FileNum, SocketItem(TheSock).BytesReceived + 1
            HostAddr = Space(u): Get SocketItem(TheSock).FileNum, , HostAddr
            SocketItem(TheSock).BytesReceived = SocketItem(TheSock).BytesReceived + u
            SendTCP TheSock, HostAddr
          Else
            Close #SocketItem(TheSock).FileNum
            SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " completed (" & GetFileName(SocketItem(TheSock).FileName) & ")"
            For u = 1 To ScriptCount
              If Scripts(u).Hooks.fa_downloadcomplete Then
                RunScriptX u, "fa_downloadcomplete", SocketItem(TheSock).RegNick, GetFileName(SocketItem(TheSock).FileName)
              End If
            Next u
            SendNextQueuedFile SocketItem(TheSock).RegNick
            RemoveSocket TheSock, 0, "", True
          End If
        End If
      Else
        SpreadLevelFileAreaMessage TheSock, "14[" & Time & "] *** DCC send to " & SocketItem(TheSock).RegNick & " failed (not found: " & GetFileName(SocketItem(TheSock).FileName) & ")!"
        SendNextQueuedFile SocketItem(TheSock).RegNick
        RemoveSocket TheSock, 0, "", True
      End If
    
    Case Else
      RemoveSocket TheSock, 0, "", True
      RemoveSocket NewSock, 0, "", True
  End Select
End Sub

Public Sub Socket_ResolvedHost(ByVal Host As String, ByVal IP As String)
  AddCacheData Host, IP
End Sub

Public Sub Socket_ResolveFailed(ByVal Host As String)
End Sub
