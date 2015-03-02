Attribute VB_Name = "Server_Parser"
Option Explicit


Public Sub Handle(ByVal Block As String) ' : AddStack "mdlWinsock_Handle(" & Block & ")"
  Dim u As Long, Char As String
  Block = Replace(Block, Chr(13), "")
  ServBuff = ServBuff & Block
  
  'Output Block
  u = InStr(1, ServBuff, Chr(10), vbBinaryCompare)
  While u <> 0
    Char = Left(ServBuff, u - 1)
    ServBuff = Mid(ServBuff, u + 1)
    GotOneLine Char
    u = InStr(1, ServBuff, Chr(10), vbBinaryCompare)
  Wend
End Sub

Public Sub GotOneLine(ByVal Line As String)
  Dim Par() As String, Dummy As String, ScNum As Long
  Par = ParamArr(Line)
  
  ' Werte alles als PING reply
  If Not GotServerPong Then GotServerPong = True
  LastEvent = Now
  
  ' Scripts übergeben
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.Raw = True And Scripts(ScNum).Hooks.RawFilter = "" Then
      RunScriptX ScNum, "Raw", Line
    ElseIf Scripts(ScNum).Hooks.Raw And InStr(LCase(Scripts(ScNum).Hooks.RawFilter), LCase(Param(Line, 2))) <> 0 Then
      RunScriptX ScNum, "Raw", Line
    End If
  Next ScNum
  
  ' Ausgeben
  Output Line & vbCrLf

  If Left(Par(1), 1) = ":" Then
    ' Anweisungen entfernen
      If InStr(1, Par(1), "!") > 1 Then Dummy = StripDP(Left(Par(1), InStr(Par(1), "!") - 1)) Else Dummy = StripDP(Par(1))
      If Dummy = MyNick Then
      If IsOrdered(Par(2) & " " & Par(3)) Then
        Output "removed order " & Par(2) & " " & Par(3) & vbCrLf
      End If
      RemOrder Par(2) & " " & Par(3)
    End If
      
    If IsNumeric(Par(2)) Then
      ' Scripts übergeben
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Numerics Then
          RunScriptX ScNum, "Numerics", GetRest(Line, 2)
        End If
      Next ScNum
      
      ' Zahlenwert
      If Par(2) < 300 Then
        '001-299
        Select Case CLng(Par(2))
          Case 1
            'RPL_WELCOME, ":Welcome to the DALnet IRC Network %s",
            RPL_WELCOME Line
          Case 4
            'RPL_MYINFO, "%s %s oiwsg biklmnopstv",
            RPL_MYINFO Line
          Case 5
            'RPL_ISUPPORT, "NOQUIT TOKEN WATCH=128 SAFELIST  are available on this server"
            ' - alternativ -
            'RPL_MAP / RPL_BOUNCE, kein einheitlicher syntax :/
            If InStr(1, Line, ":are supported by this server", vbTextCompare) > 0 Then RPL_ISUPPORT Line
        End Select
      Else
        If Par(2) < 400 Then
          ' 300-399
          Select Case CLng(Par(2))
            Case 302
              'RPL_USERHOST, :*1<reply> *( " " <reply> )
              RPL_USERHOST Line
            Case 303
              'RPL_ISON, :<nick>
              RPL_ISON Line
            Case 324
              'RPL_CHANNELMODEIS, <channel> <mode> <mode params>
              RPL_CHANNELMODEIS Line
            Case 331
              'RPL_NOTOPIC, "<channel> :No topic is set"
            Case 332
              'RPL_TOPIC, "<channel> :<topic>"
              RPL_TOPIC Line
            Case 346
              'RPL_INVITELIST, "<channel> <invitemask>"
              RPL_INVITELIST Line
            Case 347
              'RPL_ENDOFINVITELIST, "<channel> :End of channel invite list"
              RPL_ENDOFINVITELIST Line
            Case 348
              'RPL_EXCEPTLIST, "<channel> <exceptionmask>"
              RPL_EXCEPTLIST Line
            Case 349
              'RPL_ENDOFEXCEPTLIST, "<channel> :End of channel exception list"
              RPL_ENDOFEXCEPTLIST Line
            Case 352
              'RPL_WHOREPLY, "<channel> <user> <host> <server> <nick> ( "H" / "G" > ["*"] [ ( "@" / "+" ) ] :<hopcount> <real name>"
              RPL_WHOREPLY Line
            Case 315
              'RPL_ENDOFWHO, "<name> :End of WHO list"
              RPL_ENDOFWHO Line
            Case 353
              'RPL_NAMREPLY, "( "=" / "*" / "@" ) <channel> :[ "@" / "+" ] <nick> *( " " [ "@" / "+" ] <nick> )  - "@" is used for secret channels, "*" for private channels, and "=" for others (public channels).
              RPL_NAMREPLY Line
            Case 366
              'RPL_ENDOFNAMES, "<channel> :End of NAMES list"
            Case 367
              'RPL_BANLIST, "<channel> <banmask>"
              RPL_BANLIST Line
            Case 368
              'RPL_ENDOFBANLIST, "<channel> :End of channel ban list"
              RPL_ENDOFBANLIST Line
    '''            Case 375
    '''              'RPL_MOTDSTART, ":- <server> Message of the day - "
    '''            Case 372
    '''              'RPL_MOTD, ":- <text>"
    '''            Case 376
    '''              'RPL_ENDOFMOTD, ":End of MOTD command"
          End Select
        Else
          ' 400++
          Select Case CLng(Par(2))
            Case 401
              'ERR_NOSUCHNICK, "<nickname> :No such nick/channel"
              ERR_NOSUCHNICK Line
            Case 403
              'ERR_NOSUCHCHANNEL, "<channel name> :No such channel"
              ERR_NOSUCHCHANNEL Line
            Case 405
              'ERR_TOOMANYCHANNELS, "<channel name> :You have joined too many channels"
              ERR_TOOMANYCHANNELS Line
            Case 432
              'ERR_ERRONEUSNICKNAME, "<nick> :Erroneous nickname"
            Case 433
              'ERR_NICKNAMEINUSE, "<nick> :Nickname is already in use"
              ERR_NICKNAMEINUSE Line
            Case 437
              'ERR_UNAVAILRESOURCE, "<nick/channel> :Nick/channel is temporarily unavailable"
              ERR_UNAVAILRESOURCE Line
            Case 442
              'ERR_NOTONCHANNEL, "<channel> :You're not on that channel"
            Case 471
              'ERR_CHANNELISFULL, "<channel> :Cannot join channel (+l)"
              ERR_CHANNELISFULL Line
            Case 473
              'ERR_INVITEONLYCHAN, "<channel> :Cannot join channel (+i)"
              ERR_INVITEONLYCHAN Line
            Case 474
              'ERR_BANNEDFROMCHAN, "<channel> :Cannot join channel (+b)"
              ERR_BANNEDFROMCHAN Line
            Case 475
              'ERR_BADCHANNELKEY, "<channel> :Cannot join channel (+k)"
              ERR_BADCHANNELKEY Line
            Case 477
              'ERR_NOCHANMODES, "<channel> :Channel doesn't support modes
              ' - alternativ -
              'ERR_REGISTERONLY, "<channel> :Cannot join channel (+r)"
              If Par(8) = "(+r)" Then ERR_REGISTERONLY Line
            Case 484
              'ERR_RESTRICTED, ":Your connection is restricted!"
          End Select
        End If
      End If
    Else
      ' Wort
      Par(2) = UCase(Par(2))
      If Left(Par(2), 1) = "P" Then
        ' Ist ein P
        Select Case Par(2)
          Case "PONG"
            SYS_PONG Line
          Case "PART"
            SYS_PART Dummy, Line
          Case "PRIVMSG"
            SYS_PRIVMSG Dummy, Line
          Case Else
            Trace "*** unknown", Line
        End Select
      Else
        ' Ist KEIN P
        Select Case Par(2)
          Case "NICK"
            SYS_NICK Dummy, Line
          Case "JOIN"
            SYS_JOIN Dummy, Line
          Case "QUIT"
            SYS_QUIT Dummy, Line
          Case "MODE"
            SYS_MODE Dummy, Line
          Case "KICK"
            SYS_KICK Dummy, Line
          Case "NOTICE"
            SYS_NOTICE Dummy, Line
          Case "INVITE"
            SYS_INVITE Line
          Case "TOPIC"
            SYS_TOPIC Dummy, Line
          Case Else
            Trace "*** unknown", Line
        End Select
      End If
    End If
  Else
    ' Erstes Zeichen KEIN Doppelpunkt
    If UCase(Par(1)) = "PING" Then
      ' PING x'
      SYS_PING Line
    ElseIf UCase(Par(1)) = "ERROR" Then
      ' ERROR :Ping timeout
      SYS_ERROR Line
    ElseIf UCase(Par(1)) = "NOTICE" Then
      ' NOTICE AUTH :*** Found your hostname.
    Else
      Trace "*** unknown", Line
    End If
  End If
End Sub

Sub SYS_PING(Line As String)
  'Entblocken
  BlockSends = False
  If Connected Then
    'Über Buffer senden
    SendLine "PONG " & Param(Line, 2), 1
  Else
    'Sofort senden
    SendIt "PONG " & Param(Line, 2) & vbCrLf
  End If
End Sub

Sub SYS_PONG(Line As String)
  If BlockSends Then BlockSends = False
  Output "*** Unblocked Sends" & vbCrLf
End Sub

Sub SYS_ERROR(Line As String)
  SpreadFlagMessage 0, "+m", "4*** SERVER " & Line
End Sub

Sub SYS_INVITE(Line As String)
  Dim Index As Long
  Dim Chan As String
  Dim Rest As String
  If Param(Line, 3) = MyNick Then
    'Wenn der Bot geladen wurde
    Chan = StripDP(Param(Line, 4))
    If InAutoJoinChannels(Chan) Then
      Index = FindChan(Chan)
      If Index = 0 Then
        If Not IsOrdered("JOIN :" & Chan) Then
          Rest = GetChannelSetting(Chan, "EnforceModes", "x")
          If ParamCount(Rest) > 1 Then Rest = Param(Rest, ParamCount(Rest)) Else Rest = "x"
          Order "JOIN :" & Chan, 5
          SendLine "JOIN " & Chan & " " & Rest, 1
        End If
      End If
    End If
  End If
End Sub

Sub SYS_QUIT(Nick As String, Line As String)
  Dim ChNum As Long, UsNum As Long, ScNum As Long
  Dim RegUser As String, UserFlags As String, Rest As String
  
  ' Quit Message schnappen
  If Param(Line, 3) <> "" Then Rest = Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & "  ")) Else Rest = ""
  
  ' Channels durchgehen
  For ChNum = 1 To ChanCount
    UsNum = FindUser(Nick, ChNum)
    If UsNum > 0 Then
      RegUser = Channels(ChNum).User(UsNum).RegNick
      UserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Channels(ChNum).Name)
      
      'Check script Hooks
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Quit Then
          RunScriptX ScNum, "Quit", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, Rest
        End If
      Next ScNum
      
      'Generate Partyline Message
      SpreadPartylineChanEvent "quit", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", Rest
      
      'Write seen information
      If RegUser <> "" Then
        WriteSeenEntry RegUser, "", Now, Channels(ChNum).Name, Rest, Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      End If
      
      'Write extended seen information for unknown users
      If LCase(RegUser) <> LCase(Nick) Then
        WriteExtSeenEntry Nick, RegUser, Now, Channels(ChNum).Name, Rest, Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      End If
      
      'Remove Kick Requests
      UpdateBufferKicks Nick, Channels(ChNum).Name
      RemKickUser ChNum, Nick
      
      'Remove User from Channel
      RemChanUser Nick, ChNum
      
      'Check Opless
      If Channels(ChNum).UserCount = 1 And Channels(ChNum).GotOPs = False Then
        SendLine "PART " & Channels(ChNum).Name & vbCrLf & "JOIN " & Channels(ChNum).Name, 1
        SpreadFlagMessage 0, "+m", "3*** Rejoining " & Channels(ChNum).Name & " to get ops..."
      End If
    End If
  Next ChNum
End Sub

Sub SYS_KICK(Nick As String, Line As String)
  Dim Chan As String, ChNum As Long, UsNum As Long, ScNum As Long, KUsNum As Long, KNick As String, KUserFlags As String, UserFlags As String, Reason As String, Rest As String, KickedUser As Boolean
    
  'Informationen Sammeln
  Chan = Param(Line, 3)
  KNick = Param(Line, 4)
  ChNum = FindChan(Chan)
  UsNum = FindUser(Nick, ChNum)
  KUsNum = FindUser(KNick, ChNum)
  Reason = StripDP(GetRest(Line, 5))
  
  If ChNum > 0 And KUsNum > 0 Then
    'Wenn Channel und Opfer ;) bekannt sind
    If UsNum > 0 Then UserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Chan) Else UserFlags = ""
    KUserFlags = GetUserChanFlags2(Channels(ChNum).User(KUsNum).UserNum, Chan)
    If KNick = MyNick Then
      'Ich selber wurde gekickt
      SetPermChanStat Chan, ChanStat_NotOn
      SpreadFlagMessage 0, "+m", "4*** I was kicked off the channel " & Chan & " by " & Nick & IIf(Reason <> "", " (" & Reason & "4)", "!")
      Rest = GetChannelKey(ChNum)
      If Channels(ChNum).CompletedWHO Then
        For ScNum = 1 To ScriptCount
          If Scripts(ScNum).Hooks.Kick Then
            RunScriptX ScNum, "Kick", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, KNick, Channels(ChNum).User(KUsNum).RegNick, KUserFlags, Reason
          End If
        Next ScNum
        SpreadPartylineChanEvent "kick", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, KNick, Channels(ChNum).User(KUsNum).RegNick, KUserFlags, Reason
      End If
      RemChan Chan
      If Rest = "" Then SendLine "join " & Chan, 1 Else SendLine "join " & Chan & " " & Rest, 1
    Else
      'Jemand anders wurde gekickt
      KickedUser = False
      If Channels(ChNum).CompletedWHO Then
        'Check script Hooks
        HaltDefault = False
        For ScNum = 1 To ScriptCount
          If Scripts(ScNum).Hooks.Kick Then
            RunScriptX ScNum, "Kick", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, KNick, Channels(ChNum).User(KUsNum).RegNick, KUserFlags, Reason
          End If
        Next ScNum
        SpreadPartylineChanEvent "kick", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, KNick, Channels(ChNum).User(KUsNum).RegNick, KUserFlags, Reason
        If UsNum > 0 Then
          'Bekannter Chatter
          If HaltDefault = False Then
            If Channels(ChNum).ProtectFriends Then
              If MatchFlags(KUserFlags, "+f") And MatchFlags(UserFlags, "-f") And (Nick <> MyNick) Then
                PatternKickBan Nick, Chan, "4Don't kick my friends!", True, 30, 50
                KickedUser = True
              End If
            End If
            If (KickedUser = False) And MatchFlags(UserFlags, "-bnr") And (Nick <> MyNick) And MatchFlags(KUserFlags, "+r") Then
              SendLine "mode " & Chan & " -o+b " & Nick & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 1), 1
              SendLine "kick " & Chan & " " & Nick & " :4Special revenge for " & KNick & "!", 1
              TimedEvent "UnBan " & Chan & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 1), 20
              KickedUser = True
            End If
          End If
        End If
      End If
      
      If Channels(ChNum).User(KUsNum).RegNick <> "" Then
        'Eintrag im SeenLog
        WriteSeenEntry Channels(ChNum).User(KUsNum).RegNick, "", Now, Channels(ChNum).Name, "*kicked* " & Nick, Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      End If
      
      If (LCase(Channels(ChNum).User(KUsNum).RegNick) <> LCase(KNick)) And (Not Channels(ChNum).InFlood) Then
        'Eintrag in ExtSeen
        WriteExtSeenEntry KNick, Channels(ChNum).User(KUsNum).RegNick, Now, Channels(ChNum).Name, "*kicked* " & Nick, Mask(Channels(ChNum).User(KUsNum).Hostmask, 10)
      End If
      
      UpdateBufferKicks KNick, Chan
      RemKickUser ChNum, KNick
      
      RemChanUser KNick, ChNum
    End If
  End If
End Sub

Sub SYS_NICK(Nick As String, Line As String)
  Dim u As Long, NewNick As String, UserFlags As String, MatchedOne As Boolean, OUserNum As Long, Rest As String, ChNum As Long, UsNum As Long, ScNum As Long
  
  NewNick = StripDP(Param(Line, 3))
  
  UpdateBufferNicks Nick, NewNick
  
  'Keep track of IRC nicks
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If SocketItem(u).IRCNick = Nick Then SocketItem(u).IRCNick = NewNick
    End If
  Next u
  For u = 1 To QueuedFiles
    If FileSendQueue(u).IRCNick = Nick Then FileSendQueue(u).IRCNick = NewNick
  Next u
  
  'Update Textflood-Nicks
  For u = 1 To EventCount
    If Param(Events(u).DoThis, 3) = Nick Then
      Rest = Param(Events(u).DoThis, 1)
      If (Rest = "RemLine") Or (Rest = "RemRepeat") Or (Rest = "Op") Then
        Events(u).DoThis = Rest & " " & Param(Events(u).DoThis, 2) & " " & NewNick
      ElseIf Rest = "RemChars" Then
        Events(u).DoThis = Replace(Events(u).DoThis, " " & Nick & " ", " " & NewNick & " ")
      End If
    End If
  Next u
  
  'Update Channels
  For ChNum = 1 To ChanCount
    'Update KickList
    For u = 1 To Channels(ChNum).KickCount
      If LCase(Channels(ChNum).KickList(u).Nick) = LCase(Nick) Then
        Channels(ChNum).KickList(u).Nick = NewNick
        Channels(ChNum).KickList(u).Hostmask = NewNick & "!" & Mask(Channels(ChNum).KickList(u).Hostmask, 10)
      End If
    Next u
    'Update NickList
    UsNum = FindUser(Nick, ChNum)
    If UsNum > 0 Then
      Channels(ChNum).User(UsNum).Nick = NewNick
      Channels(ChNum).User(UsNum).Hostmask = NewNick & "!" & Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      If Channels(ChNum).User(UsNum).IPmask <> "" Then Channels(ChNum).User(UsNum).IPmask = NewNick & "!" & Mask(Channels(ChNum).User(UsNum).IPmask, 10)
      Channels(ChNum).User(UsNum).NickChanges = Channels(ChNum).User(UsNum).NickChanges + 1
      UserFlags = GetUserChanFlags(Channels(ChNum).User(UsNum).RegNick, Channels(ChNum).Name)
      
      'Check script Hooks
      HaltDefault = False
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Nick Then
          RunScriptX ScNum, "Nick", Channels(ChNum).Name, Nick, NewNick, Channels(ChNum).User(UsNum).RegNick, UserFlags, Channels(ChNum).User(UsNum).NickChanges
        End If
      Next ScNum
      
      SpreadPartylineChanEvent "nick", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, NewNick, "", "", ""
      
      'Nick Flood Protection
      If (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) And (Nick <> MyNick) And (Not HaltDefault) Then
        MatchedOne = IsFloodNick(Channels(ChNum).User(UsNum).Nick)
        'Flood during normal state:   4 nicks / 10 sec (with normal nicks)
        '                             2 nicks / 10 sec (with flood nicks like "r7b1c0o9g" or "}_][]}[[{")
        'Flood while chan is flooded: 2 nicks / 10 sec (with all nicks)
        If ((MatchedOne = False) And (Channels(ChNum).User(UsNum).NickChanges = 4) And MatchFlags(UserFlags, "-bn")) Or ((MatchedOne Or Channels(ChNum).InFlood) And (Channels(ChNum).User(UsNum).NickChanges = 2) And MatchFlags(UserFlags, "-bn")) Then
          FloodCheck ChNum, "+smtin", 3
          If AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, 2)) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), UnbanTime
          AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, "Nick flood"
          SpreadPartylineChanEvent "nickflood", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", ""
          Exit Sub
        End If
      End If
      
      'Nick Protection
      If MatchFlags(GetUserChanFlags(Channels(ChNum).User(UsNum).Nick, Channels(ChNum).Name), "+x") And (Nick <> MyNick) Then
        OUserNum = GetUserNum(Channels(ChNum).User(UsNum).Nick)
        If OUserNum > 0 Then
          MatchedOne = False
          For u = 1 To BotUsers(OUserNum).HostMaskCount
            If MatchHost(BotUsers(OUserNum).HostMasks(u), Channels(ChNum).User(UsNum).Hostmask) Then MatchedOne = True: Exit For
          Next u
          If Not MatchedOne Then
            If MatchFlags(UserFlags, "+n") Then
              SendLine "privmsg " & Channels(ChNum).Name & " :" & Channels(ChNum).User(UsNum).Nick & ": Please choose another nick, the one you are`using is protected.", 3
            Else
              If InStr(1, Channels(ChNum).User(UsNum).Status, "@", vbBinaryCompare) > 0 Then SendLine "mode " & Channels(ChNum).Name & " -o+b " & Nick & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 8), 2
              If Channels(ChNum).User(UsNum).Status <> "@" Then SendLine "mode " & Channels(ChNum).Name & " +b " & Mask(Channels(ChNum).User(UsNum).Hostmask, 8), 2
              SendLine "kick " & Channels(ChNum).Name & " " & Channels(ChNum).User(UsNum).Nick & " :Nick faker! Change your nick!", 2
              TimedEvent "notice " & Channels(ChNum).User(UsNum).Nick & " :Please use another nick, the one you are using is protected.", 4
              TimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 8), 120
            End If
          End If
        End If
      End If
    End If
  Next ChNum
  
  'I changed my nick
  If Nick = MyNick Then
    Output "*** I'm now known as " & NewNick & vbCrLf
    SpreadFlagMessage 0, "+m", "3*** I changed my IRC nick to " & NewNick
    MyNick = NewNick
    MyHostmask = MyNick & "!" & Mask(MyHostmask, 10)
    If MyIPmask <> "" Then MyIPmask = MyNick & "!" & Mask(MyIPmask, 10)
  End If
End Sub

Sub SYS_MODE(Nick As String, Line As String)
  Dim Index As Integer
  Dim Chan As String, ONick As String, ChNum As Long, UsNum As Long, UserFlags As String, OpString As String
  Dim Hostmask As String, KNick As String, KUsNum As Long, KUserFlags As String, RemovedOne As Boolean
  Dim IsF As Boolean, IsM As Boolean, IsN As Boolean, IsS As Boolean, IsB As Boolean
  Dim JustGainedOps As Boolean, JustGainedHOps As Boolean, JustLostOps As Boolean, TookOp As Boolean, ReOpped As Boolean
  Dim u As Long, u2 As Long, PosInONick As Long, PlusOrMinus As Byte, MatchedOne As Boolean
  Dim ChangMode As String, ScNum As Long, sNewStatus As String
  
  If IsValidChannel(Left(Param(Line, 3), 1)) Then
    'Channel Mode
    Chan = StripDP(Param(Line, 3))
    ONick = GetRest(Line, 5)
    ChNum = FindChan(Chan)
    If ChNum <> 0 Then
      UsNum = FindUser(Nick, ChNum)
      UserFlags = GetUserChanFlags(Channels(ChNum).User(UsNum).RegNick, Chan)
      OpString = LCase(Param(Line, 4))
      IsF = MatchFlags(UserFlags, "+f")
      IsM = MatchFlags(UserFlags, "+m")
      IsN = MatchFlags(UserFlags, "+n")
      IsS = MatchFlags(UserFlags, "+s")
      IsB = MatchFlags(UserFlags, "+b")
      
      'Mass deop protection
      If (Left(OpString, 4) = "-ooo") And (Nick <> MyNick) Then
        If Channels(ChNum).GotOPs Then
          If IsB = False Then
            If IsF = False Then
              SpreadMessage 0, -1, "4*** Detected a mass deop from " & Nick & " in " & Chan & " - kicking/banning (no +f)..."
              PatternKickBan Nick, Chan, "4Mass Deop! Have a nice day... OUTSIDE", True, 40, 60
              TookOp = True
            ElseIf IsM = False Then
              SpreadMessage 0, -1, "4*** Detected a mass deop from " & Nick & " in " & Chan & " - kicking (no +m)..."
              SendLine "kick " & Chan & " " & Nick & " :Mass Deop", 1
              TookOp = True
            ElseIf IsN = False Then
              SpreadMessage 0, -1, "14*** Detected a mass deop from " & Nick & " in " & Chan & " - deopping (no +n)"
              SendLine "mode " & Chan & " -o " & Nick, 1
              TookOp = True
            Else
              SpreadMessage 0, -1, "14*** Detected a mass deop from " & Nick & " in " & Chan & " - doing nothing (+n)"
            End If
          Else
            SpreadMessage 0, -1, "14*** Detected a mass deop from " & Nick & " in " & Chan & " - doing nothing (+b)"
          End If
        Else
          SpreadMessage 0, -1, "14*** Detected a mass deop from " & Nick & " in " & Chan
        End If
      End If
      
      Channels(ChNum).Mode = FormatMode(Channels(ChNum).Mode, GetRest(Line, 4))
      PosInONick = 0
      JustGainedOps = False
      JustLostOps = False
      For u = 1 To Len(OpString)
        Select Case Mid(OpString, u, 1)
          Case "+": PlusOrMinus = 1
          Case "-": PlusOrMinus = 2
          Case "l": If GetModeChar(OpString, "l") = 1 Then PosInONick = PosInONick + 1
          Case "k": PosInONick = PosInONick + 1
          Case "v"
            PosInONick = PosInONick + 1
            'VOICE
            If PlusOrMinus = 1 Then
              KUsNum = FindUser(Param(ONick, PosInONick), ChNum)
              sNewStatus = ""
              If InStr(Channels(ChNum).User(KUsNum).Status, "@") > 0 Then sNewStatus = sNewStatus & "@"
              If InStr(Channels(ChNum).User(KUsNum).Status, "%") > 0 Then sNewStatus = sNewStatus & "%"
              sNewStatus = sNewStatus & "+"
              Channels(ChNum).User(KUsNum).Status = sNewStatus
            End If
            'DEVOICE
            If PlusOrMinus = 2 Then
              KUsNum = FindUser(Param(ONick, PosInONick), ChNum)
              sNewStatus = ""
              If InStr(Channels(ChNum).User(KUsNum).Status, "@") > 0 Then sNewStatus = sNewStatus & "@"
              If InStr(Channels(ChNum).User(KUsNum).Status, "%") > 0 Then sNewStatus = sNewStatus & "%"
              Channels(ChNum).User(KUsNum).Status = sNewStatus
            End If
          Case "b"
            PosInONick = PosInONick + 1
            'ON BAN
            If PlusOrMinus = 1 Then
              Hostmask = Param(ONick, PosInONick)
              'Cut by Server but set
              If Hostmask = "" Then
                SendLine "mode " & Chan & " +b", 1
              Else
                'Check script Hooks
                For ScNum = 1 To ScriptCount
                  If Scripts(ScNum).Hooks.Ban Then
                    RunScriptX ScNum, "Ban", Nick, Channels(ChNum).User(UsNum).RegNick, Chan, Hostmask
                  End If
                Next ScNum
                
                'Remove Ban Orders
                RemOrder "ban " & Chan & " " & Hostmask
                RemDesiredBan ChNum, Hostmask
                'Add Bans to Banlist
                MatchedOne = False
                For u2 = 1 To Channels(ChNum).BanCount
                  If LCase(Channels(ChNum).BanList(u2).Mask) = LCase(Hostmask) Then MatchedOne = True
                Next u2
                If Not MatchedOne Then
                  Channels(ChNum).BanCount = Channels(ChNum).BanCount + 1
                  Channels(ChNum).BanList(Channels(ChNum).BanCount).Mask = Hostmask
                  Channels(ChNum).BanList(Channels(ChNum).BanCount).CreatedAt = Now
                End If
                'Selfban protection
                RemovedOne = False
                If Nick <> MyNick Then
                  If (MatchWM(Hostmask, MyHostmask) = True) Or (MatchWM(Hostmask, MyIPmask) = True) Then
                    If (IsS = False) Then  'Only allow super owners to ban me
                      If (IsN = False) And (IsB = False) Then AddMassMode "-o", Nick
                      AddMassMode "-b", Hostmask
                      RemovedOne = True
                    End If
                  End If
                End If
                'Protect friends
                If Channels(ChNum).ProtectFriends = True Then
                  If Nick <> MyNick Then
                    KNick = SearchUserFromHostmask2(Hostmask)
                    KUserFlags = GetUserChanFlags(KNick, Chan)
                    'Allow bots to ban everybody
                    If (IsB = False) Then
                      'If ban matches a user with +f and banner doesn't have +f -> deop, unban
                      'If a user without +n bans a bot with +o for the channel -> deop, unban
                      If (MatchFlags(KUserFlags, "+f") And (IsF = False)) Or (MatchFlags(KUserFlags, "+bo") And (IsN = False) And (IsS = False)) Then
                        AddMassMode "-o", Nick
                        AddMassMode "-b", Hostmask
                        RemovedOne = True
                      Else
                        'If banner has a lower level than banned user -> unban
                        If LevelNumber(KUserFlags) > LevelNumber(UserFlags) Then
                          AddMassMode "-b", Hostmask
                          RemovedOne = True
                        End If
                      End If
                    End If
                  End If
                End If
                'Enforce bans
                If (Channels(ChNum).EnforceBans = True) And (RemovedOne = False) Then
                  If Channels(ChNum).InFlood = False Then
                    'Wait 3 sec before kicking to give banner a chance
                    TimedEvent "CheckBans " & Channels(ChNum).Name & " " & Hostmask & " " & Nick, 3
                  Else
                    'Kick immediately during floods
                    TimedEvent "CheckBans " & Channels(ChNum).Name & " " & Hostmask & " " & Nick, Now
                  End If
                End If
              End If
            End If
            'ON UNBAN
            If PlusOrMinus = 2 Then
              Hostmask = Param(ONick, PosInONick)
              If Hostmask = "" Then
                SendLine "mode " & Chan & " +b", 1
              Else
                'Check script Hooks
                For ScNum = 1 To ScriptCount
                  If Scripts(ScNum).Hooks.UnBan Then
                    RunScriptX ScNum, "UnBan", Nick, Channels(ChNum).User(UsNum).RegNick, Chan, Hostmask
                  End If
                Next ScNum
                
                KUsNum = 0
                For u2 = 1 To Channels(ChNum).BanCount
                  If LCase(Channels(ChNum).BanList(u2).Mask) = LCase(Hostmask) Then KUsNum = u2: Exit For
                Next u2
                If KUsNum > 0 Then
                  For u2 = UsNum To Channels(ChNum).BanCount - 1
                    Channels(ChNum).BanList(u2) = Channels(ChNum).BanList(u2 + 1)
                  Next u2
                  Channels(ChNum).BanCount = Channels(ChNum).BanCount - 1
                  For u2 = 1 To BanCount
                    If Bans(u2).Channel = "*" Or LCase(Bans(u2).Channel) = LCase(Channels(ChNum).Name) Then
                      If Bans(u2).Sticky And LCase(Bans(u2).Hostmask) = LCase(Hostmask) Then
                        SendLine "mode " & Channels(ChNum).Name & " +b " & Bans(u2).Hostmask, 2
                      End If
                    End If
                  Next u2
                End If
              End If
            End If
          Case "e"
            PosInONick = PosInONick + 1
            'ON EXCEPT
            If PlusOrMinus = 1 Then
              Hostmask = Param(ONick, PosInONick)
              ' possibly cut of by server but set
              If Hostmask = "" Then
                SendLine "mode " & Chan & " +e", 1
              Else
                'Remove Except Orders
                RemOrder "Except " & Chan & " " & Hostmask
                RemDesiredExcept ChNum, Hostmask
                'Add Excepts to Exceptlist
                MatchedOne = False
                For u2 = 1 To Channels(ChNum).ExceptCount
                  If LCase(Channels(ChNum).ExceptList(u2).Mask) = LCase(Hostmask) Then MatchedOne = True
                Next u2
                If Not MatchedOne Then
                  Channels(ChNum).ExceptCount = Channels(ChNum).ExceptCount + 1
                  Channels(ChNum).ExceptList(Channels(ChNum).ExceptCount).Mask = Hostmask
                  Channels(ChNum).ExceptList(Channels(ChNum).ExceptCount).CreatedAt = Now
                End If
              End If
            End If
            'ON UNEXCEPT
            If PlusOrMinus = 2 Then
              Hostmask = Param(ONick, PosInONick)
              If Hostmask = "" Then
                SendLine "mode " & Chan & " +e", 1
              Else
                UsNum = 0
                For u2 = 1 To Channels(ChNum).ExceptCount
                  If LCase(Channels(ChNum).ExceptList(u2).Mask) = LCase(Hostmask) Then UsNum = u2: Exit For
                Next u2
                If UsNum > 0 Then
                  For u2 = UsNum To Channels(ChNum).ExceptCount - 1
                    Channels(ChNum).ExceptList(u2) = Channels(ChNum).ExceptList(u2 + 1)
                  Next u2
                  Channels(ChNum).ExceptCount = Channels(ChNum).ExceptCount - 1
                  For u2 = 1 To ExceptCount
                    If Excepts(u2).Channel = "*" Or LCase(Excepts(u2).Channel) = LCase(Channels(ChNum).Name) Then
                      SendLine "mode " & Channels(ChNum).Name & " +e " & Excepts(u2).Hostmask, 2
                    End If
                  Next u2
                End If
              End If
            End If
          Case "I"
            PosInONick = PosInONick + 1
            'ON INVITE
            If PlusOrMinus = 1 Then
              Hostmask = Param(ONick, PosInONick)
              If Hostmask = "" Then
                SendLine "MODE " & Chan & " +I", 1
              Else
                'Remove Invite Orders
                RemOrder "Invite " & Chan & " " & Hostmask
                RemDesiredInvite ChNum, Hostmask
                'Add Invites to Invitelist
                MatchedOne = False
                For u2 = 1 To Channels(ChNum).InviteCount
                  If LCase(Channels(ChNum).InviteList(u2).Mask) = LCase(Hostmask) Then MatchedOne = True
                Next u2
                If Not MatchedOne Then
                  Channels(ChNum).InviteCount = Channels(ChNum).InviteCount + 1
                  Channels(ChNum).InviteList(Channels(ChNum).InviteCount).Mask = Hostmask
                  Channels(ChNum).InviteList(Channels(ChNum).InviteCount).CreatedAt = Now
                End If
              End If
            End If
            'ON UNBAN
            If PlusOrMinus = 2 Then
              Hostmask = Param(ONick, PosInONick)
              If Hostmask = "" Then
                SendLine "MODE " & Chan & " +I", 1
              Else
                UsNum = 0
                For u2 = 1 To Channels(ChNum).InviteCount
                  If LCase(Channels(ChNum).InviteList(u2).Mask) = LCase(Hostmask) Then UsNum = u2: Exit For
                Next u2
                If UsNum > 0 Then
                  For u2 = UsNum To Channels(ChNum).InviteCount - 1
                    Channels(ChNum).InviteList(u2) = Channels(ChNum).InviteList(u2 + 1)
                  Next u2
                  Channels(ChNum).InviteCount = Channels(ChNum).InviteCount - 1
                  For u2 = 1 To InviteCount
                    If Invites(u2).Channel = "*" Or LCase(Invites(u2).Channel) = LCase(Channels(ChNum).Name) Then
                      SendLine "mode " & Channels(ChNum).Name & " +I " & Invites(u2).Hostmask, 2
                    End If
                  Next u2
                End If
              End If
            End If
          Case "o"
            PosInONick = PosInONick + 1
            'OP
            If PlusOrMinus = 1 Then
              TookOp = False
              KNick = Param(ONick, PosInONick)
              KUsNum = FindUser(KNick, ChNum)
              KUserFlags = GetUserChanFlags(Channels(ChNum).User(KUsNum).RegNick, Chan)
              sNewStatus = "@"
              If InStr(Channels(ChNum).User(KUsNum).Status, "%") > 0 Then sNewStatus = sNewStatus & "%"
              If InStr(Channels(ChNum).User(KUsNum).Status, "+") > 0 Then sNewStatus = sNewStatus & "+"
              Channels(ChNum).User(KUsNum).Status = sNewStatus
              If Param(ONick, PosInONick) = MyNick Then
                If Not Channels(ChNum).GotOPs Then JustGainedOps = True
                Channels(ChNum).GotOPs = True
              Else
                If (Channels(ChNum).GotOPs = True) And (JustGainedOps = False) Then
                  If Nick <> MyNick Then
                    If MatchFlags(KUserFlags, "+d") Then
                      AddMassMode "-o", KNick: TookOp = True
                    ElseIf (Channels(ChNum).User(KUsNum).RegNick = "") And (Channels(ChNum).DeopUnknownUsers = 1) And (IsB = False) And (IsN = False) Then
                      AddMassMode "-o", KNick: TookOp = True
                    ElseIf (Channels(ChNum).User(KUsNum).RegNick = "") And (Channels(ChNum).DeopUnknownUsers = 2) Then
                      AddMassMode "-o", KNick: TookOp = True
                    End If
                  End If
                End If
              End If
              'Check script Hooks
              For ScNum = 1 To ScriptCount
                If Scripts(ScNum).Hooks.Op Then
                  RunScriptX ScNum, "Op", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, KNick, Channels(ChNum).User(KUsNum).RegNick, KUserFlags, IIf(TookOp, "True", "False")
                End If
              Next ScNum
            End If
            'DEOP
            If PlusOrMinus = 2 Then
              ReOpped = False
              KNick = Param(ONick, PosInONick)
              KUsNum = FindUser(KNick, ChNum)
              KUserFlags = GetUserChanFlags(Channels(ChNum).User(KUsNum).RegNick, Chan)
              If Param(ONick, PosInONick) = MyNick Then
                If Channels(ChNum).GotOPs Then JustLostOps = True
                Channels(ChNum).GotOPs = False
                sNewStatus = ""
                If InStr(Channels(ChNum).User(KUsNum).Status, "%") > 0 Then sNewStatus = sNewStatus & "%"
                If InStr(Channels(ChNum).User(KUsNum).Status, "+") > 0 Then sNewStatus = sNewStatus & "+"
                Channels(ChNum).User(KUsNum).Status = sNewStatus
              Else
                If InStr(Channels(ChNum).User(KUsNum).Status, "@") > 0 Then
                  RemOrder "giveop " & Channels(ChNum).Name & " " & Channels(ChNum).User(KUsNum).Nick
                  If InStr(Channels(ChNum).User(KUsNum).Status, "+") > 0 Then Channels(ChNum).User(KUsNum).Status = "+" Else Channels(ChNum).User(KUsNum).Status = ""
                  If MatchFlags(KUserFlags, "+f") And (IsF = False) Then
                    If (Nick <> MyNick) And Channels(ChNum).ProtectFriends Then
                      If Not TookOp Then AddMassMode "-o", Nick
                      AddMassMode "+o", Param(ONick, PosInONick): ReOpped = True
                    End If
                  ElseIf MatchFlags(KUserFlags, "+v") Then
                    If (Channels(ChNum).GotOPs = True) And (InStr(Channels(ChNum).User(KUsNum).Status, "+") = 0) Then SendLine "mode " & Chan & " +v " & Param(ONick, PosInONick), 2
                  End If
                  'Immediately re-op bots deopped by non-owners and non-+o-bots
                  If MatchFlags(KUserFlags, "+bo") Then
                    If Not (IsN Or (IsB And MatchFlags(UserFlags, "+o"))) Then
                      If Nick <> MyNick Then AddMassMode "+o", Param(ONick, PosInONick): ReOpped = True
                    End If
                  End If
                End If
              End If
              'Check script Hooks
              For ScNum = 1 To ScriptCount
                If Scripts(ScNum).Hooks.Deop Then
                  RunScriptX ScNum, "Deop", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, KNick, Channels(ChNum).User(KUsNum).RegNick, KUserFlags, IIf(ReOpped, "True", "False")
                End If
              Next ScNum
            End If
          Case "h"
            PosInONick = PosInONick + 1
            'HELPER
            If PlusOrMinus = 1 Then
              KUsNum = FindUser(Param(ONick, PosInONick), ChNum)
              sNewStatus = ""
              If InStr(Channels(ChNum).User(KUsNum).Status, "@") > 0 Then sNewStatus = sNewStatus & "@"
              sNewStatus = sNewStatus & "%"
              If InStr(Channels(ChNum).User(KUsNum).Status, "+") > 0 Then sNewStatus = sNewStatus & "+"
              Channels(ChNum).User(KUsNum).Status = sNewStatus
              If Param(ONick, PosInONick) = MyNick Then
                If Not Channels(ChNum).GotHOPs Then JustGainedHOps = True
                Channels(ChNum).GotHOPs = True
              End If
            End If
            'DEHELPER
            If PlusOrMinus = 2 Then
              KUsNum = FindUser(Param(ONick, PosInONick), ChNum)
              sNewStatus = ""
              If InStr(Channels(ChNum).User(KUsNum).Status, "@") > 0 Then sNewStatus = sNewStatus & "@"
              If InStr(Channels(ChNum).User(KUsNum).Status, "+") > 0 Then sNewStatus = sNewStatus & "+"
              Channels(ChNum).User(KUsNum).Status = sNewStatus
              If Param(ONick, PosInONick) = MyNick Then
                Channels(ChNum).GotHOPs = False
              End If
            End If
        End Select
      Next u
      
      'Check script Hooks
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Mode Then
          RunScriptX ScNum, "Mode", Nick, Channels(ChNum).User(UsNum).RegNick, Chan, OpString & " " & ONick
        End If
        If Scripts(ScNum).Hooks.ModeEnd Then
          RunScriptX ScNum, "ModeEnd"
        End If
      Next ScNum
      SpreadPartylineChanEvent "mode", Chan, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", OpString & " " & ONick
      
      If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then DoMassMode Chan Else MassModeOps = "": MassModeDeops = "": MassModeBans = "": MassModeUnbans = ""
      
      If Channels(ChNum).CompletedWHO = True Then
        If JustGainedOps = True Then
          DoAutoStuff ChNum
          RemOrder "xgop " & Channels(ChNum).Name
          OfferOps ChNum
        End If
        If JustLostOps = True Then
          RemOrder "gop " & Channels(ChNum).Name
          RequestOps ChNum
        End If
      End If
      
      If ((Channels(ChNum).GotOPs = True) And (IsN = False) And (IsB = False)) Or (JustGainedOps = True) Or (JustGainedHOps = True) Then
        If Channels(ChNum).InFlood = False Then
          ChangMode = ChangeMode(GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), Channels(ChNum).Mode)
          If ChangMode <> "" Then
            If Channels(ChNum).CompletedMode Then SendLine "mode " & Chan & " " & ChangMode, 2
          End If
          If JustGainedOps Or JustGainedHOps Then
            If (LCase(GetChannelSetting(Channels(ChNum).Name, "ProtectTopic", "off")) = "on" And GetChannelSetting(Channels(ChNum).Name, "DefaultTopic", "") <> Channels(ChNum).Topic) Or (GetChannelSetting(Channels(ChNum).Name, "DefaultTopic", "") <> "" And Channels(ChNum).Topic = "") Then
              If Not IsOrdered("topic " & Chan) Then
                SendLine "topic " & Chan & " :" & GetChannelSetting(Channels(ChNum).Name, "DefaultTopic", ""), 2
                Order "topic " & Chan, 20
              End If
            End If
          End If
        End If
      End If
      
      'Trace Channels(ChNum).Mode
      If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then
        ChangMode = ChangeMode(GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), Channels(ChNum).Mode)
        If ChangMode <> "" Then
          SendLine "mode " & Channels(ChNum).Name & " " & ChangMode, 2
        End If
      End If
      
      If (JustGainedOps Or JustGainedHOps) And (Channels(ChNum).CompletedBANS = True) Then CheckBans Channels(ChNum).Name
      If (JustGainedOps Or JustGainedHOps) And (Channels(ChNum).CompletedExcepts = True) Then CheckExcepts Channels(ChNum).Name
      If (JustGainedOps Or JustGainedHOps) And (Channels(ChNum).CompletedInvites = True) Then CheckInvites Channels(ChNum).Name
    End If
  ElseIf Param(Line, 3) = MyNick Then
    'My Mode
    If RestrictCycle = True Then
      If MatchFlags(StripDP(Param(Line, 4)), "+r") Then
        RestrictedIndex = RestrictedIndex + 1
        If RestrictedIndex > 3 Then
          SpreadFlagMessage 0, "+m", "4*** My connection is restricted!!"
          For Index = 1 To BotUserCount
            If MatchFlags(BotUsers(Index).Flags, "+n") Then
              SendNote BotNetNick, BotUsers(Index).Name, "+n", "My connection seems to be permanent restricted! Stoped cycling."
            End If
          Next Index
        Else
          SpreadFlagMessage 0, "+m", "4*** My connection is restricted! Cycling Server..."
          SendTCP ServerSocket, "QUIT Cycling Server..." & vbLf
        End If
      Else
        If RestrictedIndex > 0 Then
          RestrictedIndex = 0
          SpreadFlagMessage 0, "+m", "3*** My connection no longer restricted."
        End If
      End If
    End If
  End If
End Sub

Sub SYS_JOIN(Nick As String, Line As String)
Dim RegUser As String, UserFlags As String, GaveOps As Boolean, u As Long
Dim Chan As String, ChNum As Long, UsNum As Long, ScNum As Long, Rest As String
Dim OUserFlags As String, OUserNum As Long, NewHM As String
Dim CloneCount As Long, KickedUser As Boolean, MatchedOne As Boolean
Dim FullAddress As String
  
  FullAddress = StripDP(Param(Line, 1))
  Chan = StripDP(Param(Line, 3))
  If Right(Chan, 2) = "o" Then Chan = Left(Chan, Len(Chan) - 2)
  'Somebody else joined a channel
  If Nick <> MyNick Then
    ChNum = FindChan(Chan): If ChNum = 0 Then Exit Sub
    
    RegUser = SearchUserFromHostmask(FullAddress)
    UserFlags = GetUserChanFlags2(BotUserNum, Chan)
    OUserNum = GetUserNum(Nick)
    OUserFlags = GetUserChanFlags2(OUserNum, Chan)
    'Add user to the Channels() array
    UsNum = AddChanUser(ChNum, Nick, RegUser, BotUserNum, FullAddress)
    
    If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then
      'Check internal ban list
      KickedUser = CheckPermBans(ChNum, UsNum)
      'If Double-@ (fake) -> kickban user with IP
      If FakeIDKick = True Then
        If Not KickedUser Then
          If InStr(InStr(Channels(ChNum).User(UsNum).Hostmask, "@") + 1, Channels(ChNum).User(UsNum).Hostmask, "@") > 0 Then
            KickedUser = True
            If Not IsValidIP(Mask(Channels(ChNum).User(UsNum).Hostmask, 11)) Then
              If Channels(ChNum).User(UsNum).IPmask <> "" Then
                If AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).IPmask, 2)) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).IPmask, 2), UnbanTime
                AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).IPmask, "Fake Ident"
              End If
            Else
              If AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, 2)) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), UnbanTime
              AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, "Fake Ident"
            End If
          End If
        End If
      End If
      'Bogus kickban
      If FakeIDKick = True Then
        If Not KickedUser Then
          Rest = Mask(Channels(ChNum).User(UsNum).Hostmask, 12)
          If InStr(Rest, "") > 0 Or InStr(Rest, "") > 0 Or InStr(Rest, "") > 0 Then
            Rest = Mask(Channels(ChNum).User(UsNum).Hostmask, 2)
            FloodCheck ChNum, "+smtin", 4
            If AddToBan(ChNum, Rest) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Rest, UnbanTime
            AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, "Bogus Ident"
            KickedUser = True
          End If
        End If
      End If
      'Nick Protection
      If Not KickedUser Then
        If MatchFlags(OUserFlags, "+x") And (Channels(ChNum).User(UsNum).Nick <> MyNick) Then
          MatchedOne = False
          For u = 1 To BotUsers(OUserNum).HostMaskCount
            If MatchHost(BotUsers(OUserNum).HostMasks(u), FullAddress) Then MatchedOne = True: Exit For
          Next u
          If Not MatchedOne Then
            If MatchFlags(UserFlags, "+n") Then
              SendLine "privmsg " & Channels(ChNum).Name & " :" & Channels(ChNum).User(UsNum).Nick & ": Please choose another nick, the one you are using is protected.", 3
            Else
              If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 Then SendLine "mode " & Channels(ChNum).Name & " -o+b " & Nick & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 8), 2
              If InStr(Channels(ChNum).User(UsNum).Status, "@") = 0 Then SendLine "mode " & Channels(ChNum).Name & " +b " & Mask(Channels(ChNum).User(UsNum).Hostmask, 8), 2
              SendLine "kick " & Channels(ChNum).Name & " " & Channels(ChNum).User(UsNum).Nick & " :Nick faker! Change your nick!", 2
              TimedEvent "notice " & Channels(ChNum).User(UsNum).Nick & " :Please use another nick, the one you are using is protected.", 4
              TimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 8), 120
            End If
            KickedUser = True
          End If
        End If
      End If
      'Clone Kick
      If Not KickedUser Then
        If Channels(ChNum).CloneKick Then
          NewHM = Mask(Channels(ChNum).User(UsNum).Hostmask, 2)
          CloneCount = 0
          For u = 1 To Channels(ChNum).UserCount
            If LCase(Mask(Channels(ChNum).User(u).Hostmask, 2)) = LCase(NewHM) Then If MatchFlags(GetUserChanFlags(Channels(ChNum).User(u).RegNick, Channels(ChNum).Name), "-b") Then CloneCount = CloneCount + 1
          Next u
          If CloneCount = 3 Then
            PatternKickBan Nick, Channels(ChNum).Name, "Clones from " & NewHM, True, 120, 300
            KickedUser = True
          End If
        End If
      End If
      'AutoKick (+k flag)
      If Not KickedUser Then
        If MatchFlags(UserFlags, "+k") Then
          Rest = SearchMatchedHostmask(FullAddress)
          If AddToBan(ChNum, Rest) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Rest, UnbanTime
          AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, "Banned: " & GetUserData(Channels(ChNum).User(UsNum).UserNum, "comment", "requested")
          KickedUser = True
        Else
          'AutoOp (+a flag)
          If MatchFlags(UserFlags, "+a") Then
            TimedEvent "Op " & Chan & " " & Nick, 1
            GaveOps = True
          End If
          'AutoVoice (+v flag / channel is +m)
          If GaveOps = False Then
            If Channels(ChNum).InFlood = False Then
              If MatchFlags(UserFlags, "+v") Then
                SendLine "mode " & Chan & " +v " & Nick, 2
              ElseIf InStr(Param(Channels(ChNum).Mode, 1), "m") > 0 Then
                If (Channels(ChNum).AutoVoice = 1) Or ((Channels(ChNum).AutoVoice = 2) And (RegUser <> "")) Then
                  SendLine "mode " & Chan & " +v " & Nick, 2
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    
    'Newbie Greeting
    If (Channels(ChNum).NewbieGreeting <> "") And (Channels(ChNum).InFlood = False) And (KickedUser = False) Then
      If MatchFlags(UserFlags, "-b") Then
        Rest = "greet " & Channels(ChNum).Name & " " & Mask(FullAddress, 2)
        If IsOrdered(Rest) = False Then
          If GetPPString(Channels(ChNum).Name, Mask(Channels(ChNum).User(UsNum).Hostmask, 3), "", HomeDir & "Newbie.ini") = "" Then
            WritePPString Channels(ChNum).Name, Mask(Channels(ChNum).User(UsNum).Hostmask, 3), "!", HomeDir & "Newbie.ini"
            SendLine "notice " & Nick & " :" & Channels(ChNum).NewbieGreeting, 3
            Order Rest, 10
          End If
        End If
      End If
    End If
    
    'AUTH Service
      If AuthTarget <> "" And LCase(Nick) = LCase(ParamX(ParamX(AuthTarget, "!", 1), "@", 1)) And AuthJust = False And AuthReAuth = True Then
        SendLine "PRIVMSG " & AuthTarget & " :" & AuthCommand & " " & AuthParam1 & IIf(AuthParam2 <> "", " " & AuthParam2, ""), 1
        SpreadFlagMessage 0, "+s", "14*** Trying to AUTH."
        AuthJust = True
      End If
    
    'Check script Hooks
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Join Then
        RunScriptX ScNum, "Join", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags
      End If
    Next ScNum
    
    ' Generate Partyline Message
    SpreadPartylineChanEvent "join", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", ""

    If RegUser <> "" Then
      'LastSeen
      WriteSeenEntry RegUser, "", Now, Channels(ChNum).Name, "*join*", Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      
      'Notes
      If Not IsIgnored(FullAddress) And Not KickedUser Then
        AddIgnore Mask(FullAddress, 2), 10, 1
        u = NotesCount(RegUser)
        If BotUsers(GetUserNum(RegUser)).Password <> "" Then
          If u > 0 Then
            If u = 1 Then
              SendLine "notice " & Nick & " :Hi! I've got 1 note waiting for you.", 3
              SendLine "notice " & Nick & " :To get it, type: /msg " & MyNick & " notes <pass>", 3
            Else
              SendLine "notice " & Nick & " :Hi! I've got " & CStr(u) & " notes waiting for you.", 3
              SendLine "notice " & Nick & " :To get them, type: /msg " & MyNick & " notes <pass>", 3
            End If
          End If
        Else
          If u > 0 Then
            SpreadFlagMessage 0, "+m", "14[" & Time & "] *** I told " & IIf(Nick <> RegUser, Nick & " (" & RegUser & ")", Nick) & " to set a password: Waiting notes!"
            SendLine "notice " & Nick & " :Hi! I've got notes for you, but you don't have a password set.", 3
            SendLine "notice " & Nick & " :To set one, type: /msg " & MyNick & " pass <your password>", 3
          End If
        End If
      End If
    End If
    If (LCase(RegUser) <> LCase(Nick)) And (Not KickedUser) And (Not Channels(ChNum).InFlood) Then
      WriteExtSeenEntry Nick, RegUser, Now, Channels(ChNum).Name, "*join*", Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
    End If
  Else
    'I joined a channel - update Channels() array
    AddChannel Chan
    
    'Check script Hooks
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Join Then
        RunScriptX ScNum, "Join", Chan, Nick, "", ""
        SpreadPartylineChanEvent "join", Chan, Nick, BotNetNick, "", "", "", "", ""
      End If
    Next ScNum
  End If
End Sub

Sub SYS_PART(Nick As String, Line As String)
  Dim Chan As String, ChNum As Long, UsNum As Long, UserFlags As String, ScNum As Long, RegUser As String
  Chan = Param(Line, 3)
  ChNum = FindChan(Chan)
  UsNum = FindUser(Nick, ChNum)
  If ChNum <> 0 And UsNum <> 0 Then
    UpdateBufferKicks Nick, Channels(ChNum).Name
    UserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Chan)
    
    'Check script Hooks
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Part Then
        RunScriptX ScNum, "Part", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags
      End If
    Next ScNum
    SpreadPartylineChanEvent "part", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", ""
    
    If Nick <> MyNick Then
      RegUser = Channels(ChNum).User(UsNum).RegNick
      
      'Set seen entry
      If RegUser <> "" Then
        WriteSeenEntry RegUser, "", Now, Channels(ChNum).Name, "*left*", Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      End If
      
      'Set extseen Entry
      If LCase(RegUser) <> LCase(Nick) And Channels(ChNum).InFlood = False Then
        WriteExtSeenEntry Nick, RegUser, Now, Channels(ChNum).Name, "*left*", Mask(Channels(ChNum).User(UsNum).Hostmask, 10)
      End If
      
      RemChanUser Nick, ChNum
      RemKickUser ChNum, Nick
      
      If Channels(ChNum).UserCount = 1 And Channels(ChNum).GotOPs = False Then SendLine "part " & Channels(ChNum).Name + vbCrLf & "join " & Channels(ChNum).Name, 1: SpreadFlagMessage 0, "+m", "3*** Rejoining " & Channels(ChNum).Name & " to get ops..."
    Else
      SetPermChanStat Chan, ChanStat_Left
      SpreadFlagMessage 0, "+m", "10*** I left the channel " & Chan & "."
      RemChan Chan
    End If
  End If
End Sub

Sub SYS_TOPIC(Nick As String, Line As String)
  Dim Chan As String, ChNum As Long, UsNum As Long, UserFlags As String, ScNum As Long
  
  Chan = Param(Line, 3)
  ChNum = FindChan(Chan)
  UsNum = FindUser(Nick, ChNum)
  Channels(ChNum).Topic = StripDP(Trim(Right(Line, Len(Line) - Len(Param(Line, 1)) - Len(Param(Line, 2)) - Len(Param(Line, 3)) - 3)))
  UserFlags = GetUserChanFlags(SearchUserFromHostmask(Channels(ChNum).User(UsNum).Hostmask), Chan)
  
  'Check script Hooks
  HaltDefault = False
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.Topic Then
      RunScriptX ScNum, "Topic", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, Channels(ChNum).Topic
    End If
  Next ScNum
  
  'Topic protection
  If HaltDefault = False Then
    If MatchFlags(UserFlags, "-bn") Then
      If LCase(GetChannelSetting(Channels(ChNum).Name, "ProtectTopic", "off")) = "on" Then
        If GetChannelSetting(Channels(ChNum).Name, "DefaultTopic", "") <> Channels(ChNum).Topic Then
          If Not IsOrdered("topic " & Chan) Then
            SendLine "topic " & Chan & " :" & GetChannelSetting(Channels(ChNum).Name, "DefaultTopic", ""), 2
            Order "topic " & Chan, 10
          End If
        End If
      End If
    End If
  End If
End Sub

Sub SYS_NOTICE(Nick As String, Line As String)
  Dim Rest As String, ScNum As Long, Chan As String, RegUser As String, ChNum As Long, UsNum As Long, UserFlags As String
  Rest = StripDP(GetRest(Line, 4))
  
  'SERVER NOTICES
  If InStr(1, Param(Line, 1), "!", vbBinaryCompare) = 0 Then
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Server_notice Then
        RunScriptX ScNum, "Server_notice", StripDP(Param(Line, 1)), Param(Line, 3), StripDP(GetRest(Line, 4))
      End If
    Next ScNum
    Exit Sub
  End If
  
  If LCase(Param(Line, 3)) = LCase(MyNick) Then
    If Left(Rest, 1) = "" Then
      'CTCP Reply
      If Right(Rest, 1) = "" Then Rest = Mid(Rest, 2, Len(Rest) - 2) Else Rest = Mid(Rest, 2, Len(Rest) - 1)
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Priv_ctcpreply Then
          RunScriptX ScNum, "Priv_ctcpreply", Nick, RegUser, Rest
        End If
      Next ScNum
    Else
      'Notice to Bot
      If Nick = "" Then
        SpreadFlagMessage 0, "+m", "14[" & Time & "] NOTICE from " & StripDP(Param(Line, 1)) & ": " & Rest
      Else
        SpreadFlagMessage 0, "+m", "14[" & Time & "] NOTICE from " & Nick & ": " & Rest
      End If
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Priv_notice Then
          RunScriptX ScNum, "Priv_notice", Nick, RegUser, Rest
        End If
      Next ScNum
    End If
  ElseIf UCase(Param(Line, 3)) = "&SERVERS" Then
    'Split Detection
    If UCase(Param(Line, 5)) = "SQUIT" Then AddSplitServer Param(Line, 6)
    If UCase(Param(Line, 5)) = "SERVER" Then RemoveSplitServer Param(Line, 6)
  Else
    ' Channel CTCP
    If IsIgnoredHost(StripDP(Param(Line, 1))) Then Exit Sub
    Chan = StripDP(Param(Line, 3))
    RegUser = SearchUserFromHostmask(StripDP(Param(Line, 1)))
    ChNum = FindChan(Chan)
    Rest = StripDP(GetRest(Line, 4))
    If Channels(ChNum).ColorKick Then
      If InStr(Rest, "") > 0 Then
        UsNum = FindUser(Nick, ChNum)
        UserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Chan)
        If MatchFlags(UserFlags, "-bn") Then
          If Not IsIgnored(StripDP(Param(Line, 1))) Then
            AddIgnore Mask(StripDP(Param(Line, 1)), 2), 10, 1
            SendLine "notice " & Nick & " :No colors please!", 3
          Else
            SendLine "kick " & Chan & " " & Nick & " :No colors please!!!", 2
          End If
        End If
      End If
    End If
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Chan_notice Then
        RunScriptX ScNum, "Chan_notice", Nick, RegUser, Chan, Rest
      End If
    Next ScNum
  End If
End Sub

Sub SYS_PRIVMSG(Nick As String, Line As String)
  Dim Rest As String, Chan As String, ScNum As Long, ChNum As Long, UsNum As Long, RegUser As String, UserFlags As String, MatchedOne As Boolean, KUsNum As Long, Hostmask As String
  
  Rest = StripDP(Param(Line, 4))
  
  'PRIVATE MESSAGES
  If LCase(Param(Line, 3)) = LCase(MyNick) Then
    HandlePrivateMessage Line
  ElseIf Left(Rest, 1) = "" Then
    'CHANNEL CTCP EVENTS
    Chan = Param(Line, 3)
    ChNum = FindChan(Chan)
    UsNum = FindUser(Nick, ChNum)
    If UsNum > 0 Then
      If (InStr(Rest, "PING") > 0) Or (InStr(Rest, "FINGER") > 0) Or (InStr(Rest, "TIME") > 0) Or (InStr(Rest, "VERSION") > 0) Or (InStr(Rest, "CLIENTINFO") > 0) Or (InStr(Rest, "USERINFO") > 0) Then
        Channels(ChNum).User(UsNum).CTCPs = Channels(ChNum).User(UsNum).CTCPs + 1
        FloodCheck ChNum, "+smtinC", 5
      End If
      If (Channels(ChNum).User(UsNum).CTCPs = 3) Or ((Channels(ChNum).InFlood = True) And (Channels(ChNum).User(UsNum).CTCPs = 2)) Then
        If AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, 2)) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), UnbanTime
        AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, "CTCP Flood"
        SpreadPartylineChanEvent "ctcpflood", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", ""
      End If
    End If
    RegUser = SearchUserFromHostmask(StripDP(Param(Line, 1)))
    If Right(Rest, 1) = "" And Len(Rest) > 1 Then Rest = Mid(Rest, 2, Len(Rest) - 2) Else Rest = Mid(Rest, 2, Len(Rest) - 1)
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Chan_ctcp Then
        RunScriptX ScNum, "Chan_ctcp", Nick, RegUser, Chan, Rest
      End If
    Next ScNum
  Else
    If IsIgnoredHost(StripDP(Param(Line, 1))) Then Exit Sub
    Chan = StripDP(Param(Line, 3))
    ChNum = FindChan(Chan)
    UsNum = FindUser(Nick, ChNum)
    If UsNum > 0 Then
      Rest = StripDP(GetRest(Line, 4))
      Channels(ChNum).User(UsNum).LastEvent = WinTickCount
      HaltDefault = False
      'Check script Hooks
      If LCase(Param(Rest, 1)) <> "action" Then
        For ScNum = 1 To ScriptCount
          If Scripts(ScNum).Hooks.Chan_msg Then
            RunScriptX ScNum, "Chan_msg", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, Rest
          End If
        Next ScNum
        SpreadPartylineChanEvent "chantalk", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", Rest
      Else
        Rest = GetRest(Rest, 2): Rest = Left(Rest, Len(Rest) - 1)
        For ScNum = 1 To ScriptCount
          If Scripts(ScNum).Hooks.Chan_act Then
            RunScriptX ScNum, "Chan_act", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, Rest
          End If
        Next ScNum
        SpreadPartylineChanEvent "chanaction", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", Rest
      End If
      
      'Text flood protection
      If Channels(ChNum).GotOPs Or (Channels(ChNum).GotHOPs And InStr(1, Channels(ChNum).User(UsNum).Status, "@", vbBinaryCompare) = 0) Then
        UserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Chan)
        If MatchFlags(UserFlags, "-bn") Then
          Channels(ChNum).User(UsNum).LineCount = Channels(ChNum).User(UsNum).LineCount + 1
          Channels(ChNum).User(UsNum).CharCount = Channels(ChNum).User(UsNum).CharCount + Len(Rest)
          If Channels(ChNum).User(UsNum).LastLine = Rest Then
            Channels(ChNum).User(UsNum).RepeatCount = Channels(ChNum).User(UsNum).RepeatCount + 1
            TimedEvent "RemRepeat " & Chan & " " & Nick, 10
            If (Channels(ChNum).MaxRepeats > 0) And (Channels(ChNum).User(UsNum).RepeatCount > Channels(ChNum).MaxRepeats) Then
              FloodCheck ChNum, "+smtin", 4
              If Channels(ChNum).InFlood = False Then
                MatchedOne = AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, Channels(ChNum).BanMask))
                If MatchedOne = True Then TimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, Channels(ChNum).BanMask), 30
                AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, IIf(MatchedOne, Channels(ChNum).User(UsNum).Hostmask, ""), "Repeating sucks!"
              Else
                MatchedOne = AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, 2))
                If MatchedOne = True Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), UnbanTime
                AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, IIf(MatchedOne, Channels(ChNum).User(UsNum).Hostmask, ""), "Repeating sucks!"
              End If
              SpreadPartylineChanEvent "textflood", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", ""
            End If
          Else
            Channels(ChNum).User(UsNum).LastLine = Rest
          End If
          TimedEvent "RemLine " & Chan & " " & Nick, 10
          TimedEvent "RemChars " & Chan & " " & Nick & " " & Trim(Str(Len(Rest))), 10
        End If
        If ((Channels(ChNum).MaxLines > 0) And (Channels(ChNum).User(UsNum).LineCount > Channels(ChNum).MaxLines)) Or ((Channels(ChNum).MaxChars > 0) And (Channels(ChNum).User(UsNum).CharCount > Channels(ChNum).MaxChars)) Then
          If MatchFlags(UserFlags, "-bn") Then
            FloodCheck ChNum, "+smtin", 4
            If Channels(ChNum).InFlood = False Then
              MatchedOne = AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, 3))
              If MatchedOne = True Then TimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 3), 30
              AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, IIf(MatchedOne, Channels(ChNum).User(UsNum).Hostmask, ""), "Text Flood"
            Else
              MatchedOne = AddToBan(ChNum, Mask(Channels(ChNum).User(UsNum).Hostmask, 2))
              If MatchedOne = True Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), UnbanTime
              AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, IIf(MatchedOne, Channels(ChNum).User(UsNum).Hostmask, ""), "Text Flood"
            End If
          End If
          SpreadPartylineChanEvent "textflood", Channels(ChNum).Name, Nick, Channels(ChNum).User(UsNum).RegNick, UserFlags, "", "", "", ""
        End If
      End If
      
      'Color kick
      If (Channels(ChNum).ColorKick = True) And (Channels(ChNum).GotOPs Or (Channels(ChNum).GotHOPs And InStr(1, Channels(ChNum).User(UsNum).Status, "@", vbBinaryCompare) = 0)) Then
        If InStr(Rest, "") > 0 Then
          UserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Chan)
          If MatchFlags(UserFlags, "-bn") Then
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 10, 1
              SendLine "notice " & Nick & " :No colors please!", 3
            Else
              SendLine "kick " & Chan & " " & Nick & " :No colors please!!!", 2
            End If
          End If
        End If
      End If
      
      If Not HaltDefault Then
        Select Case StripDP(LCase(Param(Line, 4)))
          Case CommandPrefix & "t"
            If Not Channels(ChNum).AllowVoiceControl Then Exit Sub
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(StripDP(Param(Line, 1)))
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel StripDP(Param(Line, 1)), UsNum + 1
            End If
            KUsNum = FindUser(Nick, ChNum)
            If KUsNum = 0 Then Exit Sub
            If (InStr(Channels(ChNum).User(KUsNum).Status, "+") = 0) And (InStr(Channels(ChNum).User(KUsNum).Status, "@") = 0) Then Exit Sub
            If LCase(GetChannelSetting(Channels(ChNum).Name, "ProtectTopic", "off")) = "on" Then
              If GetIgnoreLevel(StripDP(Param(Line, 1))) <= 1 Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, the topic here is protected.", 3
            Else
              SendLine "topic " & Chan & " :" & GetRest(Line, 5), 2
            End If
          Case CommandPrefix & "k"
            If Not Channels(ChNum).AllowVoiceControl Then Exit Sub
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(StripDP(Param(Line, 1)))
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel StripDP(Param(Line, 1)), UsNum + 1
            End If
            If Len(Param(Line, 5)) <= ServerNickLen And IsValidNick(Param(Line, 5)) Then
              KUsNum = FindUser(Nick, ChNum)
              If Channels(ChNum).User(KUsNum).Status = "" Then Exit Sub
              UsNum = FindUser(Param(Line, 5), ChNum)
              If UsNum = 0 Then Exit Sub
              If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 And InStr(Channels(ChNum).User(KUsNum).Status, "@") = 0 Then Exit Sub
              If LCase(Channels(ChNum).User(UsNum).Nick) = LCase(MyNick) Then SendLine "privmsg " & Chan & " :" & Nick & ": Ha ha.", 3: Exit Sub
              Rest = GetUserChanFlags(Channels(ChNum).User(KUsNum).RegNick, Chan)
              If MatchFlags(Rest, "-n") Then
                Rest = GetUserChanFlags(Channels(ChNum).User(UsNum).RegNick, Chan)
                If MatchFlags(Rest, "+n") Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, " & Channels(ChNum).User(UsNum).Nick & " is an owner.", 3: Exit Sub
                If MatchFlags(Rest, "+b") Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a bot.", 3: Exit Sub
              Else
                Rest = GetUserChanFlags(Channels(ChNum).User(UsNum).RegNick, Chan)
                If MatchFlags(Rest, "+b") Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a bot.", 3: Exit Sub
              End If
              If Param(Line, 6) <> "" Then
                Rest = Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3) & " " & Param(Line, 4) & " " & Param(Line, 5)) - 1)
                SendLine "kick " & Chan & " " & Param(Line, 5) & " :" & Rest, 2
              Else
                Rest = FunnyKick(Channels(ChNum).User(UsNum).Nick)
                SendLine "kick " & Chan & " " & Param(Line, 5) & " :" & Rest, 2
              End If
            End If
          Case CommandPrefix & "kb"
            If Not Channels(ChNum).AllowVoiceControl Then Exit Sub
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(StripDP(Param(Line, 1)))
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel StripDP(Param(Line, 1)), UsNum + 1
            End If
            If Len(Param(Line, 5)) <= ServerNickLen And IsValidNick(Param(Line, 5)) Then
              KUsNum = FindUser(Nick, ChNum)
              If Channels(ChNum).User(KUsNum).Status = "" Then Exit Sub
              UsNum = FindUser(Param(Line, 5), ChNum)
              If UsNum = 0 Then Exit Sub
              If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 And InStr(Channels(ChNum).User(KUsNum).Status, "@") = 0 Then Exit Sub
              If LCase(Channels(ChNum).User(UsNum).Nick) = LCase(MyNick) Then SendLine "privmsg " & Chan & " :" & Nick & ": Are you crazy? :)", 3: Exit Sub
              Rest = GetUserChanFlags(Channels(ChNum).User(KUsNum).RegNick, Chan)
              If MatchFlags(Rest, "-n") Then
                Rest = GetUserChanFlags(Channels(ChNum).User(UsNum).RegNick, Chan)
                If MatchFlags(Rest, "+n") Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, " & Channels(ChNum).User(UsNum).Nick & " is an owner.", 3: Exit Sub
                If MatchFlags(Rest, "+b") Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a bot.", 3: Exit Sub
              Else
                Rest = GetUserChanFlags(Channels(ChNum).User(UsNum).RegNick, Chan)
                If MatchFlags(Rest, "+b") Then SendLine "privmsg " & Chan & " :" & Nick & ": Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a bot.", 3: Exit Sub
              End If
              If Param(Line, 6) <> "" Then
                Rest = Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3) & " " & Param(Line, 4) & " " & Param(Line, 5)) - 1)
                If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 Then SendLine "mode " & Chan & " -o+b " & Channels(ChNum).User(UsNum).Nick & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, Channels(ChNum).BanMask), 2 Else SendLine "mode " & Chan & " +b " & Mask(Channels(ChNum).User(UsNum).Hostmask, Channels(ChNum).BanMask), 2
                SendLine "kick " & Chan & " " & Param(Line, 5) & " :Banned: " & Rest, 2
              Else
                Rest = FunnyKick(Channels(ChNum).User(UsNum).Nick)
                If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 Then SendLine "mode " & Chan & " -o+b " & Channels(ChNum).User(UsNum).Nick & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, Channels(ChNum).BanMask), 2 Else SendLine "mode " & Chan & " +b " & Mask(Channels(ChNum).User(UsNum).Hostmask, Channels(ChNum).BanMask), 2
                SendLine "kick " & Chan & " " & Param(Line, 5) & " :" & Rest, 2
              End If
            End If
          Case CommandPrefix & "seen"
            If Channels(ChNum).ReactToSeen = 0 Then Exit Sub
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(StripDP(Param(Line, 1)))
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel StripDP(Param(Line, 1)), UsNum + 1
            End If
            HaltDefault = False
            For ScNum = 1 To ScriptCount
              If Scripts(ScNum).Hooks.seen Then
                RunScriptX ScNum, "seen", Nick, Chan, Replace(GetRest(Line, 5), ", ", ",")
              End If
            Next ScNum
            If HaltDefault = True Then Exit Sub
            UsNum = FindUser(Nick, ChNum)
            Rest = LastSeen(Param(Replace(GetRest(Line, 5), ", ", ","), 1), Nick, Chan, Channels(ChNum).User(UsNum).RegNick, MatchedOne)
            If Rest <> "" Then
              If ((MatchedOne = False) And (LastSeenOutput = "°NULL " & LCase(Param(Line, 5)))) Or (Rest = LastSeenOutput) Then
                If Not DontSeen Then
                  Rest = ParamX(MakeMsg(IRC_NoRepeats, Nick), "|", Int(Rnd * ParamXCount(MakeMsg(IRC_NoRepeats, Nick), "|")) + 1)
                  Select Case Channels(ChNum).ReactToSeen
                    Case 1: SendLine "privmsg " & Chan & " :" & Rest, 3
                    Case 2: SendLine "notice " & Chan & " :" & Rest, 3
                    Case 3: SendLine "privmsg " & Nick & " :" & Rest, 3
                    Case 4: SendLine "notice " & Nick & " :" & Rest, 3
                  End Select
                End If
                DontSeen = True
              Else
                SpreadFlagMessage 0, "+m", "14[" & Time & "] " & Channels(ChNum).Name & ": <" & Nick & "> " & CommandPrefix & "seen " & Param(Line, 5) + IIf(Not MatchedOne, "  (not seen)", "")
                Select Case Channels(ChNum).ReactToSeen
                  Case 1: SendLine "privmsg " & Chan & " :" & Rest, 3
                  Case 2: SendLine "notice " & Chan & " :" & Rest, 3
                  Case 3: SendLine "privmsg " & Nick & " :" & Rest, 3
                  Case 4: SendLine "notice " & Nick & " :" & Rest, 3
                End Select
              End If
            Else
              If Not DontSeen And Len(Param(Line, 5)) < 20 Then
                If Param(Line, 5) = "" Then Rest = MakeMsg(IRC_Seen_NoNick, Nick) Else Rest = MakeMsg(IRC_Seen_ErrNick, Nick)
                Select Case Channels(ChNum).ReactToSeen
                  Case 1: SendLine "privmsg " & Chan & " :" & Rest, 3
                  Case 2: SendLine "notice " & Chan & " :" & Rest, 3
                  Case 3: SendLine "privmsg " & Nick & " :" & Rest, 3
                  Case 4: SendLine "notice " & Nick & " :" & Rest, 3
                End Select
              End If
              DontSeen = True
            End If
            If MatchedOne Then LastSeenOutput = Rest Else LastSeenOutput = "°NULL " & LCase(Param(Line, 5))
          Case CommandPrefix & "whois"
            If Not Channels(ChNum).ReactToWhois Then Exit Sub
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(StripDP(Param(Line, 1)))
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel StripDP(Param(Line, 1)), UsNum + 1
            End If
            If Left(Param(Line, 5), 1) = "@" Then RegUser = Right(Param(Line, 5), Len(Param(Line, 5)) - 1) Else RegUser = Param(Line, 5)
            If Len(RegUser) <= ServerNickLen And IsValidNick(RegUser) Then
              UsNum = GetUserNum(RegUser)
              If UsNum = 0 Then
                If LCase(RegUser) = "me" Then
                  RegUser = Channels(ChNum).User(FindUser(Nick, ChNum)).RegNick: If RegUser <> "" Then UsNum = GetUserNum(RegUser)
                Else
                  RegUser = Channels(ChNum).User(FindUser(RegUser, ChNum)).RegNick: If RegUser <> "" Then UsNum = GetUserNum(RegUser)
                End If
                If UsNum = 0 Then
                  If LCase(RegUser) = "me" Then
                    SendLine "privmsg " & Chan & " :" & MakeMsg(IRC_Whois_YUnknown, Nick), 3
                  Else
                    If Left(Param(Line, 5), 1) = "@" Then RegUser = Right(Param(Line, 5), Len(Param(Line, 5)) - 1) Else RegUser = Param(Line, 5)
                    SendLine "privmsg " & Chan & " :" & MakeMsg(IRC_Whois_UUnknown, Nick, RegUser), 3
                  End If
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] " & Channels(ChNum).Name & ": <" & Nick & "> !whois " & Param(Line, 5) & "  (unknown user)"
                  Exit Sub
                End If
              End If
              Rest = GetUserData(UsNum, "info", "")
              If Rest = "" Then
                SendLine "privmsg " & Chan & " :" & MakeMsg(IRC_Whois_NoInfo, Nick, BotUsers(UsNum).Name), 3
                SpreadFlagMessage 0, "+m", "14[" & Time & "] " & Channels(ChNum).Name & ": <" & Nick & "> !whois " & Param(Line, 5) & "  (no info set)"
                Exit Sub
              End If
              If Rest = LastWhoisOutput Then
                If Not DontSeen Then
                  Rest = ParamX(MakeMsg(IRC_NoRepeats, Nick), "|", Int(Rnd * ParamXCount(MakeMsg(IRC_NoRepeats, Nick), "|")) + 1)
                  Select Case Int(Rnd * 3) + 1
                    Case 1: SendLine "privmsg " & Chan & " :" & Rest, 3
                    Case 2: SendLine "privmsg " & Chan & " :" & Rest, 3
                    Case 3: SendLine "privmsg " & Chan & " :" & Rest, 3
                  End Select
                End If
                DontSeen = True
              Else
                SpreadFlagMessage 0, "+m", "14[" & Time & "] " & Channels(ChNum).Name & ": <" & Nick & "> !whois " & Param(Line, 5)
                SendLine "privmsg " & Chan & " :" & BotUsers(UsNum).Name & ": " & Rest, 3
                LastWhoisOutput = Rest
              End If
            Else
              If RegUser <> "" Then SendLine "privmsg " & Chan & " :" & Nick & ": This user doesn't exist.", 3
            End If
          Case CommandPrefix & "whatis"
            If Not Channels(ChNum).ReactToWhatis Then Exit Sub
            If Not IsIgnored(StripDP(Param(Line, 1))) Then
              AddIgnore Mask(StripDP(Param(Line, 1)), 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(StripDP(Param(Line, 1)))
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel StripDP(Param(Line, 1)), UsNum + 1
            End If
            Rest = Trim(Replace(GetRest(Line, 5), """", " "))
            Hostmask = ReadWhatis(Rest, False, Nick)  '"Hostmask" contains whatis answer
            If ((Hostmask <> "") And (Hostmask = LastWhatisOutput)) Or ((Hostmask = "") And (Rest = LastWhatisOutput)) Then
              If Not DontSeen Then
                Rest = ParamX(MakeMsg(IRC_NoRepeats, Nick), "|", Int(Rnd * ParamXCount(MakeMsg(IRC_NoRepeats, Nick), "|")) + 1)
                Select Case Int(Rnd * 3) + 1
                  Case 1: SendLine "privmsg " & Chan & " :" & Rest, 3
                  Case 2: SendLine "privmsg " & Chan & " :" & Rest, 3
                  Case 3: SendLine "privmsg " & Chan & " :" & Rest, 3
                End Select
              End If
              DontSeen = True
            Else
              If Hostmask = "" Then
                If Len(Rest) > 26 Then Rest = Left(Rest, 23) & "..." 'Flood protection -> cut after 26 chars
                SendLine "privmsg " & Chan & " :" & MakeMsg(IRC_Whatis_NotFound, Nick, Rest), 3
                SpreadFlagMessage 0, "+m", "14[" & Time & "] " & Channels(ChNum).Name & ": <" & Nick & "> " & CommandPrefix & "whatis " & Rest & "  (not found)"
                LastWhatisOutput = Rest
              Else
                SendLine "privmsg " & Chan & " :" & Hostmask, 3
                SpreadFlagMessage 0, "+m", "14[" & Time & "] " & Channels(ChNum).Name & ": <" & Nick & "> " & CommandPrefix & "whatis " & Rest
                LastWhatisOutput = Hostmask
              End If
            End If
        End Select
      End If
    End If
  End If
End Sub

Sub RPL_WELCOME(Line As String)
  Dim Rest As String
  If InStr(LCase(Param(Line, 1)), "psybnc") = 0 Then
    If ServerName = "" Then
      ServerName = Param(Line, 1)
      Output "*** Server Name: " & StripDP(ServerName) + vbCrLf
    End If
  End If
  MyNick = Param(Line, 3)
  Rest = Param(Line, ParamCount(Line))
  If (InStr(Rest, "@") = 0) Or (InStr(Rest, "!") = 0) Then
    'look up hostmask via userhost
    MyHostmask = ""
    SendLine "userhost " & MyNick, 1
  Else
    GotMyHost Rest
  End If
  
  'Shut down ident Socket (no longer needed)
  CloseIdentSockets
  
  'Update TrayIcon
  SetTrayIcon SI_Online
  
  'Reset Counter
  ConnectTryCounter = 0
  
  'Set myself invisible
  SendLine "mode " & MyNick & " +i", 1
  
  'AUTH if enabled
    If AuthTarget <> "" And AuthCommand <> "" And AuthParam1 <> "" Then
      SendLine "PRIVMSG " & AuthTarget & " :" & AuthCommand & " " & AuthParam1 & IIf(AuthParam2 <> "", " " & AuthParam2, ""), 1
      SpreadFlagMessage 0, "+s", "14*** Trying to AUTH."
      AuthJust = True
    End If
End Sub

Sub RPL_MYINFO(Line As String)
  Dim Rest As String, Chan As String, KeyLine As String, JoinLine As String, u As Long
  
  Rest = LCase(Param(Line, 5))
  If InStr(Rest, "euirc") > 0 Then
    ServerInfo.Network = "euIRC"
    ServerInfo.HidesHosts = True
    ServerInfo.SupportsMultiChanJoin = True
    ServerInfo.SupportsMultiChanWho = True
    ServerInfo.SupportsMultiChanMode = False
    ServerSplitDetection = False
    ServerInfo.SupportsServersChan = False
    ServerInfo.MaxNickLength = 9
  ElseIf InStr(Rest, "dal") > 0 Then
    ServerInfo.Network = "DALnet"
    ServerInfo.HidesHosts = False
    ServerInfo.SupportsMultiChanJoin = True
    ServerInfo.SupportsMultiChanWho = True
    ServerInfo.SupportsMultiChanMode = False
    ServerSplitDetection = False
    ServerInfo.SupportsServersChan = False
    ServerInfo.MaxNickLength = 30
  ElseIf InStr(Rest, "hybrid") > 0 Then
    ServerInfo.Network = "EFNet"
    ServerInfo.HidesHosts = False
    ServerInfo.SupportsMultiChanJoin = True
    ServerInfo.SupportsMultiChanWho = False
    ServerInfo.SupportsMultiChanMode = False
    ServerInfo.SupportsServersChan = False
    ServerSplitDetection = False
    ServerInfo.MaxNickLength = 9
  ElseIf Left(Rest, 2) = "cr" Then
    ServerInfo.Network = "FlirtNet"
    ServerInfo.HidesHosts = True
    ServerInfo.SupportsMultiChanJoin = True
    ServerInfo.SupportsMultiChanWho = False
    ServerInfo.SupportsMultiChanMode = False
    ServerInfo.SupportsServersChan = False
    ServerSplitDetection = False
    ServerInfo.MaxNickLength = 30
  ElseIf Left(Rest, 1) = "u" Then
    ServerInfo.Network = "Undernet"
    ServerInfo.HidesHosts = False
    ServerInfo.SupportsMultiChanJoin = True
    ServerInfo.SupportsMultiChanWho = False
    ServerInfo.SupportsMultiChanMode = False
    ServerInfo.SupportsServersChan = False
    ServerSplitDetection = False
    ServerInfo.MaxNickLength = 30
  Else
    ServerInfo.Network = "IRCnet"
    ServerInfo.HidesHosts = False
    ServerInfo.SupportsMultiChanJoin = True
    ServerInfo.SupportsMultiChanWho = False
    ServerInfo.SupportsMultiChanMode = True
    ServerInfo.SupportsServersChan = True
    ServerInfo.MaxNickLength = 9
    ServerInfo.SupportsMultiKicks = True
  End If
  
  If AutoNetSetup = True Then
    ServerChannelModes = Param(Line, 7)
    WritePPString "NET", "ChanModes", ServerChannelModes, NET_INI
  End If
  
  SpreadFlagMessage 0, "+m", "3*** Server type: " & ServerInfo.Network & " (" & Rest & ")"
  
  'connect burst
  If PermChanCount > 0 Then
    Initializing = True '<-- Don't send extra WHO's and MODE's on join till connect burst is over
    JoinLine = "": KeyLine = ""
    For u = 1 To PermChanCount
      Chan = PermChannels(u).Name
      SetPermChanStat Chan, ChanStat_NotOn
      Rest = GetChannelSetting(Chan, "EnforceModes", "x")
      If ParamCount(Rest) > 1 Then Rest = Param(Rest, ParamCount(Rest)) Else Rest = "x"
      If JoinLine = "" Then JoinLine = Chan: KeyLine = KeyLine + Rest Else JoinLine = JoinLine & "," & Chan: KeyLine = KeyLine & "," & Rest
      If ServerInfo.SupportsMultiChanJoin = False Then SendLine "join " & Chan & " " & Rest, 1
    Next u
    'Join channels
    If ServerInfo.SupportsMultiChanJoin Then
      If ServerInfo.SupportsServersChan Then
        SendLine "join &servers," & JoinLine & " x," & KeyLine, 1
      Else
        SendLine "join " & JoinLine & " " & KeyLine, 1
      End If
    Else
      If ServerInfo.SupportsServersChan Then SendLine "join &servers", 1
      For u = 1 To PermChanCount
        Chan = PermChannels(u).Name
        Rest = GetChannelSetting(Chan, "EnforceModes", "x")
        If ParamCount(Rest) > 1 Then Rest = Param(Rest, ParamCount(Rest)) Else Rest = "x"
        SendLine "join " & Chan & " " & Rest, 1
      Next u
    End If
    'winsock2_send WHO commands
    If ServerInfo.SupportsMultiChanWho Then
      SendLine "who " & JoinLine, 1
    Else
      For u = 1 To PermChanCount
        SendLine "who " & PermChannels(u).Name, 1
      Next u
    End If
    'winsock2_send MODE commands
    If ServerInfo.SupportsMultiChanMode Then
      SendLine "mode " & JoinLine, 1
      SendLine "mode " & JoinLine & " +b", 1
      If InStr(1, ServerChannelModes, "e", vbBinaryCompare) > 0 Then SendLine "mode " & JoinLine & " +e", 1
      If InStr(1, ServerChannelModes, "I", vbBinaryCompare) > 0 Then SendLine "mode " & JoinLine & " +I", 1
    Else
      For u = 1 To PermChanCount
        SendLine "mode " & PermChannels(u).Name, 1
        SendLine "mode " & PermChannels(u).Name & " +b", 1
        If InStr(1, ServerChannelModes, "e", vbBinaryCompare) > 0 Then SendLine "mode " & PermChannels(u).Name & " +e", 1
        If InStr(1, ServerChannelModes, "I", vbBinaryCompare) > 0 Then SendLine "mode " & PermChannels(u).Name & " +I", 1
      Next u
    End If
    SendLine "ISON :", 1 '<-- Initializing will be set back to FALSE when the reply to this command is received
  End If
End Sub

Sub RPL_ISUPPORT(Line As String)
  If AutoNetSetup = False Then Exit Sub
  If Dir(NET_INI) = "" Then
    ServerInfo.SupportsMultiChanJoin = True
    WritePPString "NET", "UseFullAdress", "0", NET_INI
    WritePPString "NET", "SplitDetection", "1", NET_INI
    WritePPString "NET", "MultiKicks", "0", NET_INI
    WritePPString "NET", "MultiWho", "0", NET_INI
  End If
  If InStr(Line, " MODES") <> 0 Then
    ServerNumberOfModes = CByte(Param(Mid(Line, InStr(Line, " MODES") + 7), 1))
    WritePPString "NET", "NumberOfModes", ServerNumberOfModes, NET_INI
  End If
  If InStr(Line, "TOPICLEN") <> 0 Then
    ServerTopicLen = CInt(Param(Mid(Line, InStr(Line, "TOPICLEN") + 9), 1))
    WritePPString "NET", "TopicLen", ServerTopicLen, NET_INI
  End If
  If InStr(Line, "CHANTYPES") <> 0 Then
    ServerChannelPrefixes = (Param(Mid(Line, InStr(Line, "CHANTYPES") + 10), 1))
    If InStr(1, ServerChannelPrefixes, "&", vbBinaryCompare) <> 0 Then ServerInfo.SupportsServersChan = True Else ServerInfo.SupportsServersChan = False
    WritePPString "NET", "ChanPrefixes", ServerChannelPrefixes, NET_INI
  End If
  If InStr(Line, "MAXCHANNELS") <> 0 Then
    ServerMaxChannels = CByte(Param(Mid(Line, InStr(Line, "MAXCHANNELS") + 12), 1))
    WritePPString "NET", "MaxChan", ServerMaxChannels, NET_INI
  End If
  If InStr(Line, "NICKLEN") <> 0 Then
    ServerNickLen = CByte(Param(Mid(Line, InStr(Line, "NICKLEN") + 8), 1))
    ServerInfo.MaxNickLength = ServerNickLen
    WritePPString "NET", "NickLength", ServerNickLen, NET_INI
  End If
  If InStr(Line, "NETWORK") <> 0 Then
    ServerNetwork = (Param(Mid(Line, InStr(Line, "NETWORK") + 8), 1))
    ServerInfo.Network = ServerNetwork
    WritePPString "NET", "NetworkName", ServerNetwork, NET_INI
  End If
  If InStr(Line, "PREFIX") <> 0 Then
    ServerUserModes = (Param(Mid(Line, InStr(Line, "PREFIX") + 7), 1))
    WritePPString "NET", "UserPrefixes", ServerUserModes, NET_INI
  End If
  If InStr(Line, "CHANMODES") <> 0 Then
    ServerChannelModes = (Param(Mid(Line, InStr(Line, "CHANMODES") + 10), 1))
    WritePPString "NET", "ChanModes", ServerChannelModes, NET_INI
  End If
End Sub

Sub RPL_USERHOST(Line As String)
  MyHostmask = StripDP(Param(Line, 4))
  If LCase(ParamX(MyHostmask, "=", 1)) = LCase(MyNick) Then
    MyNick = ParamX(MyHostmask, "=", 1)
    MyHostmask = ParamX(MyHostmask, "=", 2)
    If InStr("*+-", Left(MyHostmask, 1)) > 0 Then MyHostmask = Mid(MyHostmask, 2)
    GotMyHost MyNick & "!" & MyHostmask
  End If
End Sub

Sub RPL_ISON(Line As String)
  If Initializing = True Then
    Initializing = False
    SpreadFlagMessage 0, "+m", "3*** Finished connecting."
  End If
End Sub

Sub RPL_WHOREPLY(Line As String)
  Dim ChNum As Long, UsNum As Long, Rest As String
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  UsNum = FindUser(Param(Line, 8), ChNum)
  If UsNum > UBound(Channels(ChNum).User()) Then ReDim Preserve Channels(ChNum).User(UsNum + 5)
  If InStr(Param(Line, 9), "@") > 0 Then
    Channels(ChNum).User(UsNum).Status = "@"
  ElseIf InStr(Param(Line, 9), "!") > 0 Then
    Channels(ChNum).User(UsNum).Status = "!"
  ElseIf InStr(Param(Line, 9), "*") > 0 Then
    Channels(ChNum).User(UsNum).Status = "*"
  ElseIf InStr(Param(Line, 9), "%") > 0 Then
    Channels(ChNum).User(UsNum).Status = "%"
  ElseIf InStr(Param(Line, 9), "+") > 0 Then
    Channels(ChNum).User(UsNum).Status = "+"
  Else
    Channels(ChNum).User(UsNum).Status = ""
  End If
  Channels(ChNum).User(UsNum).Hostmask = Param(Line, 8) & "!" & Param(Line, 5) & "@" & Param(Line, 6)
  If IsValidIP(Param(Line, 6)) = False Then
    Rest = GetCacheIP(Param(Line, 6), False)
    If Rest <> "" Then
      Channels(ChNum).User(UsNum).IPmask = Param(Line, 8) & "!" & Param(Line, 5) & "@" & Rest
    End If
  End If
End Sub

Sub RPL_ENDOFWHO(Line As String)
  Dim ChNum As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  GetRegUsers ChNum
  Channels(ChNum).CompletedWHO = True
  DoAutoStuff ChNum
  'GetOps with 3 seconds delay
  TimedEvent "gop " & Channels(ChNum).Name, 3
End Sub

Sub RPL_TOPIC(Line As String)
  Dim ChNum As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  Channels(ChNum).Topic = StripDP(Trim(Right(Line, Len(Line) - Len(Param(Line, 1)) - Len(Param(Line, 2)) - Len(Param(Line, 3)) - Len(Param(Line, 4)) - 4)))
End Sub

Sub RPL_NAMREPLY(Line As String)
  Dim ChNum As Long, CurPos As Long, Name As String
  ChNum = FindChan(Param(Line, 5))
  If ChNum = 0 Then Exit Sub
  CurPos = 5
  Do
    CurPos = CurPos + 1
    Name = StripDP(Param(Line, CurPos))
    If Name = "" Then Exit Do
    If FindUser(Name, ChNum) = 0 Then
      Channels(ChNum).UserCount = Channels(ChNum).UserCount + 1
      If Channels(ChNum).UserCount > UBound(Channels(ChNum).User()) Then ReDim Preserve Channels(ChNum).User(UBound(Channels(ChNum).User()) + 5)
      Select Case Left(Name, 1)
        Case "@", "*", "!"
            Channels(ChNum).User(Channels(ChNum).UserCount).Nick = Right(Name, Len(Name) - 1)
            Channels(ChNum).User(Channels(ChNum).UserCount).Status = "@"
            If Right(Name, Len(Name) - 1) = MyNick Then
              Channels(ChNum).GotOPs = True
              'SendLine "mode " & Channels(ChNum).Name & " +tn", 1
            End If
        Case "%"
            If Right(Name, Len(Name) - 1) = MyNick Then
              Channels(ChNum).GotHOPs = True
            End If
            Channels(ChNum).User(Channels(ChNum).UserCount).Nick = Right(Name, Len(Name) - 1)
            Channels(ChNum).User(Channels(ChNum).UserCount).Status = "%"
        Case "+"
            Channels(ChNum).User(Channels(ChNum).UserCount).Nick = Right(Name, Len(Name) - 1)
            Channels(ChNum).User(Channels(ChNum).UserCount).Status = "+"
        Case Else
            Channels(ChNum).User(Channels(ChNum).UserCount).Nick = Name
            Channels(ChNum).User(Channels(ChNum).UserCount).Status = ""
      End Select
      Channels(ChNum).User(Channels(ChNum).UserCount).Hostmask = ""
      Channels(ChNum).User(Channels(ChNum).UserCount).IPmask = ""
      Channels(ChNum).User(Channels(ChNum).UserCount).LastLine = ""
      Channels(ChNum).User(Channels(ChNum).UserCount).CTCPs = 0
      Channels(ChNum).User(Channels(ChNum).UserCount).CharCount = 0
      Channels(ChNum).User(Channels(ChNum).UserCount).LineCount = 0
      Channels(ChNum).User(Channels(ChNum).UserCount).RepeatCount = 0
      Channels(ChNum).User(Channels(ChNum).UserCount).LastEvent = WinTickCount
    End If
  Loop
End Sub

Sub RPL_BANLIST(Line As String)
  Dim ChNum As Long, MatchedOne As Boolean, u As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  MatchedOne = False
  For u = 1 To Channels(ChNum).BanCount
    If LCase(Channels(ChNum).BanList(u).Mask) = LCase(Param(Line, 5)) Then MatchedOne = True
  Next u
  If Not MatchedOne Then
    Channels(ChNum).BanCount = Channels(ChNum).BanCount + 1
    Channels(ChNum).BanList(Channels(ChNum).BanCount).Mask = Param(Line, 5)
  End If
End Sub

Sub RPL_ENDOFBANLIST(Line As String)
  Dim ChNum As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  Channels(ChNum).CompletedBANS = True
  RemDisturbingBans ChNum
  If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then CheckBans Channels(ChNum).Name
End Sub

Sub RPL_EXCEPTLIST(Line As String)
  Dim ChNum As Long, MatchedOne As Boolean, u As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  MatchedOne = False
  For u = 1 To Channels(ChNum).ExceptCount
    If LCase(Channels(ChNum).ExceptList(u).Mask) = LCase(Param(Line, 5)) Then MatchedOne = True
  Next u
  If Not MatchedOne Then
    Channels(ChNum).ExceptCount = Channels(ChNum).ExceptCount + 1
    Channels(ChNum).ExceptList(Channels(ChNum).ExceptCount).Mask = Param(Line, 5)
  End If
End Sub

Sub RPL_ENDOFEXCEPTLIST(Line As String)
  Dim ChNum As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  Channels(ChNum).CompletedExcepts = True
  If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then CheckExcepts Channels(ChNum).Name
End Sub

Sub RPL_INVITELIST(Line As String)
  Dim ChNum As Long, MatchedOne As Boolean, u As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  MatchedOne = False
  For u = 1 To Channels(ChNum).InviteCount
    If LCase(Channels(ChNum).InviteList(u).Mask) = LCase(Param(Line, 5)) Then MatchedOne = True
  Next u
  If Not MatchedOne Then
    Channels(ChNum).InviteCount = Channels(ChNum).InviteCount + 1
    Channels(ChNum).InviteList(Channels(ChNum).InviteCount).Mask = Param(Line, 5)
  End If
End Sub

Sub RPL_ENDOFINVITELIST(Line As String)
  Dim ChNum As Long
  ChNum = FindChan(Param(Line, 4))
  If ChNum = 0 Then Exit Sub
  Channels(ChNum).CompletedInvites = True
  If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then CheckInvites Channels(ChNum).Name
End Sub

Sub RPL_CHANNELMODEIS(Line As String)
  Dim Chan As String, ChNum As Long, Rest As String, ChangMode As String
  Chan = Param(Line, 4)
  ChNum = FindChan(Param(Line, 4))
  Channels(ChNum).Mode = GetRest(Line, 5)
  Channels(ChNum).CompletedMode = True
  If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then
    Rest = GetChannelSetting(Channels(ChNum).Name, "EnforceModes", "")
    If (GetModeChar(Rest, "t") = 0) And (GetModeChar(Rest, "n") = 0) Then
      Rest = "+tn" & Rest
    ElseIf (GetModeChar(Rest, "t") = 0) Then
      Rest = "+t" & Rest
    ElseIf (GetModeChar(Rest, "n") = 0) Then
      Rest = "+n" & Rest
    End If
    ChangMode = ChangeMode(Rest, Channels(ChNum).Mode)
    If ChangMode <> "" Then SendLine "mode " & Chan & " " & ChangMode, 1
  End If
End Sub

Sub ERR_NOSUCHNICK(Line As String)
  If Left(Param(Line, 4), 13) <> "#anti-i-chan-" And Initializing = False Then
    SpreadMessage 0, -1, "5*** Server: " & Param(Line, 4) & " no such nick/channel"
  End If
End Sub

Sub ERR_NICKNAMEINUSE(Line As String)
  If Not Connected Then
    If Param(Line, 4) = PrimaryNick Then SendIt "NICK " & SecondaryNick + vbCrLf: TimedEvent "NICK " & PrimaryNick, 20: Exit Sub
    If Param(Line, 4) = SecondaryNick Then SendIt "NICK " & Left(PrimaryNick, 5) + Trim(Str(Int(Rnd * 9999))) + vbCrLf: Exit Sub
  End If
  If Connected Then
    If Param(Line, 4) = PrimaryNick Then TimedEvent "NICK " & PrimaryNick, 20
  End If
End Sub

Sub ERR_UNAVAILRESOURCE(Line As String)
  If IsValidChannel(Param(Line, 4)) Then
    SetPermChanStat Param(Line, 4), ChanStat_Duped
  Else
    If Not Connected Then
      If Param(Line, 4) = PrimaryNick Then SendIt "NICK " & SecondaryNick + vbCrLf: TimedEvent "NICK " & PrimaryNick, 20: Exit Sub
      If Param(Line, 4) = SecondaryNick Then SendIt "NICK " & Left(PrimaryNick, 5) + Trim(Str(Int(Rnd * 9999))) + vbCrLf: Exit Sub
    End If
    If Connected Then
      If Param(Line, 4) = PrimaryNick Then TimedEvent "NICK " & PrimaryNick, 20
    End If
  End If
End Sub

Sub ERR_NOSUCHCHANNEL(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_Unsup
End Sub

Sub ERR_TOOMANYCHANNELS(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_OutLimits
End Sub

Sub ERR_CHANNELISFULL(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_BadLimit
End Sub

Sub ERR_INVITEONLYCHAN(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_NeedInvite
  SendGetOps "invite", Chan, MyNick
End Sub

Sub ERR_BANNEDFROMCHAN(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_ImBanned
  SendGetOps "unban", Chan, MyHostmask & " " & MyIPmask
End Sub

Sub ERR_BADCHANNELKEY(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_ImBanned
  SendGetOps "key", Chan, MyNick
End Sub

Sub ERR_REGISTERONLY(Line As String)
  Dim Chan As String
  Chan = Param(Line, 4)
  SetPermChanStat Chan, ChanStat_RegisteredOnly
End Sub
  
Sub HandlePrivateMessage(Line As String)
Dim u As Long, u2 As Long, RegUser As String, UserFlags As String, Nick As String
Dim GaveOps As Boolean, UsNum As Long, Rest As String, HostNum As Long
Dim NoteCount As Long, NoteDate As String, NoteFrom As String
Dim NoteText As String, NoteFlag As String, HostOnly As String
Dim IP As String, IP2 As String, Port As String, strRemoteIP As String, NewSock As Long
Dim MatchedOne As Boolean, ScNum As Long
Dim FullAddress As String, ValidSession As Boolean

  FullAddress = StripDP(Param(Line, 1))
  
  If IsIgnoredHost(FullAddress) Then Exit Sub
  If Param(Line, 1) <> ServerName And InStr(Param(Line, 1), "!") > 1 Then Nick = StripDP(Left(Param(Line, 1), InStr(Param(Line, 1), "!") - 1))
  
  RegUser = SearchUserFromHostmask(FullAddress)
  UserFlags = GetUserFlags(RegUser)
  Rest = StripDP(GetRest(Line, 4))
  
  If RegUser <> "" Then
    ValidSession = (Mid(FullAddress, InStr(FullAddress, "@") + 1) = BotUsers(GetUserNum(RegUser)).OldFullAddress)
  Else
    ValidSession = False
  End If
  
  'Check script Hooks
  HaltDefault = False
  If Nick <> MyNick Then
    If LCase(Param(Rest, 1)) = "action" Then '--- /ME ACTION
      Rest = GetRest(Rest, 2): Rest = Left(Rest, Len(Rest) - 1)
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Priv_act Then
          RunScriptX ScNum, "Priv_act", Nick, RegUser, Rest
        End If
      Next ScNum
    ElseIf Left(Rest, 1) = "" Then           '--- CTCP
      If Right(Rest, 1) = "" And Len(Rest) > 1 Then Rest = Mid(Rest, 2, Len(Rest) - 2) Else Rest = Mid(Rest, 2, Len(Rest) - 1)
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Priv_ctcp Then
          RunScriptX ScNum, "Priv_ctcp", Nick, RegUser, Rest
        End If
      Next ScNum
    Else                                       '--- REAL PRIVMSG
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.Priv_msg Then
          RunScriptX ScNum, "Priv_msg", Nick, RegUser, Rest
        End If
      Next ScNum
    End If
  End If
  Rest = ""
  
  If Not HaltDefault Then
    Select Case StripDP(LCase(Param(Line, 4)))
      'First time superowner keyword
      Case FirstTimeKeyword
          Select Case AddUser(Nick, "fijmnopstvw")
            Case AU_Success
              Select Case AddHost(0, Nick, Mask(FullAddress, 23))
                Case AH_Success
                  SendLine "privmsg " & Nick & " :Hello " & Nick & " - welcome to my user list!", 1
                  SendLine "privmsg " & Nick & " :" & String(Len("Hello " & Nick & " - welcome to my user list!"), "-"), 1
                  SendLine "privmsg " & Nick & " :You are now a SUPER OWNER. That's my highest user level.", 1
                  SendLine "privmsg " & Nick & " :Your hostmask is: " & Mask(FullAddress, 23), 1
                  SendLine "privmsg " & Nick & " : ", 1
                  SendLine "privmsg " & Nick & " :You should open a DCC Chat to me now to finish my configuration.", 1
                  SendLine "privmsg " & Nick & " :To do this, just type: /dcc chat " & MyNick, 1
                  UpdateRegUsers "A " & Mask(FullAddress, 23)
                  FirstTimeKeyword = "°NOKEY"
                  DeletePPString "Identification", "FirstTimeKeyword", AnGeL_INI
                Case Else
                  SendLine "notice " & Nick & " :Sorry, I couldn't add your host. Perhaps you are already added or the user list is broken.", 1
              End Select
            Case AU_TooLong
              SendLine "notice " & Nick & " :Sorry, I couldn't add you. Your nick seems to long to me. Try nickchanging to a shorter nick.", 1
            Case AU_InvalidNick
              SendLine "notice " & Nick & " :Sorry, I couldn't add you. Your nick seems to contain invalid chars. Try a nick only with letters.", 1
            Case Else
              SendLine "notice " & Nick & " :Sorry, I couldn't add you. Perhaps you are already added or the user list is broken.", 1
          End Select
      'Help request
      Case "help"
          If RegUser <> "" Then
            If IgnoreCheck(FullAddress, 2, 20) Then
              If RegUser <> Nick Then SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & RegUser & ") requested HELP ..." Else SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " requested HELP ..."
              SendLine "privmsg " & Nick & " :/MSG commands for all users:", 3
              SendLine "privmsg " & Nick & " :   PASS <newpass>            Sets your password.", 3
              SendLine "privmsg " & Nick & " :   PASS <oldpass> <newpass>  Changes your password.", 3
              SendLine "privmsg " & Nick & " :   SEEN <nick>               Tells you when/where I last saw <nick>.", 3
              SendLine "privmsg " & Nick & " :   NOTES <pass>              Sends you all notes waiting for you.", 3
              If MatchFlags(AllFlags(RegUser), "+o") Then
                SendLine "privmsg " & Nick & " :/MSG commands for channel ops or global ops:", 3
                SendLine "privmsg " & Nick & " :   OP <pass> ([" & ServerChannelPrefixes & "]channel)      Ops you on every channel I'm on.", 3
                SendLine "privmsg " & Nick & " :   VOICE <pass> ([" & ServerChannelPrefixes & "]channel)   Voices you on every channel I'm on.", 3
                SendLine "privmsg " & Nick & " :   KEY <pass> <[" & ServerChannelPrefixes & "]channel>     Sends you the key for [" & ServerChannelPrefixes & "]channel.", 3
                SendLine "privmsg " & Nick & " :   INVITE <pass> <[" & ServerChannelPrefixes & "]channel>  Invites you to [" & ServerChannelPrefixes & "]channel.", 3
                SendLine "privmsg " & Nick & " :   GO <[" & ServerChannelPrefixes & "]channel>             Makes me rejoin [" & ServerChannelPrefixes & "]channel if I don't have ops.", 3
              End If
              If MatchFlags(GetUserFlags(RegUser), "+s") Then
                SendLine "privmsg " & Nick & " :/MSG commands for super owners:", 3
                SendLine "privmsg " & Nick & " :   RESTART <pass>            Makes me die and restart again.", 3
              End If
              SendLine "privmsg " & Nick & " :--- That's it! ---", 3
            End If
          End If
      'GO request
      Case "go"
          If RegUser <> "" Then
            u = FindChan(Param(Line, 5))
            If u = 0 Then Exit Sub
            If Not Channels(u).GotOPs Then
              GaveOps = False
              For u2 = 1 To Channels(u).UserCount
                Select Case Channels(u).User(u2).Status
                  Case "@", "@+": GaveOps = True: Exit For
                End Select
              Next u2
              If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") And Not GaveOps Then
                If Not IsOrdered("GO " & Channels(u).Name) Then
                  Order "GO " & Channels(u).Name, 15
                  SendLine "part " & Channels(u).Name & " :Rejoining... ", 1
                  TimedEvent "join " & Channels(u).Name, 5
                  If RegUser <> Nick Then
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & RegUser & ") requested me to GO " & Channels(u).Name & "..."
                  Else
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " requested me to GO " & Channels(u).Name & "..."
                  End If
                End If
              End If
            End If
          End If
      'OP request
      Case "op"
          GaveOps = False
          If (RegUser <> "") Then UsNum = GetUserNum(RegUser): If UsNum = 0 Then Exit Sub
          If (RegUser <> "") And (IIf(MatchFlags(BotUsers(UsNum).Flags, "-b"), EncryptIt(Param(Line, 5)), Param(Line, 5)) = BotUsers(UsNum).Password) And (BotUsers(UsNum).Password <> "") Then
            BotUsers(UsNum).ValidSession = True
            ValidSession = True
            'Op user everywhere he/she has op rights for
            If Not IsValidChannel(Left(Param(Line, 6), 1)) Then
              MatchedOne = False
              For u = 1 To ChanCount
                If Channels(u).GotOPs Then
                  UsNum = FindUser(Nick, u)
                  If UsNum > 0 Then
                    MatchedOne = True
                    If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") Then
                      If InStr(Channels(u).User(UsNum).Status, "@") = 0 Then GiveOp Channels(u).Name, Nick: If Rest = "" Then Rest = Channels(u).Name Else Rest = Rest & ", " & Channels(u).Name
                      GaveOps = True
                    End If
                  End If
                End If
              Next u
            'Op user on a specific channel
            Else
              u = FindChan(Param(Line, 6))
              If u = 0 Then
                If IgnoreCheck(FullAddress, 2, 5) Then
                  SendLine "notice " & Nick & " :Sorry, I'm not on this channel.", 3
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed OP for " & Param(Line, 6) & " - not one of my channels"
                End If
                Exit Sub
              End If
              If Not Channels(u).GotOPs Then
                If IgnoreCheck(FullAddress, 2, 5) Then
                  SendLine "notice " & Nick & " :Sorry, I don't have ops.", 3
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed OP for " & Param(Line, 6) & " - I don't have ops"
                End If
                Exit Sub
              End If
              UsNum = FindUser(Nick, u)
              If UsNum = 0 Then
                If IgnoreCheck(FullAddress, 2, 5) Then
                  SendLine "notice " & Nick & " :Sorry, I can't see you on this channel!", 3
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed OP for " & Param(Line, 6) & " - user is not on this channel"
                End If
                Exit Sub
              End If
              If UsNum > 0 Then
                If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") Then
                  MatchedOne = True
                  If InStr(Channels(u).User(UsNum).Status, "@") = 0 Then GiveOp Channels(u).Name, Nick: If Rest = "" Then Rest = Channels(u).Name Else Rest = Rest & ", " & Channels(u).Name
                  GaveOps = True
                End If
              End If
            End If
            If MatchedOne Then
              If GaveOps Then
                If Rest <> "" Then Rest = " for " & Rest
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested OP" & Rest
              Else
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed OP - no op flag"
                  SendLine "notice " & Nick & " :Sorry, you're not allowed to get ops from me.", 3
                End If
              End If
            Else
              If IgnoreCheck(FullAddress, 2, 10) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed OP - not on my channels"
                SendLine "notice " & Nick & " :Sorry, you're not on one of my channels.", 3
              End If
            End If
          Else
            If IgnoreCheck(FullAddress, 2, 5) Then
              If RegUser <> "" Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed OP - wrong password"
                SendLine "notice " & Nick & " :Sorry, wrong password.", 3
              Else
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & Mask(FullAddress, 10) & ") failed OP - user unknown"
              End If
            End If
          End If
      'VOICE request
      Case "voice"
          GaveOps = False
          If RegUser <> "" Then UsNum = GetUserNum(RegUser): If UsNum = 0 Then Exit Sub
          If RegUser <> "" And EncryptIt(Param(Line, 5)) = BotUsers(UsNum).Password Then
            BotUsers(UsNum).ValidSession = True
            ValidSession = True
            'Voice user everywhere he/she has voice rights for
            If Not IsValidChannel(Left(Param(Line, 6), 1)) Then
              For u = 1 To ChanCount
                If Channels(u).GotOPs Or Channels(u).GotHOPs Then
                  UsNum = FindUser(Nick, u)
                  If UsNum > 0 Then
                    If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") Or MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+v") Then
                      If InStr(Channels(u).User(UsNum).Status, "+") = 0 Then SendLine "mode " & Channels(u).Name & " +v " & Nick, 1: If Rest = "" Then Rest = Channels(u).Name Else Rest = Rest & ", " & Channels(u).Name
                      GaveOps = True
                    End If
                  End If
                End If
              Next u
            'Voice user on a specific channel
            Else
              u = FindChan(Param(Line, 6))
              If u = 0 Then
                If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :Sorry, I'm not on this channel.", 3
                Exit Sub
              End If
              If Not (Channels(u).GotOPs Or Channels(u).GotHOPs) Then
                If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :Sorry, I don't have ops.", 3
                Exit Sub
              End If
              UsNum = FindUser(Nick, u)
              If UsNum = 0 Then
                If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :Sorry, I can't see you on this channel!", 3
                Exit Sub
              End If
              If UsNum > 0 Then
                If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") Or MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+v") Then
                  If InStr(Channels(u).User(UsNum).Status, "+") = 0 Then SendLine "mode " & Channels(u).Name & " +v " & Nick, 1: If Rest = "" Then Rest = Channels(u).Name Else Rest = Rest & ", " & Channels(u).Name
                  GaveOps = True
                End If
              End If
            End If
            If GaveOps Then
              If Rest <> "" Then Rest = " for " & Rest
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested VOICE" & Rest
            Else
              If IgnoreCheck(FullAddress, 2, 15) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed VOICE - no op/voice flag"
                SendLine "notice " & Nick & " :Sorry, you're not allowed to get voice from me.", 3
              End If
            End If
          Else
            If IgnoreCheck(FullAddress, 2, 5) Then
              If RegUser <> "" Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed VOICE - wrong password"
              Else
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & Mask(FullAddress, 10) & ") failed VOICE - user unknown"
              End If
            End If
          End If
      'INVITE request
      Case "invite"
          GaveOps = False
          If LCase(Param(Line, 5)) = "" Then Exit Sub
          If RegUser <> "" Then UsNum = GetUserNum(RegUser): If UsNum = 0 Then Exit Sub
          If RegUser <> "" And EncryptIt(Param(Line, 5)) = BotUsers(UsNum).Password Then
            BotUsers(UsNum).ValidSession = True
            ValidSession = True
            'User didn't tell me a channel :)
            If Not IsValidChannel(Left(Param(Line, 6), 1)) Then
              If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :You forgot to tell me the channel :)", 3
              Exit Sub
            'Invite user to a specific channel
            Else
              u = FindChan(Param(Line, 6))
              If u = 0 Then
                If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :Sorry, I'm not on this channel.", 3
                Exit Sub
              End If
              If Not (Channels(u).GotOPs Or Channels(u).GotHOPs) Then
                If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :Sorry, I don't have ops.", 3
                Exit Sub
              End If
              If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") Then
                SendLine "invite " & Nick & " " & Channels(u).Name, 1
                If GetChannelKey(u) <> "" Then SendLine "notice " & Nick & " :Key for " & Channels(u).Name & ": " & GetChannelKey(u), 1
                GaveOps = True
              End If
            End If
            If GaveOps Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested INVITE " & Channels(u).Name
            Else
              If IgnoreCheck(FullAddress, 2, 15) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed INVITE " & Channels(u).Name & " - no op flag"
                SendLine "notice " & Nick & " :Sorry, you're not allowed to get invited by me.", 3
              End If
            End If
          Else
            If IgnoreCheck(FullAddress, 2, 5) Then
              If Param(Line, 6) <> "" Then
                If RegUser <> "" Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed INVITE " & Param(Line, 6) & " - wrong password"
                Else
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & Mask(FullAddress, 10) & ") failed INVITE " & Param(Line, 6) & " - user unknown"
                End If
              Else
                If RegUser <> "" Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed INVITE - wrong password"
                Else
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & Mask(FullAddress, 10) & ") failed INVITE - user unknown"
                End If
              End If
            End If
          End If
      'KEY request
      Case "key"
          GaveOps = False
          If LCase(Param(Line, 5)) = "" Then Exit Sub
          If RegUser <> "" Then UsNum = GetUserNum(RegUser): If UsNum = 0 Then Exit Sub
          If RegUser <> "" And EncryptIt(Param(Line, 5)) = BotUsers(UsNum).Password Then
            BotUsers(UsNum).ValidSession = True
            ValidSession = True
            'User didn't tell me a channel :)
            If Not IsValidChannel(Left(Param(Line, 6), 1)) Then
              If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :You forgot to tell me the channel :)", 3
              Exit Sub
            'Invite user to a specific channel
            Else
              u = FindChan(Param(Line, 6))
              If u = 0 Then
                If IgnoreCheck(FullAddress, 2, 5) Then SendLine "notice " & Nick & " :Sorry, I'm not on this channel.", 3
                Exit Sub
              End If
              If MatchFlags(GetUserChanFlags(RegUser, Channels(u).Name), "+o") Then
                If GetChannelKey(u) <> "" Then
                  SendLine "notice " & Nick & " :Key for " & Channels(u).Name & ": " & GetChannelKey(u), 1
                Else
                  SendLine "notice " & Nick & " :There's no key in " & Channels(u).Name & "!", 1
                End If
                GaveOps = True
              End If
            End If
            If GaveOps Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested KEY " & Channels(u).Name
            Else
              If IgnoreCheck(FullAddress, 2, 15) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed KEY " & Channels(u).Name & " - no op flag"
                SendLine "notice " & Nick & " :Sorry, you're not allowed to get the key from me.", 3
              End If
            End If
          Else
            If IgnoreCheck(FullAddress, 2, 5) Then
              If Param(Line, 6) <> "" Then
                If RegUser <> "" Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed KEY " & Param(Line, 6) & " - wrong password"
                Else
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & Mask(FullAddress, 10) & ") failed KEY " & Param(Line, 6) & " - user unknown"
                End If
              Else
                If RegUser <> "" Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed KEY - wrong password"
                Else
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " (" & Mask(FullAddress, 10) & ") failed KEY - user unknown"
                End If
              End If
            End If
          End If
      'Restart
      Case "restart"
          If RegUser <> "" Then
            UsNum = GetUserNum(RegUser): If UsNum = 0 Then Exit Sub
            If EncryptIt(Param(Line, 5)) = BotUsers(UsNum).Password Then
              BotUsers(UsNum).ValidSession = True
              ValidSession = True
              If MatchFlags(GetUserFlags(RegUser), "+s") Then
                If Dir(HomeDir & App.EXEName & ".exe") = "" Then SendLine "notice " & Nick & " :Sorry, AnGeL Binary not found!", 3: Exit Sub
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested RESTART"
                Status "*** Restart requested..." & vbCrLf
                SpreadMessage 0, -1, "7*** RESTART requested by " & RegUser & ""
                TimedEvent "RESTART", 0
              End If
            Else
              If IgnoreCheck(FullAddress, 2, 15) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " failed RESTART: wrong password"
              End If
            End If
          End If
      'Identify
      Case LCase(IdentCommand)
          If IdentCommand = "°NONE" Then
            If IgnoreCheck(FullAddress, 2, 15) Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & FullAddress
              SpreadFlagMessage 0, "+m", "14[" & Time & "]     IDENT error: command is disabled"
              If HideBot = False Then SendLine "notice " & Nick & " :Sorry, IDENT has been disabled.", 3
            End If
            Exit Sub
          End If
          If RegUser <> "" Then
            If IgnoreCheck(FullAddress, 2, 15) Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " IDENT error: hostmask already known"
              SendLine "notice " & Nick & " :Your current hostmask is already known to me!", 3
            End If
          Else
            If Param(Line, 5) = "" Then
              If IgnoreCheck(FullAddress, 2, 15) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " IDENT error: password not specified"
                If HideBot = False Then SendLine "notice " & Nick & " :Usage: '/msg " & MyNick & " " & IdentCommand & " <password> (your nick in the bot)'", 3
              End If
              Exit Sub
            End If
            If Param(Line, 6) = "" Then Line = Line & " " & Nick
            RegUser = GetRealNick(Param(Line, 6))
            If RegUser = "" Then
              If IgnoreCheck(FullAddress, 2, 15) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " IDENT error: user unknown"
                If HideBot = False Then
                  SendLine "notice " & Nick & " :Sorry, I didn't recognize you by your current nick.", 3
                  SendLine "notice " & Nick & " :Usage: '/msg " & MyNick & " " & IdentCommand & " <password> <your nick in the bot>'", 3
                End If
              End If
              Exit Sub
            End If
            If BotUsers(GetUserNum(Param(Line, 6))).Password = EncryptIt(Param(Line, 5)) Then
              BotUsers(GetUserNum(Param(Line, 6))).ValidSession = True
              ValidSession = True
              Select Case AddHost(0, RegUser, Mask(FullAddress, 23))
                Case AH_Success
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " IDENT success for: " & Mask(FullAddress, 23)
                  SendLine "notice " & Nick & " :Your hostmask " & Mask(FullAddress, 23) & " was added to my database.", 3
                  UpdateRegUsers "I " & Nick & " " & RegUser
                Case AH_MatchingUser
                  If IgnoreCheck(FullAddress, 2, 15) Then
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " IDENT error: Hostmask is matching " & ExtReply
                    SendLine "notice " & Nick & " :Sorry, I couldn't add your host - it's matching " & ExtReply & ". Please ask a master or owner for help.", 3
                  End If
                Case AH_TooManyHosts
                  If IgnoreCheck(FullAddress, 2, 15) Then
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " IDENT error: max. number of hostmasks reached (20)"
                    SendLine "notice " & Nick & " :Sorry, you've reached the maximum number of hostmasks (20). Ask a master or owner for help.", 3
                  End If
                Case Else
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " IDENT error: Couldn't add new host!"
                  SendLine "notice " & Nick & " :Sorry, I couldn't add your host. Please ask a master or owner for help.", 3
              End Select
            Else
              If IgnoreCheck(FullAddress, 2, 5) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " IDENT error: wrong password"
                If HideBot = False Then SendLine "notice " & Nick & " :Sorry, wrong password.", 3
              End If
            End If
          End If
      'Wrong ident command used / ident is disabled
      Case "ident"
          If IdentCommand = "°NONE" Then
            If IgnoreCheck(FullAddress, 2, 15) Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & FullAddress
              SpreadFlagMessage 0, "+m", "14[" & Time & "]     IDENT error: command is disabled"
              If HideBot = False Then SendLine "notice " & Nick & " :Sorry, IDENT has been disabled.", 3
            End If
          Else
            If IgnoreCheck(FullAddress, 2, 15) Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & FullAddress
              SpreadFlagMessage 0, "+m", "14[" & Time & "]     IDENT error: wrong command used"
              If HideBot = False Then SendLine "notice " & Nick & " :Sorry, the IDENT command has been changed. Please ask a bot owner for help.", 3
            End If
          End If
      'Pass
      Case "pass"
          If Param(Line, 5) = "" Then Exit Sub
          If RegUser <> "" Then
            UsNum = GetUserNum(RegUser)
            If BotUsers(UsNum).Password <> "" Then
              If Param(Line, 6) = "" Then
                'User tries to set a password, but there's already one set
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS error: password already set"
                  SendLine "notice " & Nick & " :You already have a password set.", 3
                End If
              Else
                'Change password to something else
                If BotUsers(UsNum).Password = EncryptIt(Param(Line, 5)) Then
                  BotUsers(UsNum).ValidSession = True
                  ValidSession = True
                  If Len(Param(Line, 6)) < 6 Then
                    If IgnoreCheck(FullAddress, 2, 15) Then
                      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS error: password too short"
                      SendLine "notice " & Nick & " :" & MakeMsg(ERR_Pass_TooShort), 3
                    End If
                  Else
                    If WeakPass(Param(Line, 6), RegUser) Then
                      'Weak password
                      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS error: weak password"
                      SendLine "notice " & Nick & " :" & MakeMsg(ERR_Pass_TooWeak), 3
                    Else
                      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS ..."
                      BotUsers(UsNum).Password = EncryptIt(Param(Line, 6))
                      BotUsers(UsNum).ValidSession = True
                      SendLine "notice " & Nick & " :Your password was changed to '" & Param(Line, 6) & "'.", 3
                    End If
                  End If
                End If
              End If
            Else
              If Len(Param(Line, 5)) < 6 Then
                'User specifies a password < 6 characters... too short!
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS error: password too short"
                  SendLine "notice " & Nick & " :" & MakeMsg(ERR_Pass_TooShort), 3
                End If
              Else
                If WeakPass(Param(Line, 5), RegUser) Then
                  'Weak password
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS error: weak password"
                  SendLine "notice " & Nick & " :" & MakeMsg(ERR_Pass_TooWeak), 3
                Else
                  'Set a new user password
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " PASS ..."
                  BotUsers(UsNum).Password = EncryptIt(Param(Line, 5))
                  BotUsers(UsNum).ValidSession = True
                  SendLine "notice " & Nick & " :Your password was set to '" & Param(Line, 5) & "'.", 3
                  u = NotesCount(RegUser)
                  If u > 0 Then
                    If u = 1 Then
                      SendLine "notice " & Nick & " :Hi! I've got 1 note waiting for you.", 3
                      SendLine "notice " & Nick & " :To get it, type: /msg " & MyNick & " notes <pass>", 3
                    Else
                      SendLine "notice " & Nick & " :Hi! I've got " & CStr(u) & " notes waiting for you.", 3
                      SendLine "notice " & Nick & " :To get them, type: /msg " & MyNick & " notes <pass>", 3
                    End If
                  End If
                End If
              End If
            End If
          Else
            If IgnoreCheck(FullAddress, 2, 15) Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick & " PASS error: user unknown"
              If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I don't know who you are.", 3
            End If
          End If
      'Notes
      Case "notes"
          If RegUser <> "" Then
            UsNum = GetUserNum(RegUser)
            If Param(Line, 5) = "" Then
              If IgnoreCheck(FullAddress, 2, 10) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " NOTES error: password not specified"
                If HideBot = False Then SendLine "notice " & Nick & " :Usage: '/msg " & MyNick & " notes <password>'", 3
              End If
            Else
              If BotUsers(UsNum).Password = EncryptIt(Param(Line, 5)) Then
                BotUsers(UsNum).ValidSession = True
                ValidSession = True
                NoteCount = NotesCount(RegUser)
                If NoteCount > 0 Then
                  For u = 1 To NoteCount
                    NoteFlag = NotesFlag(RegUser, u)
                    NoteFrom = NotesFrom(RegUser, u)
                    NoteDate = NotesDate(RegUser, u)
                    NoteText = NotesText(RegUser, u)
                    If GetUserData(UsNum, "colors", SF_NO) = SF_YES Then
                      If NoteFlag = "" Then
                        SendLine "notice " & Nick & " :2" & CStr(u) & ".3 " & NoteFrom & " 14(" & Format(GetDate(NoteDate), "dd.mm.yy, hh:nn") & "): " & NoteText, 3
                      Else
                        SendLine "notice " & Nick & " :2" & CStr(u) & ".3 " & NoteFrom & " 10[" & NoteFlag & "] 14(" & Format(GetDate(NoteDate), "dd.mm.yy, hh:nn") & "): " & NoteText, 3
                      End If
                    Else
                      If NoteFlag = "" Then
                        SendLine "notice " & Nick & " :" & CStr(u) & ". " & NoteFrom & " (" & Format(GetDate(NoteDate), "dd.mm.yy, hh:nn") & "): " & NoteText, 3
                      Else
                        SendLine "notice " & Nick & " :" & CStr(u) & ". " & NoteFrom & " [" & NoteFlag & "] (" & Format(GetDate(NoteDate), "dd.mm.yy, hh:nn") & "): " & NoteText, 3
                      End If
                    End If
                  Next u
                  NotesErase RegUser
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested NOTES ..."
                Else
                  If IgnoreCheck(FullAddress, 2, 10) Then
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " requested NOTES ..."
                    SendLine "notice " & Nick & " :There are no notes waiting for you.", 3
                  End If
                End If
              Else
                If IgnoreCheck(FullAddress, 2, 10) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " NOTES error: wrong password"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, wrong password.", 3
                End If
              End If
            End If
          End If
      'DCC Chat request
      Case "dcc"
          If UCase(Param(Line, 5) & " " & Param(Line, 6)) = "CHAT CHAT" Then
            If RegUser = "" Then
              If UnIgnoreTimes = 0 Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being DCC flooded. Ignoring invalid DCCs for 2 minutes.": Exit Sub
                  SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLDCCRefused, Nick, Mask(FullAddress, 2))
                  If HideBot = False Then
                    SendLine "notice " & Nick & " :Sorry, I don't chat with strangers.", 3
                    SendLine "notice " & Nick & " :If I should know you, type '/msg " & MyNick & " <ident command> <password>'.", 3
                  End If
                End If
              End If
              Exit Sub
            End If
            If MatchFlags(UserFlags, "-p") Then
              If IgnoreCheck(FullAddress, 2, 15) Then
                If HideBot = False Then SendLine "notice " & Nick & " :" & MakeMsg(ERR_Login_NoChat), 3
                SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLDCCNoAcc, Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", ""))
              End If
              Exit Sub
            End If
            IP = Param(Line, 7)
            IP2 = ""
            Port = Param(Line, 8): If Right(Port, 1) = "" Then Port = Left(Port, Len(Port) - 1)
            If (Port = "") Or (IP = "") Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Chat with " & RegUser & " failed (missing parameters)!"
              Exit Sub
            End If
            On Local Error Resume Next
            If CLng(Port) < 1024 Or CLng(Port) > 64000 Then Exit Sub
            If Err.Number <> 0 Then
              Err.Clear
              Exit Sub
            End If
            On Error GoTo 0
            
            HostOnly = Mask(FullAddress, 11)
            If IsValidIP(HostOnly) = False Then HostOnly = IrcGetLongIp(GetCacheIP(HostOnly, True)) Else HostOnly = IrcGetLongIp(HostOnly)
            'If resolved IP differs from given IP -> save to IP2
            If HostOnly <> IP Then
              If HostOnly <> "4294967295" Then IP2 = HostOnly
            End If
            
            NewSock = AddSocket
            SocketItem(NewSock).Hostmask = FullAddress
            SocketItem(NewSock).RegNick = RegUser
            SocketItem(NewSock).IRCNick = Nick
            SocketItem(NewSock).Flags = UserFlags
            SocketItem(NewSock).UserNum = GetUserNum(RegUser)
            SetSockFlag NewSock, SF_Colors, GetUserData(SocketItem(NewSock).UserNum, "colors", SF_NO)
            SetSockFlag NewSock, SF_Status, SF_Status_DCCWaiting
            SetSockFlag NewSock, SF_Echo, SF_NO
            SetSockFlag NewSock, SF_DCC, SF_YES
            SetSockFlag NewSock, SF_LF_ONLY, SF_YES
            SocketItem(NewSock).OnBot = BotNetNick
            SocketItem(NewSock).LinkStatus = "DCC"
            SocketItem(NewSock).PLChannel = BotUsers(SocketItem(NewSock).UserNum).PLChannel
            
            SpreadFlagMessageEx u, "+m", SF_Local_JP, MakeMsg(MSG_PLDCCIncoming, SocketItem(NewSock).RegNick, IrcGetAscIp(IP))
            If Err.Number = 0 Then
              If ConnectTCP(NewSock, IP, CLng(Port)) <> 0 Then
                RemoveSocket NewSock, 0, "", True
              End If
              If IP2 <> "" Then
                NewSock = AddSocket
                SocketItem(NewSock).Hostmask = FullAddress
                SocketItem(NewSock).RegNick = RegUser
                SocketItem(NewSock).IRCNick = Nick
                SocketItem(NewSock).Flags = UserFlags
                SocketItem(NewSock).UserNum = GetUserNum(RegUser)
                SetSockFlag NewSock, SF_Colors, GetUserData(SocketItem(NewSock).UserNum, "colors", SF_NO)
                SetSockFlag NewSock, SF_Status, SF_Status_DCCWaiting
                SetSockFlag NewSock, SF_Echo, SF_NO
                SetSockFlag NewSock, SF_DCC, SF_YES
                SetSockFlag NewSock, SF_LF_ONLY, SF_YES
                SocketItem(NewSock).OnBot = BotNetNick
                SocketItem(NewSock).LinkStatus = "DCC"
                SocketItem(NewSock).PLChannel = BotUsers(SocketItem(NewSock).UserNum).PLChannel
                If Err.Number = 0 Then
                  If ConnectTCP(NewSock, IP2, CLng(Port)) <> 0 Then
                    RemoveSocket NewSock, 0, "", True
                  End If
                Else
                  Err.Clear
                  RemoveSocket NewSock, 0, "", True
                End If
              Else
                Err.Clear
                RemoveSocket NewSock, 0, "", True
              End If
            Else
              Err.Clear
              RemoveSocket NewSock, 0, "", True
            End If
          End If
          'DCC Resume request
          If UCase(Param(Line, 5)) = "RESUME" Then
            Port = Param(Line, 7)
            On Local Error Resume Next
            For u = 1 To SocketCount
              If IsValidSocket(u) Then
                If GetSockFlag(u, SF_Status) = SF_Status_SendFileWaiting Then
                  If SocketItem(u).RemotePort = CLng(Port) Then
                    u2 = CLng(Left(Param(Line, 8), Len(Param(Line, 8)) - 1))
                    If (Err.Number <> 0) Or (u2 > SocketItem(u).FileSize) Or (u2 < 0) Then
                      SpreadLevelFileAreaMessage 0, "14[" & Time & "] *** DCC send to " & SocketItem(u).RegNick & " aborted (invalid RESUME: " & GetFileName(SocketItem(u).FileName) & ")"
                      RemoveSocket u, 0, "", True
                    Else
                      If ((u2 = 0) And Not (IsIgnored(FullAddress))) Or (u2 > 0) Then
                        If SocketItem(u).BytesReceived = 0 Then
                          SocketItem(u).BytesReceived = u2
                          SpreadLevelFileAreaMessage 0, "14[" & Time & "] *** DCC send to " & SocketItem(u).RegNick & " (" & GetFileName(SocketItem(u).FileName) & ") resuming at " & SizeToString(SocketItem(u).BytesReceived)
                          SendLine "PRIVMSG " & Nick & " :DCC winsock2_accept " & Param(Line, 6) & " " & Param(Line, 7) & " " & Param(Line, 8), 1
                          If u2 = 0 Then AddIgnore Mask(FullAddress, 2), 6, 1
                        Else
                          SpreadLevelFileAreaMessage 0, "14[" & Time & "] *** DCC send to " & SocketItem(u).RegNick & " aborted (duplicate RESUME: " & GetFileName(SocketItem(u).FileName) & ")"
                          RemoveSocket u, 0, "", True
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Next u
            If Err.Number > 0 Then Err.Clear
            On Error GoTo 0
          End If
          'DCC send request
          If UCase(Param(Line, 5)) = "SEND" Then
            If RegUser = "" Then
              If IgnoreCheck(FullAddress, 2, 15) Then SpreadLevelFileAreaMessage 0, "14[" & Time & "] *** DCC Get from " & Nick & " ignored (user unknown)"
              Exit Sub
            End If
            If Not IsIdentified(Nick, RegUser) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (not identified)"
                If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept files from identified users.", 3
              Exit Sub
            End If
            If LCase(Right(Param(Line, 6), 5)) = ".seen" Then
              If MatchFlags(AllFlags(RegUser), "-n") Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get of a seen list from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (+n needed)"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept seen lists from owners or channel owners.", 3
                End If
                Exit Sub
              End If
            Else
              If MatchFlags(UserFlags, "-i") Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (+i needed)"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept files from +i users.", 3
                End If
                Exit Sub
              End If
            End If
            'Superowner-only sends
            If MatchFlags(UserFlags, "-s") Then
              If LCase(GetFileName(Param(Line, 6))) = "angel.exe" Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (+s needed for update)"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept update files from super owners.", 3
                End If
                Exit Sub
              End If
              If LCase(GetFileName(Param(Line, 6))) = "motd.txt" Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (+s needed for MOTD change)"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept MOTD files from super owners.", 3
                End If
                Exit Sub
              End If
              If LCase(Right(GetFileName(Param(Line, 6)), 4)) = ".asc" Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (+s needed for script upload)"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept scripts from super owners.", 3
                End If
                Exit Sub
              End If
              If LCase(Right(GetFileName(Param(Line, 6)), 4)) = ".lng" Then
                If IgnoreCheck(FullAddress, 2, 15) Then
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** DCC Get from " & Nick + IIf(RegUser <> Nick, " (" & RegUser & ")", "") & " refused (+s needed for language file upload)"
                  If HideBot = False Then SendLine "notice " & Nick & " :Sorry, I only winsock2_accept language files from super owners.", 3
                End If
                Exit Sub
              End If
            End If
            IP = Param(Line, 7)
            Port = Param(Line, 8)
            On Local Error Resume Next
            If CLng(Port) < 1024 Or CLng(Port) > 63000 Then Exit Sub
            If Err.Number <> 0 Then
              Err.Clear
              Exit Sub
            End If
            On Error GoTo 0
            NewSock = AddSocket
            SocketItem(NewSock).RegNick = RegUser
            SocketItem(NewSock).IRCNick = Nick
            SocketItem(NewSock).Hostmask = FullAddress
            SocketItem(NewSock).FileSize = Left(Param(Line, 9), Len(Param(Line, 9)) - 1)
            SocketItem(NewSock).FileName = GetFileName(Param(Line, 6))
            If InStr(SocketItem(NewSock).FileName, "_") > 0 Then
              SocketItem(NewSock).FileName = Trim(RemUnderscore(SocketItem(NewSock).FileName))
              If Left(SocketItem(NewSock).FileName, 1) = "." Or SocketItem(NewSock).FileName = "" Then
                SocketItem(NewSock).FileName = GetFileName(Param(Line, 6))
              End If
            End If
            SetSockFlag NewSock, SF_Status, SF_Status_FileWaiting
            SetSockFlag NewSock, SF_Echo, SF_NO
            SocketItem(NewSock).OnBot = BotNetNick
            SocketItem(NewSock).BytesReceived = 0
            SocketItem(NewSock).PLChannel = 0
            HaltDefault = False
            For ScNum = 1 To ScriptCount
              If Scripts(ScNum).Hooks.fa_uploadbegin Then
                RunScriptX ScNum, "fa_uploadbegin", Nick, RegUser, SocketItem(NewSock).FileName, SocketItem(NewSock).FileSize
              End If
            Next ScNum
            If HaltDefault = True Then RemoveSocket NewSock, 0, "", True: Exit Sub
            SpreadLevelFileAreaMessage u, "14[" & Time & "] *** DCC Get from " & SocketItem(NewSock).RegNick & " starting  (" & SocketItem(NewSock).FileName & ")..."
            If Err.Number = 0 Then
              If ConnectTCP(NewSock, IrcGetAscIp(IP), CLng(Port)) <> 0 Then
                RemoveSocket NewSock, 0, "", True
              End If
            Else
              RemoveSocket NewSock, 0, "", True
            End If
          End If
      'CTCP Chat request
      Case "chat", "chat"
          If RegUser = "" Then
            If UnIgnoreTimes = 0 Then
              If IgnoreCheck(FullAddress, 2, 15) Then
                CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being CTCP flooded. Ignoring all requests for 2 minutes.": Exit Sub
                SpreadFlagMessage 0, "+m", "14[" & Time & "] CTCP CHAT from " & Nick & " (" & Mask(FullAddress, 2) & ") refused"
                If HideBot = False Then
                  SendLine "notice " & Nick & " :Sorry, I don't chat with strangers.", 3
                  SendLine "notice " & Nick & " :If I should know you, type '/msg " & MyNick & " <ident command> <password>'.", 3
                End If
              End If
            End If
            Exit Sub
          End If
          If MatchFlags(UserFlags, "-p") Then
            If UnIgnoreTimes = 0 Then
              If IgnoreCheck(FullAddress, 2, 15) Then
                CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being CTCP flooded. Ignoring all requests for 2 minutes.": Exit Sub
                SpreadFlagMessage 0, "+m", "14[" & Time & "] CTCP CHAT from " & Nick & " (" & Mask(FullAddress, 2) & ") refused"
                If HideBot = False Then SendLine "notice " & Nick & " :" & MakeMsg(ERR_Login_NoChat), 3
              End If
            End If
            Exit Sub
          End If
          SpreadFlagMessageEx u, "+m", SF_Local_JP, "14[" & Time & "] *** CTCP CHAT request from " & Nick & "..."
          InitiateDCCChat Nick, FullAddress, RegUser, UserFlags
      'CTCP VERSION request
      Case "version", "version"
          If UnIgnoreTimes = 0 Then
            If IgnoreCheck(FullAddress, 2, 15) Then
              CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being CTCP flooded. Ignoring all requests for 2 minutes.": Exit Sub
              SpreadFlagMessage 0, "+m", "14[" & Time & "] CTCP VERSION from " & Nick & " (" & Mask(FullAddress, 2) & ")" & IIf(VersionReply = "°NONE", " (ignored)", "")
              If VersionReply <> "°NONE" Then SendLine "notice " & Nick & " :VERSION " & VersionReply & "", 3
            End If
          End If
      Case "ping", "ping"
          If UnIgnoreTimes = 0 Then
            If IgnoreCheck(FullAddress, 2, 15) Then
              CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being CTCP flooded. Ignoring all requests for 2 minutes.": Exit Sub
              SpreadFlagMessage 0, "+m", "14[" & Time & "] CTCP PING from " & Nick & " (" & Mask(FullAddress, 2) & ")"
              Rest = Param(Line, 5): If Right(Rest, 1) <> "" Then Rest = Rest & ""
              SendLine "notice " & Nick & " :PING " & Rest, 2
            End If
          End If
      'Ignore Finger, Time, Clientinfo etc.
      Case "finger", "time", "clientinfo", "userinfo", "finger", "time", "clientinfo", "userinfo"
          If UnIgnoreTimes = 0 Then
            If IgnoreCheck(FullAddress, 2, 5) Then
              CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being CTCP flooded. Ignoring all requests for 2 minutes.": Exit Sub
            End If
          End If
      '                    4     5     6       7    8
      'Loona->|^AnGeL^|: ANGEL LINK REQ/REQ! Loona ...
      Case "angel"
          Select Case LCase(Param(Line, 5))
            Case "link"
              Select Case LCase(Param(Line, 6))
                'Link request
                Case "req", "req!"
                  Rest = Param(Line, 7)
                  HostOnly = Param(Line, 8)
                  If Right(HostOnly, 1) = "" Then HostOnly = Left(HostOnly, Len(HostOnly) - 1)
                  UsNum = GetUserNum(Rest)
                  If UsNum = 0 Then UsNum = GetUserNum(RegUser)
                  If UsNum = 0 Then
                    If IgnoreCheck(FullAddress, 2, 10) Then
                      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: Unknown bot"
                      SendLine "privmsg " & Nick & " :ANGEL LINK ERR This bot doesn't know me.", 3
                    End If
                    Exit Sub
                  End If
                  If MatchFlags(BotUsers(UsNum).Flags, "-b") Then
                    If IgnoreCheck(FullAddress, 2, 10) Then
                      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(RegUser), " (" & RegUser & ")", "") & " failed LINK: Not a bot!"
                      SendLine "privmsg " & Nick & " :ANGEL LINK ERR I'm not added as a bot there.", 3
                    End If
                    Exit Sub
                  End If
                  If BotUsers(UsNum).Password <> "" Then
                    If LCase(Param(Line, 6)) = "req" Then
                      If IgnoreCheck(FullAddress, 2, 10) Then
                        SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: Password required"
                        SendLine "privmsg " & Nick & " :ANGEL LINK ERR This bot wants a password but I don't have one.", 3
                      End If
                      Exit Sub
                    Else
                      HostOnly = DecryptString(HostOnly, EncryptIt(BotUsers(UsNum).Password))
                      If ParamCount(HostOnly) <> 3 Then
                        If IgnoreCheck(FullAddress, 2, 10) Then
                          SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: Wrong password"
                          SendLine "privmsg " & Nick & " :ANGEL LINK ERR Our passwords don't match.", 3
                        End If
                        Exit Sub
                      End If
                      HostOnly = Param(HostOnly, 2)
                    End If
                  Else
                    If LCase(Param(Line, 6)) = "req!" Then
                      If IgnoreCheck(FullAddress, 2, 10) Then
                        SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: I don't know the password!"
                        SendLine "privmsg " & Nick & " :ANGEL LINK ERR I want a password but this bot doesn't have one.", 3
                      End If
                      Exit Sub
                    End If
                  End If
                  If (RegUser <> "") And (RegUser <> Rest) Then
                    If GetUserNum(Rest) > 0 Then
                      If IgnoreCheck(FullAddress, 2, 10) Then
                        SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: I know this bot as """ & RegUser & """!"
                        SendLine "privmsg " & Nick & " :ANGEL LINK ERR This bot knows me as """ & RegUser & """, not as " & Rest & ".", 3
                      End If
                      Exit Sub
                    End If
                    ChangeNick RegUser, Rest
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Changed nick of " & RegUser & " to " & Rest
                  End If
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: Received " & Rest & "'s link address (" & HostOnly & ") - sending mine"
                  SetUserData UsNum, UD_LinkAddr, HostOnly
                  If BotUsers(UsNum).Password <> "" Then
                    SendLine "privmsg " & Nick & " :ANGEL LINK OK! " & BotNetNick & " " & EncryptString(RandString & " " & IrcGetAscIp(MyIP) & ":" & Trim(Str(BotnetPort)) & " " & RandString, EncryptIt(BotUsers(UsNum).Password)) & "", 2
                  Else
                    SendLine "privmsg " & Nick & " :ANGEL LINK OK " & BotNetNick & " " & IrcGetAscIp(MyIP) & ":" & Trim(Str(BotnetPort)) & "", 2
                  End If
                Case "err"
                  If IgnoreCheck(FullAddress, 2, 3) Then
                    Rest = GetRest(Line, 7)
                    If Right(Rest, 1) = "" Then Rest = Left(Rest, Len(Rest) - 1)
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Couldn't link to " & IIf((RegUser <> ""), RegUser, Nick) & ": " & Rest
                  End If
                Case "ok", "ok!"
                  Rest = Param(Line, 7)
                  HostOnly = Param(Line, 8)
                  If Right(HostOnly, 1) = "" Then HostOnly = Left(HostOnly, Len(HostOnly) - 1)
                  UsNum = GetUserNum(Rest)
                  If UsNum = 0 Then UsNum = GetUserNum(RegUser)
                  If UsNum = 0 Then Exit Sub
                  If MatchFlags(BotUsers(UsNum).Flags, "-b") Then Exit Sub
                  If BotUsers(UsNum).Password <> "" Then
                    If LCase(Param(Line, 6)) = "ok" Then
                      If IgnoreCheck(FullAddress, 2, 10) Then
                        SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: Password required"
                        SendLine "privmsg " & Nick & " :ANGEL LINK ERR This bot wants a password but I don't have one.", 3
                      End If
                      Exit Sub
                    Else
                      HostOnly = DecryptString(HostOnly, EncryptIt(BotUsers(UsNum).Password))
                      If ParamCount(HostOnly) <> 3 Then
                        If IgnoreCheck(FullAddress, 2, 10) Then
                          SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: Wrong password"
                          SendLine "privmsg " & Nick & " :ANGEL LINK ERR Our passwords don't match.", 3
                        End If
                        Exit Sub
                      End If
                      HostOnly = Param(HostOnly, 2)
                    End If
                  Else
                    If LCase(Param(Line, 6)) = "ok!" Then
                      If IgnoreCheck(FullAddress, 2, 10) Then
                        SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: I don't know the password!"
                        SendLine "privmsg " & Nick & " :ANGEL LINK ERR I want a password but this bot doesn't have one.", 3
                      End If
                      Exit Sub
                    End If
                  End If
                  If (RegUser <> "") And (RegUser <> Rest) Then
                    If GetUserNum(Rest) > 0 Then
                      If IgnoreCheck(FullAddress, 2, 10) Then
                        SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Botnet: " & Nick + IIf(LCase(Nick) <> LCase(Rest), " (" & Rest & ")", "") & " failed LINK: I know this bot as """ & RegUser & """!"
                        SendLine "privmsg " & Nick & " :ANGEL LINK ERR This bot knows me as """ & RegUser & """, not as " & Rest & ".", 3
                      End If
                      Exit Sub
                    End If
                    ChangeNick RegUser, Rest
                    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Changed nick of " & RegUser & " to " & Rest
                  End If
                  SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLBotNetLinking, Rest, "Received link address (" & HostOnly & ")")
                  SetUserData UsNum, UD_LinkAddr, HostOnly
                  InitiateBotChat UsNum, False
              End Select
          End Select
      Case "seen", CommandPrefix & "seen"
          HaltDefault = False
          For ScNum = 1 To ScriptCount
            If Scripts(ScNum).Hooks.seen Then
              RunScriptX ScNum, "seen", Nick, RegUser, Replace(GetRest(Line, 5), ", ", ",")
            End If
          Next ScNum
          If HaltDefault = True Then Exit Sub
          If UnIgnoreTimes = 0 Then
            If Not IsIgnored(FullAddress) Then
              AddIgnore Mask(FullAddress, 2), 20, 1
            Else
              CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being seen flooded. Ignoring all messages for 2 minutes.": Exit Sub
              UsNum = GetIgnoreLevel(FullAddress)
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel FullAddress, UsNum + 1
            End If
            If (HideBot = True) And (RegUser = "") Then
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *" & Nick & "* " & CommandPrefix & "seen " & Param(Line, 5) & "  (ignored - unknown user)"
            Else
              Rest = LastSeen(Param(Replace(GetRest(Line, 5), ", ", ","), 1), Nick, "", RegUser, MatchedOne)
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *" & Nick & "* " & CommandPrefix & "seen " & Param(Line, 5) + IIf(Not MatchedOne, "  (not seen)", "")
              If Rest <> "" Then SendLine "privmsg " & Nick & " :" & Rest, 3
            End If
          End If
      Case CommandPrefix & "whois", "whois"
          If UnIgnoreTimes = 0 Then
            If Not IsIgnored(FullAddress) Then
              AddIgnore Mask(FullAddress, 2), 20, 1
            Else
              CountCTCPs = CountCTCPs + 1: If CountCTCPs = 5 Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being whois flooded. Ignoring all messages for 2 minutes.": Exit Sub
              UsNum = GetIgnoreLevel(FullAddress)
              If UsNum > 3 Then Exit Sub
              SetIgnoreLevel FullAddress, UsNum + 1
            End If
            If (HideBot = True) And (RegUser = "") Then
              If LCase(Param(Line, 5)) <> "me" Then If Left(Param(Line, 5), 1) = "@" Then RegUser = Right(Param(Line, 5), Len(Param(Line, 5)) - 1) Else RegUser = Param(Line, 5)
              If Len(RegUser) <= ServerNickLen And IsValidNick(RegUser) Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *" & Nick & "* !whois " & Param(Line, 5) & "  (ignored - unknown user)"
              End If
            Else
              If LCase(Param(Line, 5)) <> "me" Then If Left(Param(Line, 5), 1) = "@" Then RegUser = Right(Param(Line, 5), Len(Param(Line, 5)) - 1) Else RegUser = Param(Line, 5)
              If Len(RegUser) <= ServerNickLen And IsValidNick(RegUser) Then
                UsNum = GetUserNum(RegUser)
                If UsNum = 0 Then
                  If LCase(RegUser) = "me" Then
                    SendLine "privmsg " & Nick & " :Sorry, I don't know who you are.", 3
                  Else
                    If Left(Param(Line, 5), 1) = "@" Then RegUser = Right(Param(Line, 5), Len(Param(Line, 5)) - 1) Else RegUser = Param(Line, 5)
                    SendLine "privmsg " & Nick & " :Sorry, I don't know who " & RegUser & " is.", 3
                  End If
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *" & Nick & "* !whois " & Param(Line, 5) & "  (unknown user)"
                  Exit Sub
                End If
                Rest = GetUserData(UsNum, "info", "")
                If Rest = "" Then
                  SendLine "privmsg " & Nick & " :" & BotUsers(UsNum).Name & " has no information set.", 3
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *" & Nick & "* !whois " & Param(Line, 5) & "  (no info set)"
                  Exit Sub
                End If
                If Rest = LastWhoisOutput Then
                  If Not DontSeen Then
                    Select Case Int(Rnd * 3) + 1
                      Case 1: SendLine "privmsg " & Nick & " : I don't like to repeat myself...", 3
                      Case 2: SendLine "privmsg " & Nick & " : I already said that!", 3
                      Case 3: SendLine "privmsg " & Nick & " : Look some lines above *^^*", 3
                    End Select
                  End If
                  DontSeen = True
                Else
                  SendLine "privmsg " & Nick & " :" & BotUsers(UsNum).Name & ": " & Rest, 3
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] *" & Nick & "* !whois " & Param(Line, 5)
                  LastWhoisOutput = Rest
                End If
              Else
                If RegUser <> "" Then SendLine "privmsg " & Nick & " :This user doesn't exist.", 3
              End If
            End If
          End If
      Case Else
          If UnIgnoreTimes = 0 Then
            CountCTCPs = CountCTCPs + 1: If CountCTCPs = MaxPrivEvents Then UnIgnoreTimes = 2: SpreadFlagMessage 0, "+m", "4*** I'm being Query flooded. Ignoring all queries for 2 minutes.": Exit Sub
            If Not IsIgnored(FullAddress) Then
              AddIgnore Mask(FullAddress, 2), 20, 1
            Else
              UsNum = GetIgnoreLevel(FullAddress)
              SetIgnoreLevel FullAddress, UsNum + 1
              If UsNum = MaxUserEvents Then
                SpreadFlagMessage 0, "+m", "14[" & Time & "] Ignoring " & Mask(FullAddress, 2) & " for 2 min: Flood."
                AddIgnore Mask(FullAddress, 2), 120, 6
                Exit Sub
              ElseIf UsNum > MaxUserEvents Then
                Exit Sub
              End If
            End If
            
            Rest = StripDP(GetRest(Line, 4))
            If LCase(Nick) <> LCase(MyNick) Then
              If LCase(Param(Rest, 1)) <> "action" Then
                If Left(Rest, 1) <> "" Then
                  If RegUser <> "" Then
                    If BotUsers(GetUserNum(RegUser)).Password <> "" Then
                      If EncryptIt(Param(Rest, 2)) = BotUsers(GetUserNum(RegUser)).Password Then
                        BotUsers(GetUserNum(RegUser)).ValidSession = True
                        ValidSession = True
                        SendLine "NOTICE " & Nick & " :Shhht! Please use your password ONLY in connection with the correct commands! Your message was not shown on the party line.", 3
                        Exit Sub
                      End If
                    End If
                  Else
                    If BotUsers(GetUserNum(Nick)).Password <> "" Then
                      If EncryptIt(Param(Rest, 2)) = BotUsers(GetUserNum(Nick)).Password Then
                        BotUsers(GetUserNum(RegUser)).ValidSession = True
                        ValidSession = True
                        SendLine "NOTICE " & Nick & " :Shh!! Please use your password ONLY in connection with the correct commands! Your message was not shown on the party line.", 3
                        Exit Sub
                      End If
                    End If
                  End If
                  SpreadFlagMessage 0, "+m", "14[" & Time & "] PRIVMSG from " & Nick & ": " & Rest
                  If HideBot = False Then
                    Rest = MakeReply(Rest, Nick)
                    If Rest <> "" Then TimedEvent "KIAnswer " & Nick & " " & Rest, 2 + Int(Len(ParamX(Rest, "|", 1)) / 6) + Int(Rnd * Len(ParamX(Rest, "|", 1)) / 10)
                  End If
                End If
              Else
                SpreadFlagMessage 0, "+m", "14[" & Time & "] PRIVMSG from " & Nick & ": * " & Nick & " " & Mid(Rest, 9, Len(Rest) - 9)
              End If
            End If
          End If
    End Select
  End If
  If ValidSession = True Then
    BotUsers(GetUserNum(RegUser)).OldFullAddress = Mid(FullAddress, InStr(FullAddress, "@") + 1)
  End If
End Sub

