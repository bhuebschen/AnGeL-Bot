Attribute VB_Name = "Botnet_Parser"
',-======================- ==-- -  -
'|   AnGeL - Botnet - Parser
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

'Handle messages received via the botnet
Public Sub Botnet(vsock As Long, Line As String) ' : AddStack "BotNet_Botnet(" & vsock & ", " & SockNum & ", " & Line & ")"
Dim Nick As String, Flags As String, UserNum As Long, u As Long, u2 As Long, OnBot As String
Dim NBefore As String, FoundOne As Boolean, Message As String, ToBot As String
Dim SNum As Long, ToNick As String, Rest As String, OrderSign As String
Dim ScNum As Long
If vsock > SocketCount Then Exit Sub
If Not IsValidSocket(vsock) Then Exit Sub
Nick = SocketItem(vsock).RegNick
Flags = SocketItem(vsock).Flags
UserNum = SocketItem(vsock).UserNum
  'Output Line + vbCrLf
  
  Select Case GetSockFlag(vsock, SF_Status)
    Case SF_Status_BotLinking
        If GetSockFlag(vsock, SF_LoggedIn) = SF_NO Then
            'If something is sent via the connection - send user/pass...
            If BotUsers(UserNum).Password = "" Then
              RTU vsock, BotNetNick
              SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinking, SocketItem(vsock).RegNick, MakeMsg(MSG_BNSendUser))
            Else
              RTU vsock, BotNetNick
              RTU vsock, BotUsers(UserNum).Password
              SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinking, SocketItem(vsock).RegNick, MakeMsg(MSG_BNSendUserPass))
            End If
            SetSockFlag vsock, SF_LoggedIn, SF_YES
            SetSockFlag vsock, SF_Silent, SF_YES
        Else
            Select Case LCase(Param(Line, 1))
              'Password was accepted
              Case "*hello!"
                  RTU vsock, "version " & LongBotVersion & " " & CStr(ServerNickLen) & " AnGeL " & BotVersion & " <" & ServerNetwork & ">"
              Case "version"
                  If StrictDupeCheck(vsock) Then Exit Sub
                  On Local Error Resume Next
                  SocketItem(vsock).SetupChan = LongToBase64(Val(Param(Line, 2)))
                  If Err.Number <> 0 Then
                    SocketItem(vsock).SetupChan = ""
                    Err.Clear
                  End If
                  On Error GoTo 0
                  LinkCheck vsock
                  PreBotNetLogin vsock
              Case "handshake"
                  'Only allow handshakes if I'm connecting to other bot
                  If (Param(Line, 2) <> "") And (SocketItem(vsock).SocketDirection = SD_Out) Then
                    BotUsers(UserNum).Password = Param(Line, 2)
                  End If
              Case "passreq"
                  If BotUsers(UserNum).Password = "" Then
                    SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinking, SocketItem(vsock).RegNick, MakeMsg(MSG_BNNoPass))
                    RTU vsock, "-"
                    RemoveSocket vsock, 0, "Password needed:", False
                  End If
              Case "badpass"
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinking, SocketItem(vsock).RegNick, MakeMsg(MSG_BNBadPass))
              Case "you"
                  If InStr(Line, "don't have access") > 0 Then SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinking, SocketItem(vsock).RegNick, MakeMsg(MSG_BNNoAccess))
              Case "error"
                  If Param(Line, 2) <> "" Then SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetError, SocketItem(vsock).RegNick, GetRest(Line, 2))
            End Select
        End If
    Case SF_Status_BotPreCache
        SocketItem(vsock).CurrentQuestion = SocketItem(vsock).CurrentQuestion + Chr(0) + Line
        Select Case LCase(Param(Line, 1))
          'Syntax: tb <bot nick>
          Case "tb", "thisbot"
              If Param(Line, 2) = "" Then Exit Sub
              If Param(Line, 2) <> SocketItem(vsock).RegNick Then
                If LCase(Param(Line, 2)) <> LCase(SocketItem(vsock).RegNick) Then
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLWrongBot, SocketItem(vsock).RegNick, Param(Line, 2))
                  RemoveSocket vsock, 0, MakeMsg(MSG_BNWrongBot, SocketItem(vsock).RegNick, Param(Line, 2)), True
                  Exit Sub
                End If
              End If
          'Syntax: n <new bot> <linked by> <version>
          Case "n", "nlinked"
              'Loop check
              If GetBotPos(Param(Line, 2)) > 0 Then
                RTU vsock, "error Loop detected ('" & Param(Line, 2) & "' is already connected!)"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Loop from " & Nick & " detected ('" & Param(Line, 2) & "' is already connected!)"
                RemoveSocket vsock, 0, MakeMsg(MSG_BNLoop, Nick, Param(Line, 2)), True
                Exit Sub
              End If
              'Leaf link check
              If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+l") Then
                RTU vsock, "error You can't link to me while you're connected to other bots."
                RTU vsock, "bye Unauthorized links"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Leaf bot " & Nick & " has other bots linked!"
                RemoveSocket vsock, 0, MakeMsg(MSG_BNLeafLinks), True
                Exit Sub
              End If
              'Bogus botname check
              If Not IsValidNick(Param(Line, 2)) Then
                RTU vsock, "error You have a bogus bot connected: " & Param(Line, 2)
                RTU vsock, "bye Bogus link"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: " & Nick & " has a bogus bot connected: " & Param(Line, 2)
                RemoveSocket vsock, 0, MakeMsg(MSG_BNBogusLink), True
                Exit Sub
              End If
          Case "error"
              If Param(Line, 2) <> "" Then SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetError, SocketItem(vsock).RegNick, GetRest(Line, 2))
          Case "bye", "*bye"
              If GetRest(Line, 2) <> "" Then
                If (SocketItem(vsock).SocketDirection = SD_Out) Then
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinking, SocketItem(vsock).RegNick, "Disconnected (" & GetRest(Line, 2) & ")")
                Else
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Disconnected (" & GetRest(Line, 2) & ")")
                End If
              End If
              RemoveSocket vsock, 0, "Disconnected while precaching", True
          Case "el"
              FinalBotNetLogin vsock, SocketItem(vsock).OrderSign
        End Select
    Case SF_Status_Bot
        HaltDefault = False
        For ScNum = 1 To ScriptCount
          If Scripts(ScNum).Hooks.Botnet = True Then
            RunScriptX ScNum, "bn", Line
          End If
        Next ScNum
        If HaltDefault = True Then Exit Sub
        Select Case LCase(Param(Line, 1))
          Case "handshake"
              'Only allow handshakes if I'm connecting to other bot
              If (Param(Line, 2) <> "") And (SocketItem(vsock).SocketDirection = SD_Out) Then
                BotUsers(UserNum).Password = Param(Line, 2)
              End If
          Case "bye", "*bye"
              RemoveSocket vsock, 0, MakeMsg(MSG_BNDisconnect, IIf(GetRest(Line, 2) <> "", "³" & GetRest(Line, 2), "")), False
          'Syntax: tb <bot nick>
          Case "tb", "thisbot"
              If Param(Line, 2) = "" Then Exit Sub
              If Param(Line, 2) <> SocketItem(vsock).RegNick Then
                If LCase(Param(Line, 2)) = LCase(SocketItem(vsock).RegNick) Then
                  ChangeNick SocketItem(vsock).RegNick, Param(Line, 2)
                  For u = 1 To BotCount
                    If LCase(Bots(u).Nick) = LCase(SocketItem(vsock).RegNick) Then Bots(u).Nick = Param(Line, 2)
                  Next u
                  SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Changed nick of " & SocketItem(vsock).RegNick & " to " & Param(Line, 2)
                  SocketItem(vsock).RegNick = Param(Line, 2)
                Else
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLWrongBot, SocketItem(vsock).RegNick, Param(Line, 2))
                  RemoveSocket vsock, 0, MakeMsg(MSG_BNWrongBot, SocketItem(vsock).RegNick, Param(Line, 2)), True
                  Exit Sub
                End If
              End If
          'Syntax: s <data>
          Case "s"
              SharingMessage vsock, Line
          'Syntac: t <ttl>:<user>@<source> <target> :<time>:<bot1>:<bot2>
          Case "t"
              TraceMessage vsock, Line
          Case "td"
              TraceReply vsock, Line
          'Syntax: n <new bot> <linked by> <version>
          Case "n", "nlinked"
              If Param(Line, 3) = "" Then
                RTU vsock, "error Missing parameters: " & Line
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Missing parameters from " & Nick & ": " & Line
                Exit Sub
              End If
              'Loop check
              If GetBotPos(Param(Line, 2)) > 0 Then
                RTU vsock, "error Loop detected ('" & Param(Line, 2) & "' is already connected!)"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Loop from " & Nick & " detected ('" & Param(Line, 2) & "' is already connected!)"
                RemoveSocket vsock, 0, MakeMsg(MSG_BNLoop, Nick, Param(Line, 2)), True
                Exit Sub
              End If
              If FakeCheck("", Param(Line, 3), vsock) Then Exit Sub
              'Leaf link check
              If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+l") Then
                RTU vsock, "error As a leaf bot, you're not allowed to link to other bots!"
                RTU vsock, "bye Unauthorized links"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Leaf bot " & Nick & " tried to link to another bot ('" & Param(Line, 2) & "')!"
                RemoveSocket vsock, 0, MakeMsg(MSG_BNLeafLink, Param(Line, 2)), True
                Exit Sub
              End If
              
              'Bogus botname check
              If Not IsValidNick(Param(Line, 2)) Then
                RTU vsock, "error Bogus link: " & Param(Line, 2)
                RTU vsock, "bye Bogus link"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: " & Nick & " tried to link to a bogus bot: " & Param(Line, 2)
                RemoveSocket vsock, 0, MakeMsg(MSG_BNBogusLink), True
                Exit Sub
              End If
              
              ToNick = Param(Line, 4)
              If Left(ToNick, 1) = "!" Then ToNick = Right(ToNick, Len(ToNick) - 1)
              NBefore = "": If Left(ToNick, 1) = "-" Or Left(ToNick, 1) = "+" Then NBefore = Left(ToNick, 1): ToNick = Right(ToNick, Len(ToNick) - 1)
              AddBot Param(Line, 2), Param(Line, 3), NBefore, ToNick, 0
              If Left(Param(Line, 4), 1) = "!" Then
                If GetSockFlag(vsock, SF_Silent) = SF_YES Then
                  ToBotNet vsock, Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3) & " -" & ToNick
                Else
                  ToBotNet vsock, Line
                  Nick = Param(Line, 3)
                  Message = MakeMsg(MSG_BNConnect, Param(Line, 2))
                  If Nick <> BotNetNick Then SpreadMessage 0, -1, "3*** 14(" & Nick & ")3 " & Message
                End If
              Else
                ToBotNet vsock, Line
              End If
          'Syntax: j <bot> <user> <channel> <user flags>
          Case "j", "join"
              OnBot = Param(Line, 2)
              If FakeCheck("", OnBot, vsock) Then Exit Sub
              If Left(OnBot, 1) <> "!" And GetSockFlag(vsock, SF_Silent) = SF_YES Then
                Message = "j !" & Trim(Right(Line, Len(Line) - Len(Param(Line, 1))))
                ToBotNet vsock, Message
              Else
                ToBotNet vsock, Line
              End If
              If AddBotNetUser(Line) Then
                If Left(OnBot, 1) <> "!" And GetSockFlag(vsock, SF_Silent) = SF_NO Then
                  If LCase(Param(Line, 1)) = "join" Then
                    SpreadMessage 0, CLng(Param(Line, 4)), MakeMsg(MSG_PLBotNetJoin, OnBot, Param(Line, 3))
                  Else
                    SpreadMessage 0, Base64ToLong(Param(Line, 4)), MakeMsg(MSG_PLBotNetJoin, OnBot, Param(Line, 3))
                  End If
                End If
              End If
          Case "un"
              If Param(Line, 2) = "" Then Exit Sub
              If FakeCheck("", Param(Line, 2), vsock) Then Exit Sub
              If LCase(Param(Line, 2)) = LCase(SocketItem(vsock).RegNick) Then
                RTU vsock, "error Fake message rejected (You can't unlink yourself!)"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Fake message from " & SocketItem(vsock).RegNick & " rejected (bot can't unlink itself)"
                Exit Sub
              End If
              ToBotNet vsock, Line
              For u = 2 To BotCount
                If LCase(Param(Line, 2)) = LCase(Bots(u).Nick) Then OnBot = Bots(u).SubBotOf: Exit For
              Next u
              If Param(Line, 3) <> "" Then SpreadMessage 0, -1, "3*** 14(" & OnBot & ")3 " & GetRest(Line, 3)
              RemBot Param(Line, 2), 0, ""
          'Idle time - Syntax: i <bot> <usernum in base64> <idle time in base64> (away message)
          Case "i"
              OnBot = Param(Line, 2)
              If FakeCheck("", OnBot, vsock) Then Exit Sub
              ToBotNet vsock, Line
              ToNick = Param(Line, 3)
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If LCase(SocketItem(u).OnBot) = LCase(OnBot) And SocketItem(u).OrderSign = ToNick Then
                    On Local Error Resume Next
                    SocketItem(u).LastEvent = CDate(Now - CDate(Base64ToLong(Param(Line, 4)) / 86400))
                    If Err.Number <> 0 Then
                      SocketItem(u).LastEvent = Now
                      Err.Clear
                    End If
                    SocketItem(u).AwayMessage = GetRest(Line, 5)
                    On Error GoTo 0
                    Exit For
                  End If
                End If
              Next u
          Case "el"
              SetSockFlag vsock, SF_Silent, SF_NO
          'Away message - Syntax: aw <bot> <usernum is base64> (message)
          Case "aw"
              OnBot = Param(Line, 2)
              If FakeCheck("", OnBot, vsock) Then Exit Sub
              ToBotNet vsock, Line
              ToNick = Param(Line, 3)
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If (LCase(SocketItem(u).OnBot) = LCase(OnBot)) And (SocketItem(u).OrderSign = ToNick) Then
                    If GetRest(Line, 4) <> "" Then
                      SpreadMessage 0, SocketItem(u).PLChannel, MakeMsg(MSG_PLBotNetAway, SocketItem(u).OnBot, SocketItem(u).RegNick, GetRest(Line, 4))
                    Else
                      If SocketItem(u).AwayMessage <> "" Then SpreadMessage 0, SocketItem(u).PLChannel, MakeMsg(MSG_PLBotNetBack, SocketItem(u).OnBot, SocketItem(u).RegNick, SocketItem(u).AwayMessage)
                    End If
                    SocketItem(u).LastEvent = Now
                    SocketItem(u).AwayMessage = GetRest(Line, 4)
                    Exit For
                  End If
                End If
              Next u
          'Motd request - Syntax: m !<socknum>:<user>@<bot> <to bot>
          Case "m"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If InStr(Nick, ":") > 0 Then Nick = Right(Nick, Len(Nick) - InStr(Nick, ":"))
              If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              If LCase(Param(Line, 3)) = LCase(BotNetNick) Then
                ToNick = Param(Line, 2): If Left(ToNick, 1) = "!" Then ToNick = Right(ToNick, Len(ToNick) - 1)
                bShowMOTD OnBot, "p " & BotNetNick & " " & ToNick & " ", Nick & "@" & OnBot, GetUserFlags(Nick)
                SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** " & Nick & "@" & OnBot & " did .motd " & BotNetNick
              Else
                SendToBot Param(Line, 3), Line
              End If
          'Nick change
          Case "nc"
              ToBotNet vsock, Line
              For u = 1 To SocketCount
                If IsValidSocket(u) Then If SocketItem(u).OnBot = Param(Line, 2) And SocketItem(u).OrderSign = Param(Line, 3) Then NBefore = SocketItem(u).RegNick: SocketItem(u).RegNick = Param(Line, 4)
              Next u
              If NBefore = "" Then SpreadFlagMessage vsock, "+m", "14[" & Param(Line, 2) & "] *** Nick change error: User '" & Param(Line, 3) & "' renamed to " & Param(Line, 4): Exit Sub
              SpreadMessage vsock, -1, "14[" & Param(Line, 2) & "]3 *** " & NBefore & " is now known as " & Param(Line, 4)
          'User part
          Case "pt"
              OnBot = Param(Line, 2): If Left(OnBot, 1) = "!" Then OnBot = Mid(OnBot, 2)
              Nick = Param(Line, 3)
              OrderSign = Param(Line, 4)
              If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              ToBotNet vsock, Line
              If Not (OnBot = "!" Or OnBot = "") Then
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then
                    If (SocketItem(u).OnBot = OnBot) And (SocketItem(u).RegNick = Nick) And (SocketItem(u).OrderSign = OrderSign) Then
                      If Left(Param(Line, 2), 1) <> "!" Then
                        If (Param(Line, 5) = "") Or (GetRest(Line, 5) = Nick) Then
                          SpreadMessage 0, SocketItem(u).PLChannel, MakeMsg(MSG_PLBotNetLeave, OnBot, Nick)
                        Else
                          SpreadMessage 0, SocketItem(u).PLChannel, MakeMsg(MSG_PLBotNetLeaveMsg, OnBot, Nick, GetRest(Line, 5))
                        End If
                      End If
                      If GetRealNick(SocketItem(u).RegNick) <> "" Then
                        WriteSeenEntry SocketItem(u).RegNick, "", Now, SocketItem(u).OnBot, "*partyline*", Mask(SocketItem(u).Hostmask, 10)
                      Else
                        WriteExtSeenEntry SocketItem(u).RegNick, "", Now, SocketItem(u).OnBot, "*partyline*", Mask(SocketItem(u).Hostmask, 10)
                      End If
                      RemoveSocket u, 0, "", True
                      Exit For
                    End If
                  End If
                Next u
              End If
          'Syntax: c <from (user@)bot> <to channel (base64)> <message>
          Case "c"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              If MatchFlags(BotUsers(GetUserNum(OnBot)).BotFlags, "+i") = True Then Exit Sub
              ToBotNet vsock, Line
              If Param(Line, 4) <> "" Then
                Message = GetRest(Line, 4)
                If Nick <> "" Then
                  If UCase(Param(Message, 1)) = "ACTION" Then
                    Message = GetRest(Message, 2)
                    SpreadMessageEx 0, Base64ToLong(Param(Line, 3)), SF_Botnet_Talk, MakeMsg(MSG_PLBotNetAct, OnBot, Nick, Left(Message, Len(Message) - 1))
                  Else
                    SpreadMessageEx 0, Base64ToLong(Param(Line, 3)), SF_Botnet_Talk, MakeMsg(MSG_PLBotNetTalk, OnBot, Nick, Message)
                  End If
                Else
                  SpreadMessageEx 0, Base64ToLong(Param(Line, 3)), SF_Botnet_Bot, MakeMsg(MSG_PLBotTalk, OnBot, Message)
                End If
              End If
          'Syntax: chan <from (user@)bot> <to channel (long)> <message>
          Case "chan"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              ToBotNet vsock, Line
              If Param(Line, 4) <> "" Then
                Message = GetRest(Line, 4)
                If Nick <> "" Then
                  If UCase(Param(Message, 1)) = "ACTION" Then
                    Message = GetRest(Message, 2)
                    SpreadMessageEx 0, CLng(Param(Line, 3)), SF_Botnet_Talk, MakeMsg(MSG_PLBotNetAct, OnBot, Nick, Left(Message, Len(Message) - 1))
                  Else
                    SpreadMessageEx 0, CLng(Param(Line, 3)), SF_Botnet_Talk, MakeMsg(MSG_PLBotNetTalk, OnBot, Nick, Message)
                  End If
                Else
                  SpreadMessageEx 0, CLng(Param(Line, 3)), SF_Botnet_Bot, MakeMsg(MSG_PLBotTalk, OnBot, Message)
                End If
              End If
          'Syntax: ct <from bot> <message>
          Case "ct", "chat"
              Nick = Param(Line, 2)
              If FakeCheck("", Nick, vsock) Then Exit Sub
              ToBotNet vsock, Line
              If vsock <> u And Param(Line, 3) <> "" Then
                Message = GetRest(Line, 3)
                SpreadMessageEx 0, -1, SF_Botnet_Bot, MakeMsg(MSG_PLBotTalk, Nick, Message)
              End If
          'Syntax: a <from user@bot> <to channel> <message>
          Case "a"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              ToBotNet vsock, Line
              If Param(Line, 4) <> "" Then
                SpreadMessageEx 0, Base64ToLong(Param(Line, 3)), SF_Botnet_Talk, MakeMsg(MSG_PLBotNetAct, OnBot, Nick, GetRest(Line, 4))
              End If
          'Private message
          Case "p"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If Left(Nick, 2) <> "*:" Then If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              If InStr(Nick, ":") > 0 Then Nick = Mid(Nick, InStr(Nick, ":") + 1)
              ToNick = GetPartNick(Param(Line, 3))
              SNum = 0
              If InStr(ToNick, ":") > 0 Then
                SNum = CLng(Left(ToNick, InStr(ToNick, ":") - 1))
                ToNick = Right(ToNick, Len(ToNick) - InStr(ToNick, ":"))
              End If
              ToBot = GetPartBot(Param(Line, 3))
              If LCase(ToBot) <> LCase(BotNetNick) Then SendToBot ToBot, Line: Exit Sub
              If Param(Line, 4) <> "" Then
                Message = Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3)) - 1)
              Else
                Message = ""
              End If
              FoundOne = False
              If SNum = 0 Then
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then
                    If (LCase(SocketItem(u).RegNick) = LCase(ToNick)) And (GetSockFlag(u, SF_LocalVisibleUser) = SF_YES) Then
                      If Nick <> "" Then
                        If SocketItem(u).AwayMessage = "" Then
                          TU u, "14[" & OnBot & "]5 *" & Nick & "* " & Message
                          FoundOne = True
                        End If
                      Else
                        TU u, "14[" & OnBot & "] " & Message
                      End If
                    End If
                  End If
                Next u
              Else
                If IsValidSocket(SNum) Then
                  If (LCase(SocketItem(SNum).RegNick) = LCase(ToNick)) And (GetSockFlag(SNum, SF_LocalVisibleUser) = SF_YES) Then
                    If Nick <> "" Then
                      If SocketItem(SNum).AwayMessage = "" Then
                        TU SNum, "14[" & OnBot & "]5 *" & Nick & "* " & Message
                        FoundOne = True
                      End If
                    Else
                      TU SNum, "14[" & OnBot & "] " & Message
                    End If
                  End If
                End If
              End If
              If Nick <> "" Then
                If FoundOne Then
                  If Left(Param(Line, 2), 2) <> "*:" Then SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " Your note was delivered."
                Else
                  u = GetUserNum(ToNick)
                  If u = 0 Then
                    If Left(Param(Line, 2), 2) <> "*:" Then SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " Sorry, there's no user called '" & ToNick & "' in my userlist."
                  Else
                    ToNick = BotUsers(u).Name
                    If SendNote(Nick & "@" & OnBot, ToNick, "", Message) Then
                      If Left(Param(Line, 2), 2) <> "*:" Then SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " Your note was stored."
                    Else
                      If Left(Param(Line, 2), 2) <> "*:" Then SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " Sorry, this user's mailbox is full."
                    End If
                  End If
                End If
              End If
          'WHO request, Syntax: w <from user@bot> <to bot> <channel>
          Case "w"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If GetBotPos(OnBot) <> 0 Then
                If LCase(Param(Line, 3)) = LCase(BotNetNick) Then
                  ReplyWHO OnBot, Line
                Else
                  SendToBot Param(Line, 3), Line
                End If
              End If
          'Botinfo request, Syntax: i? <socknum>:<from user@bot>
          Case "i?"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If InStr(Nick, ":") > 0 Then Nick = Right(Nick, Len(Nick) - InStr(Nick, ":"))
              If LCase(Bots(GetNextBot(OnBot)).Nick) = LCase(SocketItem(vsock).RegNick) Then
                SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " AnGeL " & BotVersion & IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", " <IRCnet>") & " (" & BotChannels & ") [uptime: " & TimeSince(CStr(StartUpTime)) & "]"
                ToBotNet vsock, Line
              End If
          Case "pi", "ping"
              RTU vsock, "po"
          Case "error"
              If Param(Line, 2) <> "" Then SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetError, SocketItem(vsock).RegNick, GetRest(Line, 2))
          Case "po"
              Bots(GetBotPos(Nick)).GotPingReply = True
          Case "ul"
              Nick = GetPartNick(Param(Line, 2))
              OnBot = GetPartBot(Param(Line, 2))
              If InStr(Nick, ":") > 0 Then Nick = Right(Nick, Len(Nick) - InStr(Nick, ":"))
              If FakeCheck(Nick, OnBot, vsock) Then Exit Sub
              If LCase(Param(Line, 3)) = LCase(BotNetNick) Then
                ToNick = Param(Line, 4)
                If MatchFlags(GetUserFlags(Nick), "-t") Then
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLNickFailed, Nick & "@" & OnBot, ".unlink " & ToNick)
                  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " *** Sorry, you need the +t flag here to unlink bots."
                  Exit Sub
                Else
                  SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLNickDid, Nick & "@" & OnBot, ".unlink " & ToNick)
                End If
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then
                    If LCase(SocketItem(u).RegNick) = LCase(ToNick) Then
                      Select Case GetSockFlag(u, SF_Status)
                        Case SF_Status_Bot, SF_Status_BotLinking
                          RTU u, "bye"
                          Rest = GetRest(Line, 5)
                          RemoveSocket u, 0, "Unlinked from" & IIf(Rest <> "", "³" & Rest, ""), False
                          FoundOne = True
                      End Select
                    End If
                  End If
                Next u
              Else
                SendToBot Param(Line, 3), Line
              End If
          'Syntax: z <from bot> <to bot> <message>
          Case "z"
              OnBot = Param(Line, 2)
              ToNick = Param(Line, 3)
              If GetNextBot(OnBot) <> vsock Then
                RTU vsock, "error Fake message rejected ('" & OnBot & "' not connected via '" & Nick & "'!)"
                SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Fake message from " & Nick & " rejected ('" & OnBot & "' not connected via '" & Nick & "'!)"
              Else
                If LCase(ToNick) <> LCase(BotNetNick) Then
                  'Forward message to right bot
                  SendToBot ToNick, Line
                Else
                  'It's for me!
                  Message = Trim(Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3))))
                  HaltDefault = False
                  For ScNum = 1 To ScriptCount
                    If Scripts(ScNum).Hooks.BN_Msg Then
                      RunScriptX ScNum, "BN_Msg", OnBot, GetUserFlags(OnBot), Message
                    End If
                  Next ScNum
                  If HaltDefault = False Then
                    If MatchFlags(GetUserFlags(OnBot), "+b") Then
                      Select Case LCase(Param(Message, 1))
                        'GetOps reply
                        Case "gop_resp"
                            Rest = Trim(Right(Message, Len(Message) - Len(Param(Message, 1))))
                            If (InStr(Line, "don't recognize") <> 0) Then
                              SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop addhost " & MyHostmask
                              SpreadFlagMessage 0, "+t", "14[" & Time & "] *** GetOps: requesing AddHost from " & OnBot
                            ElseIf (InStr(Rest, "I don't monitor") = 0) And (InStr(Rest, "I'm not on") = 0) Then
                              SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps reply from " & OnBot & ": " & Trim(Right(Message, Len(Message) - Len(Param(Message, 1))))
                            End If
                        'GetOps message
                        Case "gop"
                            Select Case LCase(Param(Message, 2))
                              Case "invite"
                                  u2 = FindChan(Param(Message, 3))
                                  If u2 > 0 Then
                                    If Channels(u2).GotOPs Then
                                      If MatchFlags(GetUserChanFlags(OnBot, Param(Message, 3)), "+o") Then
                                        SendLine "invite " & Param(Message, 4) & " " & Channels(u2).Name, 1
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " requested invite for " & Channels(u2).Name & "."
                                      Else
                                        SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp Sorry, you aren't +o in my userlist for " & Channels(u2).Name & "."
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " failed invite for " & Channels(u2).Name & " (+o needed)."
                                      End If
                                    End If
                                  End If
                              Case "addhost"
                                  AddHost SocketItem(vsock).UserNum, OnBot, Mask(Param(Line, 6), 3)
                                  SpreadFlagMessage 0, "+t", "14[" & Time & "] *** GetOps: Adding " & Mask(Param(Line, 6), 3) & " to " & OnBot
                              Case "unban"
                                  u2 = FindChan(Param(Message, 3))
                                  If u2 > 0 Then
                                    If Channels(u2).GotOPs Then
                                      If MatchFlags(GetUserChanFlags(OnBot, Param(Message, 3)), "+o") Then
                                        For u = 1 To Channels(u2).BanCount
                                          If MatchWM(Channels(u2).BanList(u).Mask, Param(Message, 4)) Or MatchWM(Channels(u2).BanList(u).Mask, Param(Message, 5)) Then
                                            TimedEvent "UnBan " & Param(Message, 3) & " " & Channels(u2).BanList(u).Mask, Int(Rnd * 5)
                                          End If
                                        Next u
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " requested unban for " & Channels(u2).Name & "."
                                      Else
                                        SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp Sorry, you aren't +o in my userlist for " & Channels(u2).Name & "."
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " failed unban for " & Channels(u2).Name & " (+o needed)."
                                      End If
                                    End If
                                  End If
                              Case "key" '---LASTKEY
                                  u2 = FindChan(Param(Message, 3))
                                  If u2 > 0 Then
                                    If MatchFlags(GetUserChanFlags(OnBot, Param(Message, 3)), "+o") Then
                                      Message = GetChannelKey(u2)
                                      If Message <> "" Then
                                        SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop takekey " & Channels(u2).Name & " " & Message
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " requested key for " & Channels(u2).Name & "."
                                      End If
                                    Else
                                      SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp Sorry, you aren't +o in my userlist for " & Channels(u2).Name & "."
                                      SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " failed key for " & Channels(u2).Name & " (+o needed)."
                                    End If
                                  End If
                              Case "takekey"
                                  If InAutoJoinChannels(Param(Message, 3)) Then
                                    u2 = FindChan(Param(Message, 3))
                                    If u2 = 0 Then
                                      If Param(Message, 4) <> "" And Not IsOrdered("JOIN :" & Param(Message, 3)) Then
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " gave me the key for " & Param(Message, 3) & "."
                                        Order "JOIN :" & Param(Message, 3), 5
                                        SendLine "JOIN " & Param(Message, 3) & " " & Param(Message, 4), 1
                                      End If
                                    End If
                                  End If
                              Case "op"
                                  u2 = FindChan(Param(Message, 3))
                                  If u2 > 0 Then
                                    u = FindUser(Param(Message, 4), u2)
                                    If u > 0 Then
                                      If MatchFlags(GetUserChanFlags(OnBot, Param(Message, 3)), "+o") Then
                                        If LCase(SearchUserFromHostmask(Channels(u2).User(u).Hostmask)) = LCase(OnBot) Then
                                          If Channels(u2).GotOPs Then
                                            If InStr(Channels(u2).User(u).Status, "@") = 0 Then
                                              If GiveOp(Channels(u2).Name, Param(Message, 4)) Then
                                                SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " requested op for " & Channels(u2).Name & "."
                                              End If
                                            End If
                                          Else
                                            SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp I'm not op on " & Channels(u2).Name & "."
                                            SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " failed op (I'm not op on " & Channels(u2).Name & "!)."
                                          End If
                                        Else
                                          SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp I don't recognize you on IRC (your host: " & Mask(Channels(u2).User(u).Hostmask, 10) & ")"
                                          SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " failed op (host not recognized: " & Mask(Channels(u2).User(u).Hostmask, 10) & ")."
                                        End If
                                      Else
                                        SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp Sorry, you aren't +o in my userlist for " & Channels(u2).Name & "."
                                        SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " failed op for " & Channels(u2).Name & " (+o needed)."
                                      End If
                                    Else
                                      SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " gop_resp You are not on " & Channels(u2).Name & " for me."
                                    End If
                                  End If
                            End Select
                        'Extended GetOps message (resynching channels)
                        Case "xgop"
                            Bots(GetBotPos(OnBot)).SendRequests = True
                            Select Case LCase(Param(Message, 2))
                              Case "wantops?"
                                  If (Param(Message, 5) = "") Or (LCase(Param(Message, 5)) = LCase(MyNick)) Then
                                    u2 = FindChan(Param(Message, 3))
                                    If u2 > 0 Then
                                      If Not IsOrdered("gop " & Channels(u2).Name) Then
                                        SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " xgop op " & Channels(u2).Name & " " & MyNick
                                        Order "gop " & Channels(u2).Name, 15
                                        SpreadFlagMessage 0, "+t", "14[" & Time & "] *** GetOps: Requested op for " & Channels(u2).Name & " from " & OnBot & " (offered)."
                                      Else
                                        SendToBot OnBot, "z " & BotNetNick & " " & OnBot & " xgop no thanks"
                                      End If
                                    End If
                                  End If
                              Case "op"
                                  If MatchFlags(GetUserChanFlags(OnBot, Param(Message, 3)), "+o") Then
                                    u2 = FindChan(Param(Message, 3))
                                    If u2 > 0 Then
                                      u = FindUser(Param(Message, 4), u2)
                                      If u > 0 Then
                                        If LCase(Channels(u2).User(u).RegNick) = LCase(OnBot) Then
                                          If Channels(u2).GotOPs Then
                                            If GiveOp(Channels(u2).Name, Param(Message, 4)) Then
                                              SpreadFlagMessage vsock, "+m", "14[" & Time & "] *** GetOps: " & OnBot & " requested op for " & Channels(u2).Name & " (offered)."
                                            End If
                                          End If
                                        End If
                                      End If
                                    End If
                                  End If
                            End Select
                      End Select
                    End If
                  End If
                End If
              End If
          Case "zb" 'Broadcast
            For u = 1 To SocketCount
              If IsValidSocket(u) Then
                If MatchFlags(BotUsers(SocketItem(u).UserNum).Flags, "+b") And vsock <> u Then
                  RTU u, Line
                End If
              End If
            Next u
        End Select
    'Get bot nick
    Case SF_Status_BotGetName
      If LCase(Param(Line, 1)) = "error" Then
        SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetError, SocketItem(vsock).Hostmask, GetRest(Line, 2))
        RemoveSocket vsock, 0, "Error from", False
        Exit Sub
      End If
      SocketItem(vsock).RegNick = Line
      SocketItem(vsock).UserNum = GetUserNum(SocketItem(vsock).RegNick)
      If SocketItem(vsock).UserNum = 0 Then
        SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Unknown bot! Disconnecting...")
        RemoveSocket vsock, 0, "Unknown bot:", False
        Exit Sub
      End If
      SocketItem(vsock).Flags = GetUserFlags(SocketItem(vsock).RegNick)
      If MatchFlags(SocketItem(vsock).Flags, "-b") Then
        SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Not a bot! (+b missing) Disconnecting...")
        RTU vsock, "Sorry, this port is for botnet connects only. Please use the user telnet port (" & CStr(TelnetPort) & ")."
        RemoveSocket vsock, 0, "", True
        Exit Sub
      End If
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+r") = True Then
        SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Link from " & SocketItem(vsock).RegNick & " rejected (+r)"
        RTU vsock, "error You're not allowed to link (+r set)!"
        RemoveSocket vsock, 0, "", True
        Exit Sub
      End If
      If BotUsers(SocketItem(vsock).UserNum).Password <> "" Then
        'Reject bot if nick is already in botnet but not connected directly to me
        For u = 1 To BotCount
          If LCase(Bots(u).Nick) = LCase(SocketItem(vsock).RegNick) Then
            If LCase(Bots(u).SubBotOf) <> LCase(BotNetNick) Then
              SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Link from " & SocketItem(vsock).RegNick & " rejected: Bot is already in botnet"
              RTU vsock, "error You're already in this botnet!"
              DisconnectSocket vsock
              Exit Sub
            End If
          End If
        Next u
        'Request password
        SetSockFlag vsock, SF_Status, SF_Status_BotGetPass
        SetSockFlag vsock, SF_Silent, SF_YES
        RTU vsock, "passreq"
        SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Requesting password")
      Else
        'No password set - reject bot if nick is already in botnet
        If StrictDupeCheck(vsock) Then Exit Sub
        
        'winsock2_send handshake
        Message = ""
        For u = 1 To 11 + Int(Rnd * 5)
          If Int(Rnd * 2) = 1 Then
            Message = Message & Chr(Asc("a") + Int(Rnd * 26))
          Else
            Message = Message & Chr(Asc("0") + Int(Rnd * 10))
          End If
        Next u
        BotUsers(SocketItem(vsock).UserNum).Password = Message
        SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Sending handshake")
        SetSockFlag vsock, SF_Status, SF_Status_BotLinking
        SetSockFlag vsock, SF_LoggedIn, SF_YES
        SetSockFlag vsock, SF_Silent, SF_YES
        RTU vsock, "*hello!"
        RTU vsock, "version " & LongBotVersion & " " & CStr(ServerNickLen) & " AnGeL " & BotVersion & " <" & IIf(ServerNetwork <> "", ServerNetwork, "IRCNet") & ">"
        RTU vsock, "handshake " & Message
      End If
    'Get bot password
    Case SF_Status_BotGetPass
      If LCase(Param(Line, 1)) = "error" Then
        SpreadFlagMessage 0, "+t", MakeMsg(MSG_PLBotNetError, SocketItem(vsock).RegNick, GetRest(Line, 2))
        RemoveSocket vsock, 0, "Error from", False
        Exit Sub
      End If
      If Line = "-" Then
        SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Password needed! Disconnecting...")
        RemoveSocket vsock, 0, "Password needed:", True
        Exit Sub
      End If
      If Line <> BotUsers(SocketItem(vsock).UserNum).Password Then
        SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, MakeMsg(MSG_BNBadPass))
        RTU vsock, "badpass"
        RemoveSocket vsock, 0, "Bad password:", False
        Exit Sub
      End If
      
      'Check for duplicate links
      If DupeCheck(vsock) Then Exit Sub
      
      'Check for outgoing link attempts to the same bot
      LinkCheck vsock
      
      SetSockFlag vsock, SF_Status, SF_Status_BotLinking
      SetSockFlag vsock, SF_LoggedIn, SF_YES
      SetSockFlag vsock, SF_Silent, SF_YES
      
      RTU vsock, "*hello!"
      RTU vsock, "version " & LongBotVersion & " " & CStr(ServerNickLen) & " AnGeL " & BotVersion & " <" & IIf(ServerNetwork <> "", ServerNetwork, "IRCnet") & ">"
  End Select
End Sub

Public Sub TraceMessage(vsock As Long, Line As String)
  Dim u As Long, CurCheckBot As String, FoundOne As Boolean
  'Trace...
  't <ttl>:<who>@<source> <target> :<timestamp>:<bot1>:<bot2>...
 
  RTU vsock, "td " & Param(Line, 2) & " " & Param(Line, 4) & ":" & BotNetNick
  If Param(Line, 3) <> BotNetNick Then
    'need to forward to target
    CurCheckBot = Param(Line, 3)
    SendToBot CurCheckBot, "t " & ParamX(Param(Line, 2), ":", 1) - 1 & ":" & ParamX(Param(Line, 2), ":", 2) & " " & GetRest(Line, 3) & ":" & BotNetNick
  End If
End Sub

Public Sub TraceReply(vsock As Long, Line As String)
  Dim u As Long, CurCheckBot As String, FoundOne As Boolean, cNow As Currency, cDif As Currency, cThen As Currency
  'TraceReply
  'td <ttl>:<who>@<source> :<timestamp>:<bot1>:<bot2>...
  
  If ParamX(Param(Line, 2), "@", 2) = BotNetNick Then
    ' Reply für mich :)
    If ParamX(Param(Line, 2), ":", 1) = "9" Then
      CurCheckBot = ParamX(ParamX(Param(Line, 2), ":", 2), "@", 1)
      cNow = WinTickCount
      cThen = ParamX(Param(Line, 3), ":", 1)
      cDif = cNow - cThen
      For u = 1 To SocketCount
        If SocketItem(u).RegNick = CurCheckBot Then
          TU u, "14*** Trace Reply from " & ParamX(Param(Line, 3), ":", ParamXCount(Param(Line, 3), ":")) & ": " & cDif & "ms"
        End If
      Next u
    End If
  Else
    ' Reply für jemand andren :(
    ' need to forward to target
    CurCheckBot = ParamX(Param(Line, 2), "@", 2)
    SendToBot CurCheckBot, "td " & ParamX(Param(Line, 2), ":", 1) + 1 & ":" & ParamX(Param(Line, 2), ":", 2) & " " & GetRest(Line, 3)
  End If
End Sub


Public Sub BotNetSort(vsock As Long, InLine As String) ' : AddStack "BotNet_BotNetSort(" & vsock & ", " & SockNum & ", " & InLine & ")"
Dim chPos As Long, SPos As Long, Part As String, Line As String
  If vsock > SocketCount Then Exit Sub
  Line = SocketItem(vsock).InputBuffer + Replace(InLine, Chr(13), "")
  SPos = 1
  Do
    'chPos = InStr(SPos, Line, Chr(13))
    chPos = InStr(SPos, Line, Chr(10))
    If chPos = 0 Then Exit Do
    Part = Mid(Line, SPos, chPos - SPos)
    Output Part + vbCrLf
    Botnet vsock, Part
    On Local Error Resume Next
    If Not IsValidSocket(vsock) Then Exit Sub
    If Err.Number <> 0 Then
      Err.Clear
      Exit Sub
    End If
    On Error GoTo 0
    SPos = chPos + 1
    If Mid(Line, SPos, 1) = Chr(10) Then SPos = SPos + 1
  Loop
  Part = Mid(Line, SPos, Len(Line) - SPos + 1)
  SocketItem(vsock).InputBuffer = Part
  If Len(SocketItem(vsock).InputBuffer) > 5000 Then
    SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Botnet: Closed connection from " & SocketItem(vsock).Hostmask & ": Character flood"
    RemoveSocket vsock, 0, "Disconnected for flooding:", True
  End If
End Sub

Sub ReplyWHO(OnBot As String, Line As String) ' : AddStack "BotNetRoutines_ReplyWHO(" & OnBot & ", " & Line & ")"
Dim u As Long, u2 As Long, Rest As String, Message As String, FoundOne As Boolean
  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " WHO report for " & BotNetNick & ", AnGeL " & BotVersionEx + IIf(ServerNetwork <> "", "+" & ServerNetwork, "") & ":"
  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " ------------------------------------------------------------------"
  If Connected Then
    SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "   Online as: " & MyNick & " (" & Mask(MyHostmask, 10) & ")"
  Else
    SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "   Online as: <not online!>"
  End If
  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "   Uptime   : " & TimeSince(CStr(StartUpTime))
  Rest = BotChannels
  u2 = 0: Message = ""
  FoundOne = False
  For u = 1 To ParamCount(Rest)
    If (u2 = 5) Or ((Message <> "") And (Len(Message & " " & Param(Rest, u)) > 51)) Then
      If FoundOne = False Then
        SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "   Channels : " & Message
      Else
        SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "              " & Message
      End If
      FoundOne = True
      u2 = 0
      Message = ""
    End If
    u2 = u2 + 1: If Message <> "" Then Message = Message & " " & Param(Rest, u) Else Message = Param(Rest, u)
  Next u
  If u2 > 0 Then
    If FoundOne = False Then
      SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "   Channels : " & Message
    Else
      SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "              " & Message
    End If
  End If
  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2) & "   Admin    : " & GetPPString("Identification", "Admin", "Nobody <no@admin.de>", AnGeL_INI)
  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2)
  PListBotUsers OnBot, "p " & BotNetNick & " " & Param(Line, 2) & " "
  SendToBot OnBot, "p " & BotNetNick & " " & Param(Line, 2)
End Sub

'winsock2_send information about the botnet to a new bot
Sub BotNetLogin(vsock As Long) ' : AddStack "BotNetRoutines_BotNetLogin(" & vsock & ")"
Dim u As Long
  On Local Error Resume Next
  RTU vsock, "tb " & BotNetNick
  For u = 2 To BotCount
    RTU vsock, "n " & Bots(u).Nick & " " & Bots(u).SubBotOf & " " & Bots(u).SharingFlag + Bots(u).Version
  Next u
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If (GetSockFlag(u, SF_LocalVisibleUser) = SF_YES) Or (GetSockFlag(u, SF_Status) = SF_Status_BotNetParty) Then
        RTU vsock, "j !" & SocketItem(u).OnBot & " " & SocketItem(u).RegNick & " A " & GetLevelSign(SocketItem(u).Flags) + SocketItem(u).OrderSign & " " & Mask(SocketItem(u).Hostmask, 10)
        If SocketItem(u).AwayMessage <> "" Then
          RTU vsock, "i " & SocketItem(u).OnBot & " " & SocketItem(u).OrderSign & " " & LongToBase64(GetSecondsTillNow(SocketItem(u).LastEvent)) & " " & SocketItem(u).AwayMessage
        End If
      End If
    End If
  Next u
  RTU vsock, "el"
End Sub

Sub PreBotNetLogin(ByVal vsock As Long) ' : AddStack "BotNetRoutines_PreBotNetLogin(" & vsock & ")"
  SetSockFlag vsock, SF_Status, SF_Status_BotPreCache
  SocketItem(vsock).CurrentQuestion = ""
  BotNetLogin vsock
  TimedEvent "FinalBotNetLogin " & Trim(Str(vsock)) & " " & SocketItem(vsock).OrderSign, 4
End Sub

Sub FinalBotNetLogin(ByVal vsock As Long, ByVal OrderSign As String) ' : AddStack "BotNetRoutines_FinalBotNetLogin(" & vsock & ", " & OrderSign & ")"
Dim u As Long, SockNum As Long, BotVersion As String, BotBuffer As String
  If Not IsValidSocket(vsock) Then Exit Sub
  If GetSockFlag(vsock, SF_Status) <> SF_Status_BotPreCache Then Exit Sub
  If Trim(SocketItem(vsock).OrderSign) <> Trim(OrderSign) Then Exit Sub
  
  SetSockFlag vsock, SF_Status, SF_Status_Bot
  BotVersion = SocketItem(vsock).SetupChan
  BotBuffer = SocketItem(vsock).CurrentQuestion
  SocketItem(vsock).CurrentQuestion = ""
  AddBot SocketItem(vsock).RegNick, BotNetNick, "-", BotVersion, vsock
  ToBotNet vsock, "n " & SocketItem(vsock).RegNick & " " & BotNetNick & " !-" & BotVersion
  SpreadMessageEx 0, -1, SF_Local_Bot, MakeMsg(MSG_PLBotNetConnect, SocketItem(vsock).RegNick)
  For u = 1 To ParamXCount(BotBuffer, Chr(0))
    If Not IsValidSocket(vsock) Then Exit For
    Botnet vsock, ParamX(BotBuffer, Chr(0), u)
  Next u
  'Start userfile sharing if enabled
  CheckSharing vsock
End Sub

Public Function AddBotNetUser(Line As String) As Boolean
  Dim u As Long, OrderSign As String, TheBot As String, FoundOne As Boolean, NewSock As Long
  Dim ErrPos As Long
'A10'
  On Error GoTo AUError
ErrPos = 1
  TheBot = Param(Line, 2)
ErrPos = 2
  OrderSign = Right(Param(Line, 5), Len(Param(Line, 5)) - 1)
ErrPos = 3
  If Left(TheBot, 1) = "!" Then TheBot = Right(TheBot, Len(TheBot) - 1)
ErrPos = 4
  If GetBotPos(TheBot) = 0 Then SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Fake user from non-existing " & TheBot & ": " & Param(Line, 3): Exit Function
  FoundOne = False
  For u = 1 To SocketCount
ErrPos = 5
    If IsValidSocket(u) Then
ErrPos = 6
      If LCase(SocketItem(u).OnBot) = LCase(TheBot) And SocketItem(u).OrderSign = OrderSign Then
        Select Case Left(Param(Line, 5), 1)
          Case "*": SocketItem(u).Flags = "n"
          Case "+": SocketItem(u).Flags = "m"
          Case "%": SocketItem(u).Flags = "t"
          Case "@": SocketItem(u).Flags = "o"
          Case "-": SocketItem(u).Flags = ""
        End Select
        If LCase(Param(Line, 1)) = "join" Then
ErrPos = 7
          SocketItem(u).PLChannel = CLng(Param(Line, 4))
        Else
ErrPos = 8
          SocketItem(u).PLChannel = Base64ToLong(Param(Line, 4))
        End If
        FoundOne = True
      End If
    End If
  Next u
  If FoundOne Then AddBotNetUser = False: Exit Function
ErrPos = 9
  NewSock = AddSocket
ErrPos = 10
  SocketItem(NewSock).RegNick = Param(Line, 3)
  SocketItem(NewSock).Hostmask = Param(Line, 6)
  SocketItem(NewSock).OnBot = TheBot
ErrPos = 11
  If GetRealNick(SocketItem(NewSock).RegNick) <> "" Then
    WriteSeenEntry SocketItem(NewSock).RegNick, "", Now, SocketItem(NewSock).OnBot, "*partyline*", Mask(SocketItem(NewSock).Hostmask, 10)
  Else
    WriteExtSeenEntry SocketItem(NewSock).RegNick, "", Now, SocketItem(NewSock).OnBot, "*partyline*", Mask(SocketItem(NewSock).Hostmask, 10)
  End If
ErrPos = 12
  SetSockFlag NewSock, SF_Status, SF_Status_BotNetParty
  SetSockFlag NewSock, SF_Echo, SF_NO
  SetSockFlag NewSock, SF_Colors, SF_YES
  SocketItem(NewSock).OrderSign = OrderSign
  SocketItem(NewSock).SocketNumber = 0
  SocketItem(NewSock).UserNum = 0
ErrPos = 13
  Select Case Left(Param(Line, 5), 1)
    Case "*": SocketItem(NewSock).Flags = "n"
    Case "+": SocketItem(NewSock).Flags = "m"
    Case "%": SocketItem(NewSock).Flags = "t"
    Case "@": SocketItem(NewSock).Flags = "o"
    Case "-": SocketItem(NewSock).Flags = ""
  End Select
  AddBotNetUser = True
Exit Function
AUError:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  PutLog "||| ] AddUser ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
  PutLog "||| ] Der Fehler " & ErrNumber & " (" & ErrDescription & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & ErrPos
  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
'  SendNote ("AnGeL AddBotNetUser ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten (" & cStr(ErrPos)) & "): " & Line
'  SendNote AnGeL AddBotNetUser ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") ist beim Bearbeiten folgender Zeile aufgetreten (" & cStr(ErrPos)) & "): " & Line
End Function

Function GetPartNick(Complete As String) As String
  If InStr(Complete, "@") > 0 Then
    GetPartNick = Left(Complete, InStr(Complete, "@") - 1)
  Else
    GetPartNick = ""
  End If
End Function

Function GetPartBot(Complete As String) As String
  If InStr(Complete, "@") > 0 Then
    GetPartBot = Right(Complete, Len(Complete) - InStr(Complete, "@"))
  Else
    GetPartBot = Complete
  End If
End Function

