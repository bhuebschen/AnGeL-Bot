Attribute VB_Name = "Partyline_Commands"
',-======================- ==-- -  -
'|   AnGeL - Partyline - Commands
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Private Type TMatch
  UserNum As Long
  Hostmask As String
End Type


Sub AddInvite(vsock As Long, Line As String)
  Dim u As Long, u2 As Long, ToInvite As String, Rest As String, Nick As String
  Dim Flags As String, Messg As String
  Nick = SocketItem(vsock).RegNick
  Flags = SocketItem(vsock).Flags
  ToInvite = Param(Line, 2)
  If ToInvite = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+invite <hostmask> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If IsValidHostmask(ToInvite) = False Then TU vsock, "5*** A valid hostmask must look like this: nick!identd@host.domain": Exit Sub
  If Len(Mask(ToInvite, 12)) > 10 Then TU vsock, "5*** Sorry, the ident is too long (longer than 10 characters).": Exit Sub
  If Len(Mask(ToInvite, 13)) > ServerNickLen Then TU vsock, "5*** Sorry, the nick is too long (longer than " & CStr(ServerNickLen) & " characters).": Exit Sub
  If IsValidChannel(Left(Param(Line, 3), 1)) Then Rest = Param(Line, 3) Else Rest = "*"
  If (Rest = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to place global invitions. Use: .+invite <hostmask> <#YOUR CHANNEL> (comment)": Exit Sub
  If MatchFlags(GetUserChanFlags(Nick, Rest), "-m") Then TU vsock, "5*** Sorry, you're not allowed to place invitations in this channel.": Exit Sub
  For u = 1 To InviteCount
    If (LCase(Invites(u).Hostmask) = LCase(ToInvite)) And (LCase(Invites(u).Channel) = LCase(Rest)) Then TU vsock, "5*** This hostmask is already invited.": Exit Sub
  Next u
  WritePPString Param(Line, 2), "Channel", Rest, HomeDir & "Invites.ini"
  WritePPString Param(Line, 2), "CreatedAt", Now, HomeDir & "Invites.ini"
  WritePPString Param(Line, 2), "CreatedBy", Nick, HomeDir & "Invites.ini"
  InviteCount = InviteCount + 1: If InviteCount > UBound(Invites()) Then ReDim Preserve Invites(UBound(Invites()) + 5)
  Invites(InviteCount).Hostmask = Param(Line, 2)
  Invites(InviteCount).Channel = Rest
  Invites(InviteCount).CreatedAt = Now
  Invites(InviteCount).CreatedBy = Nick
  If Rest = "*" Then
    For u = 1 To ChanCount
      CheckInvites Channels(u).Name
    Next u
  Else
    CheckInvites Rest
  End If
  TU vsock, "3*** Added " & IIf(Rest = "*", "global", Rest & " channel") & " invite '10" & Param(Line, 2) & "3' as number " & CStr(InviteCount) & "" & "."
End Sub
Sub AddExcept(vsock As Long, Line As String) ' : AddStack "GUIServer_AddExcept(" & vsock & ", " & Line & ")"
  Dim u As Long, u2 As Long, ToExcept As String, Rest As String, Nick As String
  Dim Flags As String, Messg As String
  Nick = SocketItem(vsock).RegNick
  Flags = SocketItem(vsock).Flags
  ToExcept = Param(Line, 2)
  If ToExcept = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+except <hostmask> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If IsValidHostmask(ToExcept) = False Then TU vsock, "5*** A valid hostmask must look like this: nick!identd@host.domain": Exit Sub
  If Len(Mask(ToExcept, 12)) > 10 Then TU vsock, "5*** Sorry, the ident is too long (longer than 10 characters).": Exit Sub
  If Len(Mask(ToExcept, 13)) > ServerNickLen Then TU vsock, "5*** Sorry, the nick is too long (longer than " & CStr(ServerNickLen) & " characters).": Exit Sub
  If IsValidChannel(Left(Param(Line, 3), 1)) Then Rest = Param(Line, 3) Else Rest = "*"
  If (Rest = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to place global exceptions. Use: .+except <hostmask> <#YOUR CHANNEL> (comment)": Exit Sub
  If MatchFlags(GetUserChanFlags(Nick, Rest), "-m") Then TU vsock, "5*** Sorry, you're not allowed to place exceptions in this channel.": Exit Sub
  For u = 1 To ExceptCount
    If (LCase(Excepts(u).Hostmask) = LCase(ToExcept)) And (LCase(Excepts(u).Channel) = LCase(Rest)) Then TU vsock, "5*** This hostmask is already excepted.": Exit Sub
  Next u
  WritePPString Param(Line, 2), "Channel", Rest, HomeDir & "Excepts.ini"
  WritePPString Param(Line, 2), "CreatedAt", Now, HomeDir & "Excepts.ini"
  WritePPString Param(Line, 2), "CreatedBy", Nick, HomeDir & "Excepts.ini"
  ExceptCount = ExceptCount + 1: If ExceptCount > UBound(Excepts()) Then ReDim Preserve Excepts(UBound(Excepts()) + 5)
  Excepts(ExceptCount).Hostmask = Param(Line, 2)
  Excepts(ExceptCount).Channel = Rest
  Excepts(ExceptCount).CreatedAt = Now
  Excepts(ExceptCount).CreatedBy = Nick
  If Rest = "*" Then
    For u = 1 To ChanCount
      CheckExcepts Channels(u).Name
    Next u
  Else
    CheckExcepts Rest
  End If
  TU vsock, "3*** Added " & IIf(Rest = "*", "global", Rest & " channel") & " except '10" & Param(Line, 2) & "3' as number " & CStr(ExceptCount) & "" & "."
End Sub

Public Sub AddServer(vsock As Long, Line As String) ' : AddStack "GUIServer_AddServer(" & vsock & ", " & Line & ")"
Dim u As Long, Rest As String
  If Param(Line, 2) = "" Or InStr(Param(Line, 2), ":") = 0 Then TU vsock, MakeMsg(ERR_CommandUsage, ".+server <address:port> (|proxy(:port))"): Exit Sub
  u = 1
  Do
    If u = 1 Then Rest = "Server" Else Rest = "Server" & CStr(u)
    Rest = GetPPString("Server", Rest, "", AnGeL_INI)
    If LCase(GetRest(Line, 2)) = LCase(Rest) Then TU vsock, "5*** Sorry, this server already exists as number " & CStr(u) & ".": Exit Sub
    If Rest = "" Then Exit Do
    u = u + 1
  Loop
  If u = 1 Then Rest = "Server" Else Rest = "Server" & CStr(u)
  WritePPString "Server", Rest, GetRest(Line, 2), AnGeL_INI
  TU vsock, "3*** Added '10" & Param(Line, 2) & "3' as server number " & CStr(u) & "."
  If HubBot = True Then TU vsock, "3*** Hub mode left. Type '.jump 1' to connect to this server.": HubBot = False: SetTrayIcon SI_Offline
End Sub

Public Sub RemServer(vsock As Long, Line As String) ' : AddStack "GUIServer_RemServer(" & vsock & ", " & Line & ")"
Dim u As Long, Rest As String, ServerCount As Long, Servers() As String, RemOne As String, RemNum As Long
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-server <address:port or number>"): Exit Sub
  u = 1
  ReDim Preserve Servers(5)
  Do
    If u = 1 Then Rest = "Server" Else Rest = "Server" & CStr(u)
    Rest = GetPPString("Server", Rest, "", AnGeL_INI)
    If Rest = "" Then Exit Do
    If Not (LCase(GetRest(Line, 2)) = LCase(Rest) Or u = Val(Param(Line, 2))) Then
      ServerCount = ServerCount + 1
      If ServerCount > UBound(Servers()) Then ReDim Preserve Servers(UBound(Servers()) + 5)
      Servers(ServerCount) = Rest
    Else
      RemOne = Rest
      RemNum = u
    End If
    u = u + 1
  Loop
  If RemNum = 0 Then TU vsock, "5*** Sorry, I didn't find this server. Type '.servers' to get a list.": Exit Sub
  For u = 1 To ServerCount
    If u = 1 Then Rest = "Server" Else Rest = "Server" & CStr(u)
     WritePPString "Server", Rest, Servers(u), AnGeL_INI
  Next u
  u = ServerCount + 1
  If u = 1 Then Rest = "Server" Else Rest = "Server" & CStr(u)
  WritePPString "Server", Rest, "", AnGeL_INI
  DeletePPString "Server", Rest, AnGeL_INI
  TU vsock, "3*** Removed server '10" & Param(RemOne, 1) & "3' (number " & CStr(RemNum) & ")."
  If ServerCount = 0 Then TU vsock, "3*** This was my last server! Hub mode activated.": HubBot = True: SetTrayIcon SI_Hub
End Sub



Public Sub GUIChattr(vsock As Long, Line As String) ' : AddStack "GUI_GUIChattr(" & vsock & ", " & Line & ")"
Dim u As Long, Nick As String, ChangeLine As String, ChgPos As String, Rest As String
Dim ChattrNick As String, NewFlags As String, UsNum As Long
Dim HostNum As Long, Invalid As String, Flags As String, Chan As String
  Nick = SocketItem(vsock).RegNick
  Flags = SocketItem(vsock).Flags
  If Left(Param(Line, 3), 1) <> "+" And Left(Param(Line, 3), 1) <> "-" Then TU vsock, MakeMsg(ERR_CommandUsage, ".chattr <nick> <+/-flags> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If Param(Line, 4) = "" Then
    ChgPos = "Flags"
  Else
    If Not (IsValidChannel(Left(Param(Line, 4), 1))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".chattr <nick> <+/-flags> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
    Chan = Param(Line, 4)
    ChgPos = "ChFlags" & Chan
  End If
  If MatchFlags(Flags, "-n") Then
    If Chan = "" Then TU vsock, "5*** Sorry, you're not allowed to change global flags.": Exit Sub
    If Chan <> "" Then If MatchFlags(GetUserChanFlags(Nick, Chan), "-n") Then TU vsock, "5*** Sorry, you're not allowed to change user flags for this channel.": Exit Sub
  End If
  ChangeLine = Param(Line, 3)
  
  'Don't allow .chattr <nick> +b / -b
  If InStr(LCase(ChangeLine), "b") > 0 Then
    If InStr(LCase(GetPosFlags(Param(Line, 3))), "b") > 0 Then
      TU vsock, "5*** You can't give somebody the +b (bot) flag! Add bots"
      TU vsock, "5    with '.+bot' and they automatically get this flag."
    Else
      TU vsock, "5*** You can't remove the bot flag! Add users with the"
      TU vsock, "5    commands '.adduser' or '.+user' instead."
    End If
    Exit Sub
  End If
  
  ChattrNick = GetRealNick(Param(Line, 2))
  If Not UserExist(ChattrNick) Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
  If MatchFlags(Flags, "-s") And MatchFlags(GetUserFlags(ChattrNick), "+s") Then TU vsock, "5*** " & ChattrNick & " is a super owner. You can't change the flags of this user.": Exit Sub
  If MatchFlags(ChangeLine, "+h") Then TU vsock, "5*** Sorry, chattr is no longer able to change bot flags.": TU vsock, "5    Please use the new command '.botattr'.": Exit Sub
  
  'Check if flags can be applied for bots
  Rest = GetPosFlags(ChangeLine)
  If MatchFlags(GetUserFlags(ChattrNick), "+b") Then
    If MatchFlags(Rest, "+i") Or MatchFlags(Rest, "+j") Or MatchFlags(Rest, "+m") Or MatchFlags(Rest, "+n") Or MatchFlags(Rest, "+p") Or MatchFlags(Rest, "+s") Or MatchFlags(Rest, "+t") Or MatchFlags(Rest, "+w") Then
      TU vsock, "5*** Sorry, the following flags can't be used for bots: i,j,m,n,p,s,t,w": Exit Sub
    End If
  End If
  
  Invalid = CheckForInvalidFlags(ChangeLine)
  If Chan = "" Then
    If MatchFlags(Rest, "+s") And MatchFlags(Flags, "-s") Then
      TU vsock, "5*** You can't give anybody +s without having +s yourself!": Exit Sub
    End If
  Else
    ChangeLine = ChangeLine & " " & Chan
  End If
  
  'Do the chattr ;)
  Select Case Chattr(ChattrNick, ChangeLine)
    Case CH_Success
      If Invalid <> "" Then TU vsock, "14*** Ignored invalid flag" & IIf(Len(Invalid) > 1, "s", "") & ": " & Invalid & ""
      If Chan = "" Then
        TU vsock, "3*** Flags for " & ChattrNick & " are now """ & ExtReply & """"
      Else
        TU vsock, "3*** Flags for " & ChattrNick & " in " & Chan & " are now """ & ExtReply & """"
      End If
      'winsock2_send sharing message. (without +s)
      If Replace(Param(ChangeLine, 1), "s", "") <> "+" Then SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .chattr " & ChattrNick & " " & Replace(Param(ChangeLine, 1), "s", "") & IIf(GetRest(ChangeLine, 2) <> "", " " & GetRest(ChangeLine, 2), "")
    Case CH_NoChanges
      If Invalid <> "" Then
        TU vsock, "5*** The flags you specified were already set or invalid."
        TU vsock, "5    Type '.help whois' to get a list of all valid flags."
      Else
        TU vsock, "5*** The flags you specified were already set -> no changes."
      End If
      If Chan = "" Then
        TU vsock, "3*** Flags for " & ChattrNick & " are still """ & ExtReply & """"
      Else
        TU vsock, "3*** Flags for " & ChattrNick & " in " & Chan & " are still """ & ExtReply & """"
      End If
    Case CH_NoChanFlag
      TU vsock, "5*** Sorry, the following flags can only be used as global flags: b,i,j,p,s,t,w"
  End Select
End Sub

Public Sub GUIBotattr(vsock As Long, Line As String) ' : AddStack "GUI_GUIBotattr(" & vsock & ", " & Line & ")"
Dim u As Long, Nick As String, ChangeLine As String, ChgPos As String, Rest As String
Dim FlagsBefore As String, NewFlags As String, UsNum As Long
  If (Left(Param(Line, 3), 1) <> "+") And (Left(Param(Line, 3), 1) <> "-") Then TU vsock, MakeMsg(ERR_CommandUsage, ".botattr <bot> <+/-flags>"): Exit Sub
  ChgPos = "BotFlags"
  ChangeLine = Param(Line, 3)
  UsNum = GetUserNum(Param(Line, 2))
  If UsNum = 0 Then TU vsock, MakeMsg(ERR_BotNotFound, Param(Line, 2)): Exit Sub
  FlagsBefore = BotUsers(UsNum).BotFlags
  Rest = GetPosFlags(ChangeLine)
  '.botattr can only be used for bots
  If MatchFlags(BotUsers(UsNum).Flags, "-b") Then TU vsock, "5*** Sorry, this command can only be used for bots.": Exit Sub
  If Param(Line, 4) = "" Then
    If InStr(Rest, "h") > 0 Then ChangeLine = ChangeLine & "-a"
    If InStr(Rest, "a") > 0 Then ChangeLine = ChangeLine & "-h"
  End If
  NewFlags = GetBotattrResult(FlagsBefore, ChangeLine)
  If FlagsBefore <> NewFlags Then
    BotUsers(UsNum).BotFlags = NewFlags
    'winsock2_send sharing message.
    If Replace(Replace(ChangeLine, "h", ""), "a", "") <> "+" Then SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .botattr " & Param(Line, 2) & " " & Replace(Replace(ChangeLine, "a", ""), "h", "")
  Else
    TU vsock, "5*** The flags you specified couldn't be set. Perhaps the flags"
    TU vsock, "5    are not valid or they were already set. Use '.whois <bot>'"
    TU vsock, "5    to get this bot's current flags or '.help whois' to get a"
    TU vsock, "5    list of all valid bot flags."
    If Param(Line, 4) = "" Then TU vsock, "3*** Bot flags for " & BotUsers(UsNum).Name & " are still """ & NewFlags & """"
    Exit Sub
  End If
  TU vsock, "3*** Bot flags for " & BotUsers(UsNum).Name & " are now """ & NewFlags & """"
End Sub

Public Sub GUIAddHost(vsock As Long, Line As String) ' : AddStack "GUI_GUIAddHost(" & vsock & ", " & Line & ")"
  If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+host <nick> <hostmask>"): Exit Sub
  Select Case AddHost(SocketItem(vsock).UserNum, Param(Line, 2), Param(Line, 3))
    Case AH_UserNotFound
      TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2))
    Case AH_InvalidHost
      TU vsock, "5*** A valid hostmask must look like this: nick!identd@host.domain"
    Case AH_AlreadyThere
      TU vsock, "5*** This hostmask already belongs to " & ExtReply & "."
    Case AH_MatchingUser
      FailCommand vsock, "+m", Line
      TU vsock, "5*** Can't add this hostmask - it belongs to " & ExtReply & "."
    Case AH_TooManyHosts
      TU vsock, "5*** Sorry, maximum number of hostmasks reached (20)."
    Case AH_Success
      SucceedCommand vsock, "+m", Line
      TU vsock, "3*** Added this hostmask to " & ExtReply & "."
      UpdateRegUsers "A " & Param(Line, 3)
      SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .+host " & Line
  End Select
End Sub

Public Sub GUIRemHost(vsock As Long, Line As String) ' : AddStack "GUI_GUIRemHost(" & vsock & ", " & Line & ")"
Dim Flags As String, Nick As String, RealNick As String, HostToRemove As String
Dim UsNum As Long, u As Long, FoundOne As Boolean, RemoveThis As String
  Flags = SocketItem(vsock).Flags
  Nick = SocketItem(vsock).RegNick
  
  If IsValidHostmask(Param(Line, 2)) Then
    RealNick = Nick
    HostToRemove = Param(Line, 2)
  Else
    RealNick = GetRealNick(Param(Line, 2))
    HostToRemove = Param(Line, 3)
  End If
  If HostToRemove = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-host (nick) <hostmask>"): Exit Sub
  If RealNick = "" Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
  If MatchFlags(Flags, "-m") Then
    If RealNick <> Nick Then
      If MatchFlags(Flags, "-t") Then
        TU vsock, "5*** Error: You're not allowed to remove other users' hostmasks.": FailCommand vsock, "+m", Line: Exit Sub
      Else
        If MatchFlags(GetUserFlags(RealNick), "-b") Then
          TU vsock, "5*** Error: You're only allowed to remove bot hostmasks.": FailCommand vsock, "+m", Line: Exit Sub
        End If
      End If
    End If
  End If
  If MatchFlags(Flags, "-s") And MatchFlags(GetUserFlags(RealNick), "+s") Then
    TU vsock, "5*** Error: " & RealNick & " is a super owner. You can't remove hostmasks of this user.": FailCommand vsock, "+m", Line: Exit Sub
  End If
    
  Select Case RemHost(RealNick, HostToRemove)
    Case RH_HostNotFound
      UsNum = GetUserNum(RealNick)
      For u = BotUsers(UsNum).HostMaskCount To 1 Step -1
        If MatchWM(HostToRemove, BotUsers(UsNum).HostMasks(u)) Then
          FoundOne = True
          RemoveThis = BotUsers(UsNum).HostMasks(u)
          Select Case RemHost(RealNick, RemoveThis)
            Case RH_Success
              TU vsock, "3*** Removed matching hostmask: " & RemoveThis
            Case Else
              TU vsock, "5*** Error: Couldn't remove matching hostmask '" & RemoveThis & "'!"
          End Select
        End If
      Next u
      If FoundOne Then
        TU vsock, "3*** Done."
      Else
        If RealNick = Nick Then
          TU vsock, "5*** Error: This is none of your hostmasks."
        Else
          TU vsock, "5*** Error: This is no hostmask of " & RealNick & "."
        End If
      End If
    Case RH_Success
      If RealNick = Nick Then
        TU vsock, "3*** Removed this hostmask from your user record."
      Else
        TU vsock, "3*** Removed this hostmask from " & RealNick & "."
      End If
      SucceedCommand vsock, "+m", Line
      UpdateRegUsers "R " & RealNick
      SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .-host " & Line
  End Select
End Sub

Public Sub ListSplits(vsock As Long) ' : AddStack "GUI_ListSplits(" & vsock & ")"
Dim u As Long
  If SplitServerCount = 0 Then TU vsock, "5*** There are currently no registered splits.": Exit Sub
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "2*** Listing splits:"
    TU vsock, "0,1 Splitted server:                     | Since:     "
  Else
    TU vsock, "*** Listing splits:"
    TU vsock, " Splitted server:                       Since:"
    TU vsock, " -------------------------------------  -----------"
  End If
  For u = 1 To SplitServerCount
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, " " & Spaces2(36, SplitServers(u).Name) & "14 | " & TimeSpan2(SplitServers(u).SplittedAt)
    Else
      TU vsock, " " & Spaces2(36, SplitServers(u).Name) & "   " & TimeSpan2(SplitServers(u).SplittedAt)
    End If
  Next u
  TU vsock, EmptyLine
End Sub

Public Sub ShowNews(vsock As Long) ' : AddStack "GUI_ShowNews(" & vsock & ")"
  TU vsock, "2*** News in AnGeL " & BotVersionEx + IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", "")
  TU vsock, " 7-4=5>  1.  Introduced new userfile format (combines hostmask.txt and extusers.ini into userlist.txt)"
  TU vsock, " 7-4=5>  2.  Made '+host' and '-host' much faster"
  TU vsock, " 7-4=5>  3.  Fixed GetObject security hole"
  TU vsock, " 7-4=5>  4.  Made AllowCreateObject read-only for scripts"
  TU vsock, " 7-4=5>  5.  Blocked script write access to AnGeL.ini"
  TU vsock, " 7-4=5>  6.  Added 'Srv_connect'/'fa_command'/'fa_userjoin'/'fa_userleft'/'fa_downloadbegin'/'fa_downloadcomplete'/'fa_uploadbegin'/'fa_uploadcomplete'/'seen'/'KI' hook"
  TU vsock, " 7-4=5>  7.  Added local port check in IdentD"
  TU vsock, " 7-4=5>  8.  Fixed bug when using <ReplyToSub> in scripts"
  TU vsock, " 7-4=5>  9.  Fixed psyBNC-bug (flood protection stalled)"
  TU vsock, " 7-4=5>  10. Fixed VB5<->VB6 Prev. Instance bug"
  TU vsock, " 7-4=5>  11. Compressed AnGeL.exe (it's much smaller now!)"
  TU vsock, " 7-4=5>  12. Telnet Server 1.2 (colors)"
  TU vsock, " 7-4=5>  13. Added Setup for better Network configuration (.netsetup)"
  TU vsock, " 7-4=5>  14. Made DCC-PortRange customizable (.botsetup)"
  TU vsock, " 7-4=5>  15. Made Command-Prefix customizable (.botsetup)"
  TU vsock, " 7-4=5>  16. Added KI-Setup (.kisetup)"
  TU vsock, " 7-4=5>  17. some new Script-Commands (listed @ http://angel.web-pp.de)"
  TU vsock, " 7-4=5>  18. Added support for usermodes +h,+a and +q (HalfOp, Admin, Owner)"
  TU vsock, " 7-4=5>  19. Reworked complete NT-Service Handler."
  TU vsock, " 7-4=5>  20. Changed nearly all arrays to dynamics. It is now possible to join up to ~65000 Channels"
  TU vsock, " 7-4=5>  21. Added FirstTimeKeyword routines. Should no longer kill the Word if an error occurs."
  TU vsock, " 7-4=5>  22. Lots of new Partyline commands including .cycle / .rejoin to 'hop' a channel and .reload to reload a script."
  TU vsock, " 7-4=5>  23. many many more not to be mentioned here."
  TU vsock, EmptyLine
End Sub

Public Sub GUIAddScript(vsock As Long, Line As String) ' : AddStack "GUI_GUIAddScript(" & vsock & ", " & Line & ")"
Dim NewScript As String, FileName As String, Description As String
Dim Rest As String, ScriptStarted As Boolean, u As Long
  NewScript = GetFileName(Param(Line, 2))
  If NewScript = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+script <filename>"): Exit Sub
  If Right(LCase(NewScript), 4) <> ".asc" And Right(LCase(NewScript), 5) <> ".perl" Then NewScript = NewScript & ".asc"
  For u = 1 To ScriptCount
    FileName = Scripts(u).Name
    If LCase(FileName) = LCase(NewScript) Then
      Description = Scripts(u).Description
      TU vsock, "5*** Sorry, this script is already running: " & Description
      Exit Sub
    End If
  Next u
  Rest = GetPPString("Scripts", "Load", "", AnGeL_INI)
  For u = 1 To ParamCount(Rest)
    FileName = Param(Rest, u)
    If LCase(FileName) = LCase(NewScript) Then
      TU vsock, "3*** The script should be running, but it isn't. Restarting..."
      LoadScript FileAreaHome & "Scripts\" & FileName
      Exit Sub
    End If
  Next u
  ScriptStarted = LoadScript(FileAreaHome & "Scripts\" & NewScript)
  If Not ScriptStarted Then TU vsock, "5*** Sorry, the script couldn't be added.": Exit Sub
  Rest = Rest & " " & NewScript
  WritePPString "Scripts", "Load", Rest, AnGeL_INI
  TU vsock, "3*** Loaded script: " & NewScript
End Sub

Public Sub GUIRemScript(vsock As Long, Line As String) ' : AddStack "GUI_GUIRemScript(" & vsock & ", " & Line & ")"
Dim OldScript As String, FileName As String, Description As String
Dim Rest As String, NewLine As String, FoundScript As Boolean, u As Long
  If ScriptCMDs.CurrentScript <> 0 Then SpreadFlagMessage 0, "+n", "4*** Scripting error: Scripts can't be removed via ExecuteCommand!": Exit Sub
  OldScript = GetFileName(Param(Line, 2))
  If OldScript = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-script <filename>"): Exit Sub
  If Right(LCase(OldScript), 4) <> ".asc" And Right(LCase(OldScript), 5) <> ".perl" Then OldScript = OldScript & ".asc"
  For u = 1 To ScriptCount
    FileName = Scripts(u).Name
    If LCase(FileName) = LCase(OldScript) Then
      Description = Scripts(u).Description
      SpreadFlagMessage 0, "+n", "14*** Removing script: " & Description
      RemScript u
      FoundScript = True
      Exit For
    End If
  Next u
  Rest = GetPPString("Scripts", "Load", "", AnGeL_INI)
  For u = 1 To ParamCount(Rest)
    FileName = Param(Rest, u)
    If LCase(FileName) = LCase(OldScript) Then
      SpreadFlagMessage 0, "+n", "14*** Removing link to " & FileName & "..."
      FoundScript = True
    Else
      If NewLine <> "" Then NewLine = NewLine & " " & FileName Else NewLine = FileName
    End If
  Next u
  If FoundScript Then
    WritePPString "Scripts", "Load", NewLine, AnGeL_INI
    TU vsock, "3*** Done."
  Else
    TU vsock, "5*** Sorry, I couldn't find this script."
  End If
End Sub

Public Function ShowMOTD(vsock As Long) As Boolean ' : AddStack "GUI_ShowMOTD(" & vsock & ")"
Dim FileNum As Integer, MLine As String, FoundSth As Boolean, ShowToThisUser As Boolean
  On Local Error Resume Next
  If Dir(HomeDir & "Motd.txt") <> "" Then
    FileNum = FreeFile: Open HomeDir & "Motd.txt" For Input As #FileNum
    ShowToThisUser = True
    Do While Not EOF(FileNum)
      Line Input #FileNum, MLine
      If InStr(MLine, "%{") > 0 And (InStr(MLine, "}") > InStr(MLine, "%{")) Then
        MLine = Mid(MLine, InStr(MLine, "%{") + 2, InStr(MLine, "}") - InStr(MLine, "%{") - 1)
        If LCase(MLine) = "+color" Then
          ShowToThisUser = (GetSockFlag(vsock, SF_Colors) = SF_YES)
        ElseIf LCase(MLine) = "-color" Then
          ShowToThisUser = (GetSockFlag(vsock, SF_Colors) = SF_NO)
        Else
          ShowToThisUser = MatchFlags(SocketItem(vsock).Flags, MLine)
        End If
      Else
        Do
          FoundSth = False
          If InStr(MLine, "%N") > 0 Then MLine = Left(MLine, InStr(MLine, "%N") - 1) & SocketItem(vsock).RegNick & Right(MLine, Len(MLine) - InStr(MLine, "%N") - 1): FoundSth = True
          If InStr(MLine, "%B") > 0 Then MLine = Left(MLine, InStr(MLine, "%B") - 1) & BotNetNick & Right(MLine, Len(MLine) - InStr(MLine, "%B") - 1): FoundSth = True
          If InStr(MLine, "%V") > 0 Then MLine = Left(MLine, InStr(MLine, "%V") - 1) & "AnGeL " & BotVersion & IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", "") & Right(MLine, Len(MLine) - InStr(MLine, "%V") - 1): FoundSth = True
          If InStr(MLine, "%C") > 0 Then MLine = Left(MLine, InStr(MLine, "%C") - 1) & BotChannels + Right(MLine, Len(MLine) - InStr(MLine, "%C") - 1): FoundSth = True
          If InStr(MLine, "%T") > 0 Then MLine = Left(MLine, InStr(MLine, "%T") - 1) & Time & Right(MLine, Len(MLine) - InStr(MLine, "%T") - 1): FoundSth = True
        Loop While FoundSth
        If ShowToThisUser Then If MLine = "" Then TU vsock, EmptyLine Else TU vsock, MLine
      End If
    Loop
    Close #FileNum
    TU vsock, EmptyLine
    ShowMOTD = True
  Else
    ShowMOTD = False
  End If
End Function
Public Sub ShowDCStat(vsock As Long)
  Dim i As Long, Dummy As String
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "2*** Listing sockets:"
    TU vsock, "0,1 Nr | " & MakeLength("Handle", ServerNickLen) & " | " & MakeLength("Local", 24) & " | Way | " & MakeLength("Remote", 24) & " | State        "
    For i = LBound(SocketItem) To UBound(SocketItem)
      If IsValidSocket(i) Then
        If SocketItem(i).SocketNumber > 0 And SocketItem(i).SocketStatus > SS_Idle Then
          Dummy = SocketItem(i).RegNick
          If Dummy = "" Then Dummy = "<NONE>"
          TU vsock, " " & IIf(i < 10, "0", "") & i & " 14| " & IIf(Left(Dummy, 1) = "<", "14" & MakeLength(Dummy, ServerNickLen) & "", MakeLength(Dummy, ServerNickLen)) & " 14| " & MakeLength(WSAGetAscAddr(SocketItem(i).LocalAddress), 24) & " 14| " & IIf(SocketItem(i).SocketDirection = SD_In, "<--", "-->") & " 14| " & MakeLength(WSAGetAscAddr(SocketItem(i).RemoteAddress), 24) & " 14| " & IIf(SocketItem(i).SocketType = SOCK_STREAM, "TCP, ", "UDP, ") & IIf(SocketItem(i).SocketStatus = SS_Listening, "listen", IIf(SocketItem(i).SocketStatus = SS_Connected, "CONNECTED", "CONNECTING"))
        End If
      End If
    Next i
  Else
    TU vsock, "*** Listing sockets:"
    TU vsock, " Nr  " & MakeLength("Handle", ServerNickLen) & "  " & MakeLength("Local", 24) & "  Way    " & MakeLength("Remote", 24) & "  Type"
    TU vsock, " --  " & String(ServerNickLen, "-") & "  -----------------------  ---  -----------------------  -------------"
    For i = LBound(SocketItem) To UBound(SocketItem)
      If IsValidSocket(i) Then
        If SocketItem(i).SocketNumber > 0 And SocketItem(i).SocketStatus > SS_Idle Then
          TU vsock, " " & IIf(i < 10, "0", "") & i & "     " & MakeLength(IIf(SocketItem(i).RegNick <> "", SocketItem(i).RegNick, "<none>"), ServerNickLen) & "     " & MakeLength(WSAGetAscAddr(SocketItem(i).LocalAddress), 24) & "     " & MakeLength(WSAGetAscAddr(SocketItem(i).RemoteAddress), 24) & "     " & IIf(SocketItem(i).SocketType = SOCK_STREAM, "TCP, ", "UDP, ") & IIf(SocketItem(i).SocketStatus = SS_Listening, "listen", IIf(SocketItem(i).SocketStatus = SS_Connected, "CONNECTED", "CONNECTING"))
        End If
      End If
    Next i
  End If
End Sub
Public Sub bShowMOTD(ToBot As String, prefix As String, UserName As String, UserFlags As String) ' : AddStack "GUI_bShowMOTD(" & ToBot & ", " & prefix & ", " & UserName & ", " & UserFlags & ")"
Dim FileNum As Integer, MLine As String, FoundSth As Boolean, ShowToThisUser As Boolean
  On Local Error Resume Next
  If Dir(HomeDir & "Motd.txt") <> "" Then
    SendToBot ToBot, prefix & "*** Here's my MOTD file:"
    SendToBot ToBot, prefix
    FileNum = FreeFile: Open HomeDir & "Motd.txt" For Input As #FileNum
    ShowToThisUser = True
    Do While Not EOF(FileNum)
      Line Input #FileNum, MLine
      If InStr(MLine, "%{") > 0 And (InStr(MLine, "}") > InStr(MLine, "%{")) Then
        MLine = Mid(MLine, InStr(MLine, "%{") + 2, InStr(MLine, "}") - InStr(MLine, "%{") - 1)
        If (LCase(MLine) = "+color") Or (LCase(MLine) = "-color") Then
          ShowToThisUser = True
        Else
          ShowToThisUser = (CombineFlags(UserFlags, MLine) = UserFlags)
        End If
        FoundSth = False
      Else
        Do
          FoundSth = False
          If InStr(MLine, "%N") > 0 Then MLine = Left(MLine, InStr(MLine, "%N") - 1) & ToBot & Right(MLine, Len(MLine) - InStr(MLine, "%N") - 1): FoundSth = True
          If InStr(MLine, "%B") > 0 Then MLine = Left(MLine, InStr(MLine, "%B") - 1) & BotNetNick & Right(MLine, Len(MLine) - InStr(MLine, "%B") - 1): FoundSth = True
          If InStr(MLine, "%V") > 0 Then MLine = Left(MLine, InStr(MLine, "%V") - 1) & "AnGeL " & BotVersion & IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", "") & Right(MLine, Len(MLine) - InStr(MLine, "%V") - 1): FoundSth = True
          If InStr(MLine, "%C") > 0 Then MLine = Left(MLine, InStr(MLine, "%C") - 1) & BotChannels + Right(MLine, Len(MLine) - InStr(MLine, "%C") - 1): FoundSth = True
          If InStr(MLine, "%T") > 0 Then MLine = Left(MLine, InStr(MLine, "%T") - 1) & Time & Right(MLine, Len(MLine) - InStr(MLine, "%T") - 1): FoundSth = True
        Loop While FoundSth
        If ShowToThisUser Then SendToBot ToBot, prefix + MLine
      End If
    Loop
    Close #FileNum
    SendToBot ToBot, prefix
  Else
    SendToBot ToBot, prefix & "*** Sorry, no MOTD file found."
  End If
End Sub

Public Sub ChanList(vsock As Long, Line As String) ' : AddStack "GUI_ChanList(" & vsock & ", " & Line & ")"
  Dim u As Long, u2 As Long, CountOps As Long, CountVoices As Long, CountUsers As Long, CountHOps As Long
  Dim ChNum As Long, UserState As String, Messg As String, ChanStats As String
  Dim Status As String, FrontSpace As String
  
  If Line = "4" Then FrontSpace = "    " Else FrontSpace = ""
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, FrontSpace & "2*** Listing channels:"
    TU vsock, FrontSpace & "0,1 Nr | Channel:               | Rank:  | Users:              | Status:     "
  Else
    TU vsock, FrontSpace & "*** Listing channels:"
    TU vsock, FrontSpace & " #   Channel:                 Rank:    Users:                Status:"
    TU vsock, FrontSpace & " --  -----------------------  -------  --------------------  ------------"
  End If
  For u = 1 To PermChanCount
    Messg = PermChannels(u).Name
    ChNum = FindChan(Messg)
    If ChNum <> 0 Then
      CountUsers = 0: CountHOps = 0: CountVoices = 0: CountOps = 0
      For u2 = 1 To Channels(ChNum).UserCount
        Select Case Channels(ChNum).User(u2).Status
          Case "@", "@+", "@%", "@%+"
            CountOps = CountOps + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "3CHANOP"
          Case "%", "%+"
            CountHOps = CountHOps + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "10HALFOP"
          Case "+"
            CountVoices = CountVoices + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "7VOICE "
          Case Else
            CountUsers = CountUsers + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "4USER  "
        End Select
      Next u2
      If InStr(ServerUserModes, "%") <> 0 Then
        ChanStats = "@" & CStr(CountOps) & "14," & Spaces(3, CStr(CountOps)) & "%" & CStr(CountHOps) & "14," & Spaces(3, CStr(CountHOps)) & "+" & CStr(CountVoices) & "14," & Spaces(3, CStr(CountVoices)) & "/" & CStr(CountUsers) + Spaces(3, CStr(CountUsers))
      Else
        ChanStats = "@" & CStr(CountOps) & "14," & Spaces(3, CStr(CountOps)) & "+" & CStr(CountVoices) & "14," & Spaces(3, CStr(CountVoices)) & "/" & CStr(CountUsers) + Spaces(3, CStr(CountUsers)) & String(5, " ")
      End If
    Else
      UserState = "14--    "
      ChanStats = "                   "
    End If
    If Channels(ChNum).InFlood = False Then
      Select Case PermChannels(u).Status
        Case ChanStat_OK: Status = "3OK"
        Case ChanStat_ImBanned: Status = "4I'm banned!"
        Case ChanStat_NeedInvite: Status = "4It's +i!"
        Case ChanStat_NeedKey: Status = "4It's +k!"
        Case ChanStat_RegisteredOnly: Status = "4It's +r!"
        Case ChanStat_BadLimit: Status = "4It's full!"
        Case ChanStat_NotOn: Status = "10Joining..."
        Case ChanStat_Left: Status = "10I left."
        Case ChanStat_Duped: Status = "7Duped/Split"
        Case ChanStat_OutLimits: Status = "7Out of Limit"
        Case ChanStat_Unsup: Status = "7Unsupported"
        Case Else: Status = "4?!?"
      End Select
    Else
      Status = "4FLOOD"
    End If
    If Len(Messg) > 22 Then Messg = Left(Messg, 19) & "..."
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, FrontSpace & " " & IIf(u > 9, CStr(u), "0" & CStr(u)) & "14 | " & Messg & String(22 - Len(Messg), " ") & "14 | " & UserState & "14 | " & ChanStats & "14 | " & Status
    Else
      TU vsock, FrontSpace & " " & IIf(u > 9, CStr(u), "0" & CStr(u)) & "  " & Messg & String(22 - Len(Messg), " ") & "   " & Strip(UserState) & "   " & ChanStats & "   " & Status
    End If
  Next u
  For ChNum = 1 To ChanCount
    If (Left(Channels(ChNum).Name, 1) <> "&") And (IsPermChan(Channels(ChNum).Name) = False) Then
      Messg = Channels(ChNum).Name
      CountUsers = 0: CountHOps = 0: CountVoices = 0: CountOps = 0
      For u2 = 1 To Channels(ChNum).UserCount
        Select Case Channels(ChNum).User(u2).Status
          Case "@", "@+", "@%", "@%+"
            CountOps = CountOps + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "3CHANOP"
          Case "%", "%+"
            CountHOps = CountHOps + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "10HALFOP"
          Case "+"
            CountVoices = CountVoices + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "7VOICE "
          Case Else
            CountUsers = CountUsers + 1
            If Channels(ChNum).User(u2).Nick = MyNick Then UserState = "4USER  "
        End Select
      Next u2
      If InStr(ServerUserModes, "%") <> 0 Then
        ChanStats = "@" & CStr(CountOps) & "14," & Spaces(3, CStr(CountOps)) & "%" & CStr(CountHOps) & "14," & Spaces(3, CStr(CountHOps)) & "+" & CStr(CountVoices) & "14," & Spaces(3, CStr(CountVoices)) & "/" & CStr(CountUsers) + Spaces(3, CStr(CountUsers))
      Else
        ChanStats = "@" & CStr(CountOps) & "14," & Spaces(3, CStr(CountOps)) & "+" & CStr(CountVoices) & "14," & Spaces(3, CStr(CountVoices)) & "/" & CStr(CountUsers) + Spaces(3, CStr(CountUsers)) & String(5, " ")
      End If
      Status = "14Temporary"
      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
        TU vsock, FrontSpace & " " & IIf(u > 9, CStr(u), "0" & CStr(u)) & "14 | " & Messg & String(22 - Len(Messg), " ") & "14 | " & UserState & "14 | " & ChanStats & "14 | " & Status
      Else
        TU vsock, FrontSpace & " " & IIf(u > 9, CStr(u), "0" & CStr(u)) & "  " & Messg & String(22 - Len(Messg), " ") & "   " & Strip(UserState) & "   " & ChanStats & "   " & Status
      End If
      u = u + 1
    End If
  Next ChNum
  TU vsock, EmptyLine
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 2
End Sub

Public Sub WList(vsock As Long, Line As String) ' : AddStack "GUI_WList(" & vsock & ", " & Line & ")"
Dim u As Long, Temp As String, FileNum As Integer, SeeLine As String, WItem As String
Dim WText As String
On Local Error Resume Next
  
  If Dir(HomeDir & "Whatis.txt") = "" Then TU vsock, "5*** There are no whatis entries.": Exit Sub
  If Err.Number > 0 Then TU vsock, "5*** There are no whatis entries.": Exit Sub
  
  TU vsock, "2*** Listing whatis entries:"
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "0,1 Item:                         | Text:                            "
  Else
    TU vsock, " Item:                           Text:"
    TU vsock, " ------------------------------  ---------------------------------"
  End If
  FileNum = FreeFile: Open HomeDir & "Whatis.txt" For Input As #FileNum
    Do While Not EOF(FileNum)
      Line Input #FileNum, SeeLine
      WItem = AddSpaces(Param(SeeLine, 1))
      WText = AddSpaces(Param(SeeLine, 2))
      If Len(WItem) > 29 Then WItem = Left(WItem, 26) & "..."
      If Len(WText) > 32 Then WText = Left(WText, 29) & "..."
      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
        TU vsock, " " & WItem & " " & Spaces(29, WItem) & "14| " & WText
      Else
        TU vsock, " " & WItem & " " & Spaces(31, WItem) + WText
      End If
    Loop
  Close #FileNum
  TU vsock, EmptyLine
End Sub

Public Sub ListExcepts(vsock As Long, Line As String) ' : AddStack "GUI_ListExcepts(" & vsock & ", " & Line & ")"
Dim u As Long, u2 As Long, Temp As String, CheckDesc As String, DPart As String, FirstLine As Boolean
Dim Stripped As String
  If ExceptCount = 0 Then TU vsock, "5*** There are no excepts.": Exit Sub
  
  TU vsock, "2*** Listing excepts:"
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "0,1 Nr: | Except information:                                 "
  Else
    TU vsock, " Nr:   Hostmask:                                        "
    TU vsock, " ----  -------------------------------------------------"
  End If
  For u = 1 To ExceptCount
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      Temp = String(5 - Len(CStr(u)), " ") + CStr(u) & "14 |10 Hostmask:3 " & Excepts(u).Hostmask & " 2(" & IIf(Excepts(u).Channel = "*", "global", Excepts(u).Channel) & ")"
      TU vsock, Temp + Spaces(65, Temp) & "14 "
      Temp = String(4, " ") & "14 |10 Created : " & Excepts(u).CreatedBy & " 14(" & Format(CDate(Excepts(u).CreatedAt), "dd.mm.yy, hh:nn") & ")"
      TU vsock, Temp + Spaces(66, Temp) & "14 |"
      DPart = "": FirstLine = True
      Stripped = String(4, " ") & "14 |10           " & Strip(DPart)
      Temp = String(4, " ") & "14 |10           " & DPart
      TU vsock, Temp + Spaces(62, Stripped) & "14 |"
    Else
      Temp = String(5 - Len(CStr(u)), " ") + CStr(u) & "   Hostmask: " & Excepts(u).Hostmask & " (" & IIf(Excepts(u).Channel = "*", "global", Excepts(u).Channel) & ")"
      TU vsock, Temp + Spaces(55, Temp) & "   "
      TU vsock, String(4, " ") & "   Created : " & Excepts(u).CreatedBy & " (" & Format(CDate(Excepts(u).CreatedAt), "dd.mm.yy, hh:nn") & ")"
    End If
  Next u
  TU vsock, EmptyLine
End Sub

Public Sub ListInvites(vsock As Long, Line As String) ' : AddStack "GUI_ListInvites(" & vsock & ", " & Line & ")"
Dim u As Long, u2 As Long, Temp As String, CheckDesc As String, DPart As String, FirstLine As Boolean
Dim Stripped As String
  If InviteCount = 0 Then TU vsock, "5*** There are no invitations.": Exit Sub
  TU vsock, "2*** Listing invitations:"
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "0,1 Nr: | Invite information:                                 "
  Else
    TU vsock, " Nr:   Hostmask:                                        "
    TU vsock, " ----  -------------------------------------------------"
  End If
  For u = 1 To InviteCount
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      Temp = String(5 - Len(CStr(u)), " ") + CStr(u) & "14 |10 Hostmask:3 " & Invites(u).Hostmask & " 2(" & IIf(Invites(u).Channel = "*", "global", Invites(u).Channel) & ")"
      TU vsock, Temp + Spaces(65, Temp) & "14 "
      Temp = String(4, " ") & "14 |10 Created : " & Invites(u).CreatedBy & " 14(" & Format(CDate(Invites(u).CreatedAt), "dd.mm.yy, hh:nn") & ")"
      TU vsock, Temp + Spaces(66, Temp) & "14 |"
      DPart = "": FirstLine = True
      Stripped = String(4, " ") & "14 |10           " & Strip(DPart)
      Temp = String(4, " ") & "14 |10           " & DPart
      TU vsock, Temp + Spaces(62, Stripped) & "14 |"
    Else
      Temp = String(5 - Len(CStr(u)), " ") + CStr(u) & "   Hostmask: " & Invites(u).Hostmask & " (" & IIf(Invites(u).Channel = "*", "global", Invites(u).Channel) & ")"
      TU vsock, Temp + Spaces(55, Temp) & "   "
      TU vsock, String(4, " ") & "   Created : " & Invites(u).CreatedBy & " (" & Format(CDate(Invites(u).CreatedAt), "dd.mm.yy, hh:nn") & ")"
    End If
  Next u
  TU vsock, EmptyLine
End Sub
Public Sub ListBans(vsock As Long, Line As String) ' : AddStack "GUI_ListBans(" & vsock & ", " & Line & ")"
Dim u As Long, u2 As Long, Temp As String, CheckDesc As String, DPart As String, FirstLine As Boolean
Dim Stripped As String
  If BanCount = 0 Then TU vsock, "5*** There are no bans.": Exit Sub
  
  TU vsock, "2*** Listing bans:"
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "0,1 Nr: | Ban information:                                 | Info:   "
  Else
    TU vsock, " Nr:   Hostmask:                                          Info:"
    TU vsock, " ----  -------------------------------------------------  -------"
  End If
  For u = 1 To BanCount
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      Temp = String(5 - Len(CStr(u)), " ") + CStr(u) & "14 |10 Hostmask:3 " & Bans(u).Hostmask & " 2(" & IIf(Bans(u).Channel = "*", "global", Bans(u).Channel) & ")"
      TU vsock, Temp + Spaces(65, Temp) & "14 | " & IIf(Bans(u).Sticky = True, "3sticky", "normal")
      Temp = String(4, " ") & "14 |10 Created : " & Bans(u).CreatedBy & " 14(" & Format(CDate(Bans(u).CreatedAt), "dd.mm.yy, hh:nn") & ")"
      TU vsock, Temp + Spaces(66, Temp) & "14 |"
      CheckDesc = Bans(u).Comment & " "
      DPart = "": FirstLine = True
      For u2 = 1 To Len(CheckDesc)
        If Mid(CheckDesc, u2, 1) <> " " Then
          DPart = DPart + Mid(CheckDesc, u2, 1)
        Else
          If InStr(u2 + 1, CheckDesc, " ") - u2 > 38 - Len(DPart) Then
            DPart = Trim(DPart)
            If FirstLine Then
              Stripped = String(4, " ") & "14 |10 Comment : " & Strip(DPart)
              Temp = String(4, " ") & "14 |10 Comment : " & DPart
              TU vsock, Temp + Spaces(62, Stripped) & "14 |"
              FirstLine = False
            Else
              Stripped = String(4, " ") & "14 |10           " & Strip(DPart)
              Temp = String(4, " ") & "14 |10           " & DPart
              TU vsock, Temp + Spaces(62, Stripped) & "14 |"
            End If
            DPart = ""
          Else
            DPart = DPart & " "
          End If
        End If
      Next u2
      DPart = Trim(DPart)
      If FirstLine Then
        Stripped = String(4, " ") & "14 |10 Comment : " & Strip(DPart)
        Temp = String(4, " ") & "14 |10 Comment : " & DPart
        TU vsock, Temp + Spaces(62, Stripped) & "14 |"
        FirstLine = False
      Else
        Stripped = String(4, " ") & "14 |10           " & Strip(DPart)
        Temp = String(4, " ") & "14 |10           " & DPart
        TU vsock, Temp + Spaces(62, Stripped) & "14 |"
      End If
    Else
      Temp = String(5 - Len(CStr(u)), " ") + CStr(u) & "   Hostmask: " & Bans(u).Hostmask & " (" & IIf(Bans(u).Channel = "*", "global", Bans(u).Channel) & ")"
      TU vsock, Temp + Spaces(55, Temp) & "   " & IIf(Bans(u).Sticky = True, "sticky", "normal")
      TU vsock, String(4, " ") & "   Created : " & Bans(u).CreatedBy & " (" & Format(CDate(Bans(u).CreatedAt), "dd.mm.yy, hh:nn") & ")"
      TU vsock, String(4, " ") & "   Comment : " & Bans(u).Comment
    End If
  Next u
  TU vsock, EmptyLine
End Sub

Public Sub seen(vsock As Long, Line As String) ' : AddStack "GUI_seen(" & vsock & ", " & Line & ")"
Dim Rest As String, RegUser As String, ChNum As Long, u As Long, Chan As String
Dim MatchedOne As Boolean
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".seen <nick>"): Exit Sub
  Rest = LastSeen(Param(Replace(GetRest(Line, 2), ", ", ","), 1), SocketItem(vsock).RegNick, "", SocketItem(vsock).RegNick, MatchedOne)
  If Rest <> "" Then
    TU vsock, "3*** " & Rest
  Else
    TU vsock, "5*** Invalid nick."
  End If
End Sub

Public Sub ChanBans(vsock As Long, Line As String)
  Dim Chan As String, ChNum As Integer, u As Integer
  Chan = Param(Line, 2)
  If Chan = "" Then TU vsock, "5*** Usage: .chanbans <[" & ServerChannelPrefixes & "]channel>": Exit Sub
  ChNum = FindChan(Chan)
  If ChNum > 0 Then
    If Channels(ChNum).BanCount > 0 Then
      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
        TU vsock, "2*** Listing channel bans:"
        TU vsock, "0,1 Nr: | " & Spaces2(45, "Ban information:") & " | " & Spaces2(20, "Created at:") & " "
        For u = 1 To Channels(ChNum).BanCount
          TU vsock, " " & SpacesC(3, "", CStr(u)) & " 14|2 " & Spaces2(45, Channels(ChNum).BanList(u).Mask) & " 14|3 " & Spaces2(20, CStr(Channels(ChNum).BanList(u).CreatedAt))
        Next u
      Else
        TU vsock, "*** Listing channel bans:"
        TU vsock, " Nr:   " & Spaces2(45, "Ban information:") & "   " & Spaces2(20, "Created at:")
        TU vsock, " ----  ----------------------------------------------  --------------------"
        For u = 1 To Channels(ChNum).BanCount
          TU vsock, " " & SpacesC(3, "", CStr(u)) & "   " & Spaces2(45, Channels(ChNum).BanList(u).Mask) & "   " & Spaces2(20, CStr(Channels(ChNum).BanList(u).CreatedAt))
        Next u
      End If
    Else
      TU vsock, "5*** There are no bans."
    End If
  Else
    TU vsock, "5*** I'm not on " & Chan & "."
  End If
End Sub

Public Sub Channel(vsock As Long, Line As String) ' : AddStack "GUI_Channel(" & vsock & ", " & Line & ")"
Dim Rest As String, RegUser As String, ChNum As Long, u As Long, Chan As String, u2 As Long
Dim FUsers() As Long, FUserCount As Long, u3 As Long, ClNick As String, Host As String
Dim NUserStat As Long, OUserStat As Long, SortedIn As Boolean, CountOps As Long
Dim CountVoices As Long, CountUsers As Long, ChanStats As String, CountHOps As Long
  ReDim Preserve FUsers(5)
  Chan = Param(Line, 2)
  If Chan = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".channel <[" & ServerChannelPrefixes & "]channel> (+)"): Exit Sub
  If MatchFlags(GetUserChanFlags(SocketItem(vsock).RegNick, Chan), "-o") Then TU vsock, "5*** Sorry, you don't have op rights for this channel.": Exit Sub
  ChNum = FindChan(Chan)
  If ChNum > 0 Then
    For u = 1 To Channels(ChNum).UserCount
      Select Case Channels(ChNum).User(u).Status
        Case "@", "@+", "@%", "@%+": NUserStat = 4
        Case "%", "%+": NUserStat = 3
        Case "+": NUserStat = 2
        Case Else: NUserStat = 1
      End Select
      SortedIn = False
      For u2 = 1 To FUserCount
        Select Case Channels(ChNum).User(FUsers(u2)).Status
          Case "@", "@+", "@%", "@%+": OUserStat = 4
          Case "%", "%+": OUserStat = 3
          Case "+": OUserStat = 2
          Case Else: OUserStat = 1
        End Select
        If NUserStat > OUserStat Then
          FUserCount = FUserCount + 1: If FUserCount > UBound(FUsers()) Then ReDim Preserve FUsers(UBound(FUsers()) + 5)
          For u3 = FUserCount To u2 + 1 Step -1
            FUsers(u3) = FUsers(u3 - 1)
          Next u3
          FUsers(u2) = u
          SortedIn = True
          Exit For
        ElseIf NUserStat = OUserStat Then
          If UCase(Channels(ChNum).User(FUsers(u2)).Nick) > UCase(Channels(ChNum).User(u).Nick) Then
            FUserCount = FUserCount + 1: If FUserCount > UBound(FUsers()) Then ReDim Preserve FUsers(UBound(FUsers()) + 5)
            For u3 = FUserCount To u2 + 1 Step -1
              FUsers(u3) = FUsers(u3 - 1)
            Next u3
            FUsers(u2) = u
            SortedIn = True
            Exit For
          End If
        End If
      Next u2
      If Not SortedIn Then
        FUserCount = FUserCount + 1: If FUserCount > UBound(FUsers()) Then ReDim Preserve FUsers(UBound(FUsers()) + 5)
        FUsers(FUserCount) = u
      End If
    Next u
    For u2 = 1 To Channels(ChNum).UserCount
      Select Case Channels(ChNum).User(u2).Status
        Case "@", "@+", "@%", "@%+": CountOps = CountOps + 1
        Case "%", "%+": CountHOps = CountHOps + 1
        Case "+": CountVoices = CountVoices + 1
        Case Else: CountUsers = CountUsers + 1
      End Select
    Next u2
    If InStr(ServerUserModes, "%") <> 0 Then
      ChanStats = "@" & CStr(CountOps) & "14,%" & CStr(CountHOps) & "14,+" & CStr(CountVoices) & "14,/" & CStr(CountUsers)
    Else
      ChanStats = "@" & CStr(CountOps) & "14,+" & CStr(CountVoices) & "14,/" & CStr(CountUsers)
    End If
    TU vsock, "2*** Channel info for " & Channels(ChNum).Name & ":"
    TU vsock, " Modes: Currently [" & Channels(ChNum).Mode & "], enforcing [" & GetChannelSetting(Channels(ChNum).Name, "EnforceModes", "") & "]"
    TU vsock, " Users: " & CStr(Channels(ChNum).UserCount) & " user" & IIf(Channels(ChNum).UserCount = 1, "", "s") & " 3(" & ChanStats & "3)"
    If Channels(ChNum).Topic <> "" Then TU vsock, " Topic: " & Channels(ChNum).Topic
    TU vsock, EmptyLine
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, "0,1 " & Spaces2(ServerNickLen, "User:") & "  | " & Spaces2(ServerNickLen, "Handle:") & " | Hostmask:                          "
    Else
      TU vsock, " " & Spaces2(ServerNickLen, "User:") & "    " & Spaces2(ServerNickLen, "Handle:") & "   Hostmask:"
      TU vsock, " " & String(ServerNickLen, "-") & "-   " & String(ServerNickLen, "-") & "   -------------------------------------"
    End If
    For u = 1 To FUserCount
      ClNick = Channels(ChNum).User(FUsers(u)).Nick
      Select Case Channels(ChNum).User(FUsers(u)).Status
        Case "@", "@+", "@%", "@%+": ClNick = "@" & ClNick
        Case "%", "%+": ClNick = "%" & ClNick
        Case "+": ClNick = "+" & ClNick
        Case Else: ClNick = " " & ClNick
      End Select
      RegUser = Channels(ChNum).User(FUsers(u)).RegNick
      Host = Mask(Channels(ChNum).User(FUsers(u)).Hostmask, 10)
      If (Param(Line, 3) = "+") And (Channels(ChNum).User(FUsers(u)).IPmask <> "") Then
        If GetSockFlag(vsock, SF_Colors) = SF_YES Then
          If Channels(ChNum).User(FUsers(u)).Nick = MyNick Then RegUser = "<-- me!"
          TU vsock, " " & Spaces2(ServerNickLen + 1, ClNick) & "14 | " & Spaces2(ServerNickLen, RegUser) & "14 | " & Mask(Channels(ChNum).User(FUsers(u)).IPmask, 10)
        Else
          TU vsock, " " & Spaces2(ServerNickLen + 1, ClNick) & "   " & Spaces2(ServerNickLen, RegUser) & "   " & Mask(Channels(ChNum).User(FUsers(u)).IPmask, 10)
        End If
      Else
        If GetSockFlag(vsock, SF_Colors) = SF_YES Then
          If Channels(ChNum).User(FUsers(u)).Nick = MyNick Then RegUser = "<-- me!"
          TU vsock, " " & Spaces2(ServerNickLen + 1, ClNick) & "14 | " & Spaces2(ServerNickLen, RegUser) & "14 | " & Host
        Else
          TU vsock, " " & Spaces2(ServerNickLen + 1, ClNick) & "   " & Spaces2(ServerNickLen, RegUser) & "   " & Host
        End If
      End If
    Next u
    TU vsock, EmptyLine
  Else
    TU vsock, "5*** I'm not on " & Chan & "."
  End If
End Sub

Public Sub ListBotNetUsers(vsock As Long) ' : AddStack "GUI_ListBotNetUsers(" & vsock & ")"
Dim AllBots As Long, passes As Byte, passmatch As Byte
Dim cur As Long, Level As String, OnBotName As String, Rest As String
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "2*** Listing users: (14Grey2 users are away)"
    TU vsock, "0,1 " & Spaces2(ServerNickLen, "User:") & " | LVL | " & Spaces2(ServerNickLen, "Bot:") & " | Hostmask / Away reason:            "
  Else
    TU vsock, "*** Listing users: (Users with ""*"" are away)"
    TU vsock, " " & Spaces2(ServerNickLen, "User:") & "  LVL:  " & Spaces2(ServerNickLen, "Bot:") & "  Hostmask / Away reason:"
    TU vsock, " " & String(ServerNickLen, "-") & "  ----  " & String(ServerNickLen, "-") & "  -------------------------------------"
  End If
  'Alle Bots durchgehen
  For AllBots = 1 To BotCount
    'User auf der Partyline, schön sortiert
    For passes = 1 To 4
      For cur = 1 To SocketCount
        If IsValidSocket(cur) Then
          If SocketItem(cur).IRCNick <> "²*SCRIPT*²" Then
            If ((AllBots = 1) And (GetSockFlag(cur, SF_LocalVisibleUser) = SF_YES)) Or ((AllBots > 1) And (LCase(SocketItem(cur).OnBot) = LCase(Bots(AllBots).Nick))) Then
              Level = "3USR": passmatch = 4
              If MatchFlags(SocketItem(cur).Flags, "+t") Then Level = "10NET": passmatch = 3
              If MatchFlags(SocketItem(cur).Flags, "+m") Then Level = "12MAS": passmatch = 2
              If MatchFlags(SocketItem(cur).Flags, "+n") Then Level = "4OWN": passmatch = 1
              If (passmatch = passes) And (SocketItem(cur).AwayMessage = "") Then
                OnBotName = SocketItem(cur).OnBot
                If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                  Rest = SocketItem(cur).RegNick
                  TU vsock, " " & Spaces2(ServerNickLen, Rest) & "14 | " & Level & "14 | " & Spaces2(ServerNickLen, OnBotName) & "14 | " & Mask(SocketItem(cur).Hostmask, 10)
                Else
                  Rest = SocketItem(cur).RegNick
                  TU vsock, " " & Spaces2(ServerNickLen, Rest) & "  " & Strip(Level) & "   " & Spaces2(ServerNickLen, OnBotName) & "  " & Mask(SocketItem(cur).Hostmask, 10)
                End If
              End If
            End If
          End If
        End If
      Next cur
    Next passes
  Next AllBots
  'User, die Away sind
  For cur = 1 To SocketCount
    If IsValidSocket(cur) Then
      If SocketItem(cur).IRCNick <> "²*SCRIPT*²" Then
        Level = "USR"
        If MatchFlags(SocketItem(cur).Flags, "+t") Then Level = "NET"
        If MatchFlags(SocketItem(cur).Flags, "+m") Then Level = "MAS"
        If MatchFlags(SocketItem(cur).Flags, "+n") Then Level = "OWN"
        If SocketItem(cur).AwayMessage <> "" Then  'Wenn AwayMessage gesetzt ist, muß cur ein User sein
          OnBotName = SocketItem(cur).OnBot
          If GetSockFlag(vsock, SF_Colors) = SF_YES Then
            Rest = SocketItem(cur).RegNick
            TU vsock, "14 " & Spaces2(ServerNickLen, Rest) & " | " & Level & " | " & Spaces2(ServerNickLen, OnBotName) & " | " & SpacesC(34, SocketItem(cur).AwayMessage, " [AWAY " & LTrim(TimeSpan2(SocketItem(cur).LastEvent)) & "]")
          Else
            Rest = SocketItem(cur).RegNick
            TU vsock, "*" & Spaces2(ServerNickLen, Rest) & "  " & Strip(Level) & "   " & Spaces2(ServerNickLen, OnBotName) & "  " & SpacesC(34, SocketItem(cur).AwayMessage, " [AWAY " & LTrim(TimeSpan2(SocketItem(cur).LastEvent)) & "]")
          End If
        End If
      End If
    End If
  Next cur
  TU vsock, EmptyLine
End Sub

Public Sub ListBotUsers(vsock As Long) ' : AddStack "GUI_ListBotUsers(" & vsock & ")"
Dim passes As Byte, passmatch As Byte, cur As Long, Level As String
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "2*** Listing users: (14Grey2 users are away)"
    TU vsock, "0,1 " & Spaces2(ServerNickLen, "User:") & " | LVL | Hostmask / Away reason:                       "
  Else
    TU vsock, "*** Listing users: (Users with ""*"" are away)"
    TU vsock, " " & Spaces2(ServerNickLen, "User:") & "   LVL:  Hostmask / Away reason:"
    TU vsock, " " & String(ServerNickLen, "-") & "-  ----  ----------------------------------------------"
  End If
  'User auf der Partyline, schön sortiert
  For passes = 1 To 4
    For cur = 1 To SocketCount
      If IsValidSocket(cur) Then
        Level = "3USR": passmatch = 4
        If MatchFlags(SocketItem(cur).Flags, "+t") Then Level = "10NET": passmatch = 3
        If MatchFlags(SocketItem(cur).Flags, "+m") Then Level = "12MAS": passmatch = 2
        If MatchFlags(SocketItem(cur).Flags, "+n") Then Level = "4OWN": passmatch = 1
        If (passmatch = passes) And (GetSockFlag(cur, SF_LocalVisibleUser) = SF_YES) And (SocketItem(cur).AwayMessage = "") Then
          If GetSockFlag(vsock, SF_Colors) = SF_YES Then
            TU vsock, " " & Spaces2(ServerNickLen, SocketItem(cur).RegNick) & "14 | " & Level & "14 | " & Mask(SocketItem(cur).Hostmask, 10)
          Else
            TU vsock, " " & Spaces2(ServerNickLen, SocketItem(cur).RegNick) & "   " & Strip(Level) & "   " & Mask(SocketItem(cur).Hostmask, 10)
          End If
        End If
      End If
    Next cur
  Next passes
  'User, die Away sind
  For cur = 1 To SocketCount
    If IsValidSocket(cur) Then
      Level = "USR"
      If MatchFlags(SocketItem(cur).Flags, "+t") Then Level = "NET"
      If MatchFlags(SocketItem(cur).Flags, "+m") Then Level = "MAS"
      If MatchFlags(SocketItem(cur).Flags, "+n") Then Level = "OWN"
      If (GetSockFlag(cur, SF_LocalVisibleUser) = SF_YES) And (SocketItem(cur).AwayMessage <> "") Then
        If GetSockFlag(vsock, SF_Colors) = SF_YES Then
          TU vsock, "14 " & Spaces2(ServerNickLen, SocketItem(cur).RegNick) & "14 |14 " & Level & "14 |14 " & SpacesC(45, SocketItem(cur).AwayMessage, " [AWAY " & LTrim(TimeSpan2(SocketItem(cur).LastEvent)) & "]")
        Else
          TU vsock, " " & Spaces2(ServerNickLen + 1, SocketItem(cur).RegNick & "*") & "  " & Strip(Level) & "   " & SpacesC(45, SocketItem(cur).AwayMessage, " [AWAY " & LTrim(TimeSpan2(SocketItem(cur).LastEvent)) & "]")
        End If
      End If
    End If
  Next cur
  If BotCount > 1 Then
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, "14-" & String(ServerNickLen, "-") & "-|-----|-----------------------------------------------"
    Else
      TU vsock, EmptyLine
    End If
    For cur = 1 To SocketCount
      If IsValidSocket(cur) Then
        If GetSockFlag(cur, SF_Status) = SF_Status_Bot Then
          If GetSockFlag(vsock, SF_Colors) = SF_YES Then
            TU vsock, " " & Spaces2(ServerNickLen, SocketItem(cur).RegNick) & "14 |14 BOT14 | " & IIf(SocketItem(cur).RemotePort <> 0, "->", "<-") & " " & Mask(SocketItem(cur).Hostmask, 11)
          Else
            TU vsock, " " & Spaces2(ServerNickLen, SocketItem(cur).RegNick) & "   BOT   " & IIf(SocketItem(cur).RemotePort <> 0, "->", "<-") & " " & Mask(SocketItem(cur).Hostmask, 11)
          End If
        End If
      End If
    Next cur
  End If
  TU vsock, EmptyLine
End Sub

Public Sub PListBotUsers(ToBot As String, prefix As String) ' : AddStack "GUI_PListBotUsers(" & ToBot & ", " & prefix & ")"
Dim passes As Byte, passmatch As Byte, cur As Long, Level As String, FoundOne As Boolean
  For cur = 1 To SocketCount
    If IsValidSocket(cur) Then If GetSockFlag(cur, SF_Status) = SF_Status_Party Then FoundOne = True: Exit For
  Next cur
  If Not FoundOne Then
    SendToBot ToBot, prefix & "Party line members:  (* = away)"
    SendToBot ToBot, prefix & " -none-"
    Exit Sub
  End If
  SendToBot ToBot, prefix & "Party line members:  (* = away)"
  SendToBot ToBot, prefix & " User:        LVL:  Hostmask:"
  SendToBot ToBot, prefix & " -----------  ----  ----------------------------------------------"
  'User auf der Partyline, schön sortiert
  For passes = 1 To 4
    For cur = 1 To SocketCount
      If IsValidSocket(cur) Then
        Level = "3USR": passmatch = 4
        If MatchFlags(SocketItem(cur).Flags, "+t") Then Level = "10NET": passmatch = 3
        If MatchFlags(SocketItem(cur).Flags, "+m") Then Level = "12MAS": passmatch = 2
        If MatchFlags(SocketItem(cur).Flags, "+n") Then Level = "4OWN": passmatch = 1
        If (passmatch = passes) And (GetSockFlag(cur, SF_LocalVisibleUser) = SF_YES) And (SocketItem(cur).AwayMessage = "") Then
          SendToBot ToBot, prefix & " " & SocketItem(cur).RegNick + Spaces(10, SocketItem(cur).RegNick) & "   " & Strip(Level) & "   " & Mask(SocketItem(cur).Hostmask, 0)
        End If
      End If
    Next cur
  Next passes
  'User, die Away sind
  For cur = 1 To SocketCount
    If IsValidSocket(cur) Then
      Level = "USR"
      If MatchFlags(SocketItem(cur).Flags, "+t") Then Level = "NET"
      If MatchFlags(SocketItem(cur).Flags, "+m") Then Level = "MAS"
      If MatchFlags(SocketItem(cur).Flags, "+n") Then Level = "OWN"
      If (GetSockFlag(cur, SF_LocalVisibleUser) = SF_YES) And (SocketItem(cur).AwayMessage <> "") Then
        SendToBot ToBot, prefix & " " & SocketItem(cur).RegNick & "*" & Spaces(9, SocketItem(cur).RegNick) & "   " & Strip(Level) & "   " & Mask(SocketItem(cur).Hostmask, 0)
      End If
    End If
  Next cur
End Sub

Public Sub Colors(vsock As Long, Line As String) ' : AddStack "GUI_Colors(" & vsock & ", " & Line & ")"
Dim NewValue As String
  If LCase(Param(Line, 2)) <> "on" And LCase(Param(Line, 2)) <> "off" Then TU vsock, MakeMsg(ERR_CommandUsage, ".colors <on/off>"): Exit Sub
  NewValue = IIf(LCase(Param(Line, 2)) = "on", SF_YES, SF_NO)
  SetSockFlag vsock, SF_Colors, NewValue
  SetUserData SocketItem(vsock).UserNum, "colors", NewValue
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "3*** You will now see 12colors3 here!"
  Else
    TU vsock, "*** All colors turned off."
  End If
End Sub

Public Sub Match(vsock As Long, Line As String) ' : AddStack "GUI_Match(" & vsock & ", " & Line & ")"
Dim passes As Byte, passmatch As Byte, Level As String
Dim CheckNick As String, CheckHost As String, HostNum As Long
Dim OtherUserFlags As String, Matched As Boolean, UsNum As Long
Dim MatchType As Long, MatchString As String, MatchChan As String
Dim MatchedUsers() As TMatch, MatchedUserCount As Long
Dim LongString As String, LongCount As Integer, LongMatch As Boolean
Dim MHostmask As String, MGlobalFlags As String, MChannelFlags As String
Dim MChannelMatch As String, MBotFlags As String, MMatchString As String
Dim SplNick As String, SplIdent As String, SplHost As String
Dim u As Long
  ReDim Preserve MatchedUsers(5)
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".(l)match (hostmask) (+/-<flags> ([" & ServerChannelPrefixes & "]channel)) (+/-<botflags> bot)"): Exit Sub
  
  ConvertMatchString GetRest(Line, 2), MHostmask, MGlobalFlags, MChannelFlags, MChannelMatch, MBotFlags
  If MHostmask <> "" Then MMatchString = MMatchString + IIf(MMatchString = "", "", ", ") & "hostmask """ & MHostmask & """"
  If MGlobalFlags <> "" Then MMatchString = MMatchString + IIf(MMatchString = "", "", ", ") & "global """ & MGlobalFlags & """"
  If MBotFlags <> "" Then MMatchString = MMatchString + IIf(MMatchString = "", "", ", ") & "bot """ & MBotFlags & """"
  If (MChannelFlags <> "") And Not (IsValidChannel(MChannelMatch)) Then
    MMatchString = MMatchString + IIf(MMatchString = "", "", ", ") & "channel """ & MChannelFlags & """"
  Else
    MMatchString = MMatchString + IIf(MMatchString = "", "", ", ") & "channel """ & MChannelFlags & """ in """ & MChannelMatch & """"
  End If
  
  If MMatchString = "" Then TU vsock, "5*** Invalid match string! Take a look at '.help match'.": Exit Sub
  
  If MMatchString = "hostmask """ & MHostmask & """" Then
    If Mask(MHostmask, 0) = "*!*@*" Then
      If (InStr(Mask(MHostmask, 13), "*") = 0) And (InStr(Mask(MHostmask, 13), "?") = 0) Then
        Whois vsock, ".match " & Mask(MHostmask, 13), False
        Exit Sub
      End If
    End If
  End If
  
  'Remove ident chars in match string
  If MHostmask <> "" Then
    If StrictHost = False Then
      SplitHostmask MHostmask, SplNick, SplIdent, SplHost
      If InStr("~-+^=", Left(SplIdent, 1)) > 0 Then If Len(SplIdent) > 1 Then SplIdent = Mid(SplIdent, 2): MHostmask = SplNick & "!" & SplIdent & "@" & SplHost
    End If
  End If
  
  LongMatch = (LCase(Param(Line, 1)) = ".lmatch")
  
  For UsNum = 1 To BotUserCount
    CheckNick = BotUsers(UsNum).Name
    Matched = True
    If (Matched = True) And (MHostmask <> "") Then
      For HostNum = 1 To BotUsers(UsNum).HostMaskCount
        CheckHost = BotUsers(UsNum).HostMasks(HostNum)
        SplitHostmask CheckHost, SplNick, SplIdent, SplHost
        If SplNick = "*" Then SplNick = CheckNick
        If MatchWM2(MHostmask, SplNick & "!" & SplIdent & "@" & SplHost) = True Then Exit For
        If HostNum = BotUsers(UsNum).HostMaskCount Then CheckHost = "": Matched = False
      Next HostNum
      If BotUsers(UsNum).HostMaskCount = 0 Then
        CheckHost = "none"
        If Mask(MHostmask, 10) <> "*@*" Then
          Matched = False
        ElseIf MatchWM2(MHostmask, CheckNick & "!*@*") = False Then
          Matched = False
        End If
      End If
    Else
      CheckHost = IIf(BotUsers(UsNum).HostMaskCount > 0, BotUsers(UsNum).HostMasks(1), "none")
    End If
    If (Matched = True) And (MGlobalFlags <> "") Then
      If MatchFlags(BotUsers(UsNum).Flags, MGlobalFlags) = False Then Matched = False
    End If
    If (Matched = True) And (MBotFlags <> "") Then
      If MatchFlags(BotUsers(UsNum).BotFlags, MBotFlags) = False Then Matched = False
    End If
    If (Matched = True) And (MChannelFlags <> "") Then
      For u = 1 To BotUsers(UsNum).ChannelFlagCount
        If ChannelMatch(MChannelMatch, BotUsers(UsNum).ChannelFlags(u).Channel) = True Then
          If MatchFlags(CombineFlags(BotUsers(UsNum).Flags, "+" & BotUsers(UsNum).ChannelFlags(u).Flags + ChattrChanges(BotUsers(UsNum).ChannelFlags(u).Flags)), MChannelFlags) = True Then Exit For
        End If
        If u = BotUsers(UsNum).ChannelFlagCount Then
          Matched = MatchFlags(BotUsers(UsNum).Flags, MChannelFlags)
        End If
      Next u
      If BotUsers(UsNum).ChannelFlagCount = 0 Then
        Matched = MatchFlags(BotUsers(UsNum).Flags, MChannelFlags)
      End If
    End If
    
    If Matched = True Then
      MatchedUserCount = MatchedUserCount + 1
      If MatchedUserCount > UBound(MatchedUsers()) Then ReDim Preserve MatchedUsers(UBound(MatchedUsers()) + 5)
      MatchedUsers(MatchedUserCount).UserNum = UsNum
      MatchedUsers(MatchedUserCount).Hostmask = CheckHost
    End If
  Next UsNum
  
  If MatchedUserCount = 0 Then
    TU vsock, "5*** Sorry, I couldn't find a user matching this."
    Exit Sub
  ElseIf (LongMatch = True) And (MatchedUserCount > 20) Then
    TU vsock, "5*** Sorry, too many matches for a long match. Try a normal match."
    Exit Sub
  End If
  If Not LongMatch Then
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, "2*** MATCHING " & MMatchString
      TU vsock, "0,1 User:      | LVL | Hostmask:                                     "
    Else
      TU vsock, "*** MATCHING " & MMatchString
      TU vsock, " User:        LVL:  Hostmask:"
      TU vsock, " -----------  ----  ----------------------------------------------"
    End If
  End If
  For passes = 1 To 5
    For UsNum = 1 To MatchedUserCount
      OtherUserFlags = BotUsers(MatchedUsers(UsNum).UserNum).Flags
      CheckNick = BotUsers(MatchedUsers(UsNum).UserNum).Name
      CheckHost = MatchedUsers(UsNum).Hostmask
      Level = "3USR": passmatch = 4
      If MatchFlags(OtherUserFlags, "+t") Then Level = "10NET": passmatch = 3
      If MatchFlags(OtherUserFlags, "+m") Then Level = "12MAS": passmatch = 2
      If MatchFlags(OtherUserFlags, "+n") Then Level = "4OWN": passmatch = 1
      If MatchFlags(OtherUserFlags, "+b") Then Level = "14BOT": passmatch = 5
      If passmatch = passes Then
        If LongMatch Then
          Whois vsock, ".whois " & CheckNick, True
        Else
          LongCount = LongCount + 1
          If LongString <> "" Then LongString = LongString + IIf(GetSockFlag(vsock, SF_LF_ONLY) = SF_YES, vbLf, vbCrLf)
          If GetSockFlag(vsock, SF_Colors) = SF_YES Then
            LongString = LongString & " " & CheckNick + Spaces(10, CheckNick) & "14 | " & Level & "14 | " & CheckHost
          Else
            LongString = LongString & " " & CheckNick + Spaces(10, CheckNick) & "   " & Strip(Level) & "   " & CheckHost
          End If
        End If
      End If
      If LongCount > 19 Then TU vsock, LongString: LongString = "": LongCount = 0
    Next UsNum
  Next passes
  If LongCount > 0 Then TU vsock, LongString: LongString = "": LongCount = 0
  If MatchedUserCount > 1 Then TU vsock, "2*** Found " & Trim(Str(MatchedUserCount)) & " matches." Else TU vsock, "2*** Found 1 match."
  TU vsock, EmptyLine
End Sub

Public Sub ListScripts(vsock As Long) ' : AddStack "GUI_ListScripts(" & vsock & ")"
Dim FileName As String, Description As String, u As Long, u2 As Long
Dim Rest As String, MissedOne As Boolean, FoundOne As Boolean
  If ScriptCount > 0 Then
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, "2*** Listing running scripts:"
      TU vsock, "0,1 File name:             | Description:                          "
    Else
      TU vsock, "*** Listing running scripts:"
      TU vsock, " File name:               Description:"
      TU vsock, " -----------------------  --------------------------------------"
    End If
  End If
  For u = 1 To ScriptCount
    FileName = Scripts(u).Name
    Description = Scripts(u).Description
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, " " & Spaces2(22, FileName) & "14 | " & Description
    Else
      TU vsock, " " & Spaces2(22, FileName) & "   " & Description
    End If
  Next u
  Rest = GetPPString("Scripts", "Load", "", AnGeL_INI)
  For u = 1 To ParamCount(Rest)
    FileName = Param(Rest, u)
    FoundOne = False
    For u2 = 1 To ScriptCount
      If LCase(FileName) = LCase(Scripts(u2).Name) Then FoundOne = True: Exit For
    Next u2
    If Not FoundOne Then
      If Not MissedOne Then
        MissedOne = True
        If ScriptCount = 0 Then
          TU vsock, "2*** The following scripts are NOT running:"
          If GetSockFlag(vsock, SF_Colors) = SF_YES Then
            TU vsock, "0,1 File name:             | Description:                          "
          Else
            TU vsock, " File name:               Description:"
            TU vsock, " -----------------------  --------------------------------------"
          End If
        End If
      End If
      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
        TU vsock, " " & Spaces2(22, FileName) & "14 | 14<not running due to errors>"
      Else
        TU vsock, " " & Spaces2(22, FileName) & "   <not running due to errors>"
      End If
    End If
  Next u
  If (ScriptCount = 0) And (Not MissedOne) Then
    TU vsock, "5*** There are no scripts. You can add one by sending it"
    TU vsock, "5    to the bot via DCC and using '.+script'."
    Exit Sub
  End If
  TU vsock, EmptyLine
End Sub

Public Sub Whois(vsock As Long, Line As String, ShortWhois As Boolean) ' : AddStack "GUI_Whois(" & vsock & ", " & Line & ", " & ShortWhois & ")"
Dim SeeChannel As String, ChFlags As String, HisFlags As String, HisBotFlags As String
Dim UserState As String, Setting As String, InRightSection As Boolean
Dim UsNum As Long, u As Long, ScNum As Long
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".whois <nick>"): Exit Sub
  UsNum = GetUserNum(Param(Line, 2))
  If UsNum = 0 Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
  HisFlags = BotUsers(UsNum).Flags
  HisBotFlags = BotUsers(UsNum).BotFlags
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    TU vsock, "2*** WHOIS report for " & BotUsers(UsNum).Name & ""
    TU vsock, "0,1 Channel:               | Flags:                  "
    TU vsock, " global (all channels) 14 | " & HisFlags
    If HisBotFlags <> "" Then TU vsock, " bot flags             14 | +" & HisBotFlags
  Else
    TU vsock, "*** WHOIS report for " & BotUsers(UsNum).Name
    TU vsock, " Channel:                 Flags:                  "
    TU vsock, " -----------------------  ------------------------"
    TU vsock, " global (all channels)    " & HisFlags
    If HisBotFlags <> "" Then TU vsock, " bot flags                +" & HisBotFlags
  End If
  For u = 1 To BotUsers(UsNum).ChannelFlagCount
    SeeChannel = BotUsers(UsNum).ChannelFlags(u).Channel
    ChFlags = BotUsers(UsNum).ChannelFlags(u).Flags
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then TU vsock, " " & SeeChannel + Spaces(22, SeeChannel) & "14 | +" & ChFlags
    If GetSockFlag(vsock, SF_Colors) = SF_NO Then TU vsock, " " & SeeChannel + Spaces(22, SeeChannel) & "   +" & ChFlags
  Next u
  UserState = "3User"
  If MatchFlags(HisFlags, "+t") Then UserState = "10Botnet Master"
  If MatchFlags(HisFlags, "+m") Then UserState = "12Master"
  If MatchFlags(HisFlags, "+n") Then UserState = "4Owner"
  If MatchFlags(HisFlags, "+s") Then UserState = "4Super Owner14 (unchangeable flags)"
  If MatchFlags(HisFlags, "+b") And MatchFlags(HisBotFlags, "-h") Then UserState = "14BOT"
  If MatchFlags(HisFlags, "+b") And MatchFlags(HisBotFlags, "+h") Then UserState = "14HUB BOT, auto-linked"
  TU vsock, EmptyLine
  TU vsock, " Status   : " & UserState
  If MatchFlags(SocketItem(vsock).Flags, "+t") Then
    If (GetUserData(UsNum, UD_LinkAddr, "") <> "") And MatchFlags(HisFlags, "+b") Then TU vsock, " Address  : " & GetUserData(UsNum, UD_LinkAddr, "")
  End If
  'Only display "colors" setting for non-bots
  If MatchFlags(HisFlags, "-b") Then UserState = "Colors " & IIf(GetUserData(UsNum, "colors", SF_NO) = SF_YES, "[3ON]", "[14OFF]") & " - " Else UserState = ""
  TU vsock, " Settings : " & UserState & "Password " & IIf(BotUsers(UsNum).Password <> "", "[3YES]", "[14NO]")
  If GetUserData(UsNum, "comment", "") <> "" Then
    TU vsock, " Comment  : " & GetUserData(UsNum, "comment", "")
  End If
  If GetUserData(UsNum, "info", "") <> "" Then
    TU vsock, " Info     : " & GetUserData(UsNum, "info", "")
  End If
  'TU vsock, " UserData : " & BotUsers(UsNum).UserData
  'Check script Hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.Whois Then
      RunScriptX ScNum, "Whois", vsock, SocketItem(vsock).RegNick, SocketItem(vsock).Flags, Line
    End If
  Next ScNum
  For u = 1 To BotUsers(UsNum).HostMaskCount
    If u = 1 Then TU vsock, " Hostmasks: " & BotUsers(UsNum).HostMasks(u) Else TU vsock, "            " & BotUsers(UsNum).HostMasks(u)
  Next u
  TU vsock, EmptyLine
  If ShortWhois = False Then
    TUEx vsock, SF_ExtraHelp, "14 Type '10.help whois14' to get a list of all flags."
    TUEx vsock, SF_ExtraHelp, EmptyLine
  End If
End Sub

Public Sub Kick(vsock As Long, Line As String) ' : AddStack "GUI_Kick(" & vsock & ", " & Line & ")"
Dim SockNum As Long, u As Long, FoundOne As Boolean, UsNum As Long
Dim ChanFlags As String, OtherUserFlags As String, Messg As String
Dim KickedOne As Boolean, ChNum As Long
  SockNum = SocketItem(vsock).SocketNumber
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".kick <nick> ([" & ServerChannelPrefixes & "]channel) (reason)"): Exit Sub
  If LCase(Param(Line, 2)) = LCase(MyNick) Then TU vsock, "5*** Ha ha.": Exit Sub
  If Not (IsValidChannel(LCase(Left(Param(Line, 3), 1)))) Then
    u = Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " ")
    If u > 0 Then Messg = Right(Line, u) Else Messg = "Requested"
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags2(SocketItem(vsock).UserNum, Channels(u).Name)
        OtherUserFlags = GetUserChanFlags2(Channels(u).User(UsNum).UserNum, Channels(u).Name)
        If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+b") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is a bot.": Exit Sub
        If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+n") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is an owner.": Exit Sub
        If MatchFlags(ChanFlags, "-m") And MatchFlags(OtherUserFlags, "+m") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is a master.": Exit Sub
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Or (Channels(u).GotHOPs And InStr(1, Channels(u).User(UsNum).Status, "@", vbBinaryCompare) = 0) Then
            SendLine "kick " & Channels(u).Name & " " & Channels(u).User(UsNum).Nick & " :" & Messg, 1
            TU vsock, "3*** Kicked " & Channels(u).User(UsNum).Nick & " from " & Channels(u).Name & "."
            KickedOne = True
          Else
            TU vsock, "5*** Couldn't kick " & Channels(u).User(UsNum).Nick & " from " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not KickedOne Then TU vsock, "5*** I didn't kick anybody.": Exit Sub
  Else
    u = Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3) & " ")
    If u > 0 Then Messg = Right(Line, u) Else Messg = "Requested"
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags2(SocketItem(vsock).UserNum, Channels(ChNum).Name)
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have kick rights for " & Channels(ChNum).Name & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    OtherUserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Channels(ChNum).Name)
    If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+b") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a bot.": Exit Sub
    If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+n") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is an owner.": Exit Sub
    If MatchFlags(ChanFlags, "-m") And MatchFlags(OtherUserFlags, "+m") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a master.": Exit Sub
    If Channels(ChNum).GotOPs Or (Channels(ChNum).GotHOPs And InStr(1, Channels(ChNum).User(UsNum).Status, "@", vbBinaryCompare) = 0) Then SendLine "kick " & Param(Line, 3) & " " & Channels(ChNum).User(UsNum).Nick & " :" & Messg, 1
    If Not (Channels(ChNum).GotOPs Or (Channels(ChNum).GotHOPs And InStr(1, Channels(ChNum).User(UsNum).Status, "@", vbBinaryCompare) = 0)) Then TU vsock, "5*** Couldn't kick " & Channels(ChNum).User(UsNum).Nick & " from " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Kicked " & Channels(ChNum).User(UsNum).Nick & " from " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 2
End Sub

Public Sub KickBan(vsock As Long, Line As String) ' : AddStack "GUI_KickBan(" & vsock & ", " & Line & ")"
Dim SockNum As Long, u As Long, FoundOne As Boolean, UsNum As Long
Dim ChanFlags As String, OtherUserFlags As String, Messg As String
Dim KickedOne As Boolean, ChNum As Long, u2 As Long
  SockNum = SocketItem(vsock).SocketNumber
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".kickban <nick/mask> ([" & ServerChannelPrefixes & "]channel) (reason)"): Exit Sub
  If LCase(Param(Line, 2)) = LCase(MyNick) Then TU vsock, "5*** Are you crazy?!": Exit Sub
  If Not (IsValidChannel(LCase(Left(Param(Line, 3), 1)))) Then
    u = Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " ")
    If u > 0 Then Messg = Right(Line, u) Else Messg = "Requested"
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags2(SocketItem(vsock).UserNum, Channels(u).Name)
        OtherUserFlags = GetUserChanFlags2(Channels(u).User(UsNum).UserNum, Channels(u).Name)
        If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+b") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is a bot.": Exit Sub
        If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+n") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is an owner.": Exit Sub
        If MatchFlags(ChanFlags, "-m") And MatchFlags(OtherUserFlags, "+m") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is a master.": Exit Sub
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Or Channels(u).GotHOPs Then
            PatternKickBan Channels(u).User(UsNum).Nick, Channels(u).Name, Messg, False, 120, 300
            TU vsock, "3*** Kickbanned " & Channels(u).User(UsNum).Nick & " from " & Channels(u).Name & "."
            KickedOne = True
          Else
            TU vsock, "5*** Couldn't kick " & Channels(u).User(UsNum).Nick & " from " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not KickedOne Then TU vsock, "5*** I didn't kick anybody.": Exit Sub
  Else
    u = Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3) & " ")
    If u > 0 Then Messg = Right(Line, u) Else Messg = "Requested"
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags2(SocketItem(vsock).UserNum, Channels(ChNum).Name)
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have kick rights for " & Channels(ChNum).Name & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    OtherUserFlags = GetUserChanFlags2(Channels(ChNum).User(UsNum).UserNum, Channels(ChNum).Name)
    If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+b") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a bot.": Exit Sub
    If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+n") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is an owner.": Exit Sub
    If MatchFlags(ChanFlags, "-m") And MatchFlags(OtherUserFlags, "+m") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a master.": Exit Sub
    If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then PatternKickBan Channels(ChNum).User(UsNum).Nick, Channels(ChNum).Name, Messg, False, 120, 300
    If Not (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) Then TU vsock, "5*** Couldn't kickban " & Channels(ChNum).User(UsNum).Nick & " from " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Kickbanned " & Channels(ChNum).User(UsNum).Nick & " from " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 2
End Sub

Public Sub GUIJump(vsock As Long, Line As String) ' : AddStack "GUI_GUIJump(" & vsock & ", " & Line & ")"
  Dim u As Long, Rest As String, TmpStr As String
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".jump <server number> or <server(:port)> (|proxy(:port))"): Exit Sub
  If IsNumeric(GetRest(Line, 2)) Then
    u = CLng(GetRest(Line, 2))
    Rest = GetPPString("Server", IIf(u = 1, "Server", "Server" & CStr(u)), "", AnGeL_INI)
    If Rest = "" Then TU vsock, "5*** I couldn't find this server. Type '.servers' to see my list.": Exit Sub
  Else
    Rest = GetRest(Line, 2)
  End If
  Status "*** Server jump requested..." & vbCrLf
  SpreadMessage 0, -1, "7*** SERVER JUMP requested by " & SocketItem(vsock).RegNick & ""
  DontConnect = False
  SendLine "quit :Changing servers...", 1
  Disconnect
  Output vbCrLf
  Output "*** Socket Closed (me)" & vbCrLf
  RemoveTimedEvent "ConnectServer"
  ConnectServer JumpConnectDelay, Rest
End Sub


Public Sub GUIVoice(vsock As Long, Line As String)
  Dim u As Long, UsNum As Long, ChanFlags As String, FoundOne As Boolean, oppedone As Boolean, ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".voice <nick> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If Param(Line, 3) = "" Then
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags(Nick, Channels(u).Name)
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Or Channels(u).GotHOPs Then
            SendLine "mode " & Channels(u).Name & " +v " & Channels(u).User(UsNum).Nick, 1: oppedone = True
            TU vsock, "3*** Gave voice to " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & "."
          Else
            TU vsock, "5*** Couldn't voice " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not oppedone Then TU vsock, "5*** You don't have voice rights for the channels this user is on.": Exit Sub
  Else
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags(Nick, Param(Line, 3))
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have voice rights for " & Param(Line, 3) & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then SendLine "mode " & Param(Line, 3) & " +v " & Channels(ChNum).User(UsNum).Nick, 1
    If Not (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) Then TU vsock, "5*** Couldn't voice " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Gave voice to " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
End Sub

Public Sub GUIDeVoice(vsock As Long, Line As String)
  Dim u As Long, UsNum As Long, ChanFlags As String, FoundOne As Boolean, oppedone As Boolean, ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".devoice <nick> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If Param(Line, 3) = "" Then
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags(Nick, Channels(u).Name)
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Or Channels(u).GotHOPs Then
            SendLine "mode " & Channels(u).Name & " -v " & Channels(u).User(UsNum).Nick, 1: oppedone = True
            TU vsock, "3*** Took voice from " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & "."
          Else
            TU vsock, "5*** Couldn't devoice " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not oppedone Then TU vsock, "5*** You don't have devoice rights for the channels this user is on.": Exit Sub
  Else
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags(Nick, Param(Line, 3))
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have devoice rights for " & Param(Line, 3) & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then SendLine "mode " & Param(Line, 3) & " -v " & Channels(ChNum).User(UsNum).Nick, 1
    If Not (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) Then TU vsock, "5*** Couldn't devoice " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Took voice from " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
End Sub

Public Sub GUIHop(vsock As Long, Line As String)
  Dim u As Long, UsNum As Long, ChanFlags As String, FoundOne As Boolean, oppedone As Boolean, ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".hop <nick> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If Param(Line, 3) = "" Then
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags(Nick, Channels(u).Name)
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Then
            SendLine "mode " & Channels(u).Name & " +h " & Channels(u).User(UsNum).Nick, 1: oppedone = True
            TU vsock, "3*** Gave halfop to " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & "."
          Else
            TU vsock, "5*** Couldn't halfop " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not oppedone Then TU vsock, "5*** You don't have halfop rights for the channels this user is on.": Exit Sub
  Else
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags(Nick, Param(Line, 3))
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have halfop rights for " & Param(Line, 3) & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    If Channels(ChNum).GotOPs Then SendLine "mode " & Param(Line, 3) & " +h " & Channels(ChNum).User(UsNum).Nick, 1
    If Not Channels(ChNum).GotOPs Then TU vsock, "5*** Couldn't halfop " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Gave halfop to " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
End Sub

Public Sub GUIDeHop(vsock As Long, Line As String)
  Dim u As Long, UsNum As Long, ChanFlags As String, FoundOne As Boolean, oppedone As Boolean, ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".devoice <nick> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If Param(Line, 3) = "" Then
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags(Nick, Channels(u).Name)
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Then
            SendLine "mode " & Channels(u).Name & " -h " & Channels(u).User(UsNum).Nick, 1: oppedone = True
            TU vsock, "3*** Took halfop from " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & "."
          Else
            TU vsock, "5*** Couldn't dehalfop " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not oppedone Then TU vsock, "5*** You don't have dehalfop rights for the channels this user is on.": Exit Sub
  Else
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags(Nick, Param(Line, 3))
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have dehalfop rights for " & Param(Line, 3) & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    If Channels(ChNum).GotOPs Then SendLine "mode " & Param(Line, 3) & " -h " & Channels(ChNum).User(UsNum).Nick, 1
    If Not Channels(ChNum).GotOPs Then TU vsock, "5*** Couldn't dehalfop " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Took halfop from " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
End Sub

Public Sub GUIOp(vsock As Long, Line As String)
  Dim u As Long, UsNum As Long, ChanFlags As String, FoundOne As Boolean, oppedone As Boolean, ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".op <nick> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If Param(Line, 3) = "" Then
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags(Nick, Channels(u).Name)
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Then
            GiveOp Channels(u).Name, Channels(u).User(UsNum).Nick: oppedone = True
            TU vsock, "3*** Gave op to " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & "."
          Else
            TU vsock, "5*** Couldn't op " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not oppedone Then TU vsock, "5*** You don't have op rights for the channels this user is on.": Exit Sub
  Else
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags(Nick, Param(Line, 3))
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have op rights for " & Param(Line, 3) & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    If Channels(ChNum).GotOPs Then GiveOp Param(Line, 3), Channels(ChNum).User(UsNum).Nick
    If Not Channels(ChNum).GotOPs Then TU vsock, "5*** Couldn't op " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Gave op to " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 2
End Sub

Public Sub GUIDeOp(vsock As Long, Line As String)
  Dim u As Long, UsNum As Long, OtherUserFlags As String, ChanFlags As String, FoundOne As Boolean, oppedone As Boolean, ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".deop <nick> ([" & ServerChannelPrefixes & "]channel)"): Exit Sub
  If LCase(Param(Line, 2)) = LCase(MyNick) Then TU vsock, "5*** Ha ha.": Exit Sub
  If Param(Line, 3) = "" Then
    For u = 1 To ChanCount
      UsNum = FindUser(Param(Line, 2), u)
      If UsNum > 0 Then
        ChanFlags = GetUserChanFlags(Nick, Channels(u).Name)
        OtherUserFlags = GetUserChanFlags(SearchUserFromHostmask(Channels(u).User(UsNum).Hostmask), Channels(u).Name)
        If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+b") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is a bot.": Exit Sub
        If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+n") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is an owner.": Exit Sub
        If MatchFlags(ChanFlags, "-m") And MatchFlags(OtherUserFlags, "+m") Then TU vsock, "5*** Sorry, " & Channels(u).User(UsNum).Nick & " is a master.": Exit Sub
        FoundOne = True
        If MatchFlags(ChanFlags, "+o") Then
          If Channels(u).GotOPs Then
            SendLine "mode " & Channels(u).Name & " -o " & Channels(u).User(UsNum).Nick, 1: oppedone = True
            TU vsock, "3*** Took op from " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & "."
          Else
            TU vsock, "5*** Couldn't deop " & Channels(u).User(UsNum).Nick & " on " & Channels(u).Name & " - I don't have OPs."
          End If
        End If
      End If
    Next u
    If Not FoundOne Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Param(Line, 2)): Exit Sub
    If FoundOne And Not oppedone Then TU vsock, "5*** You don't have deop rights for the channels this user is on.": Exit Sub
  Else
    ChNum = FindChan(Param(Line, 3))
    If ChNum = 0 Then TU vsock, "5*** I'm not in " & Param(Line, 3) & ".": Exit Sub
    ChanFlags = GetUserChanFlags(Nick, Param(Line, 3))
    If MatchFlags(ChanFlags, "-o") Then TU vsock, "5*** You don't have deop rights for " & Param(Line, 3) & ".": Exit Sub
    UsNum = FindUser(Param(Line, 2), ChNum)
    If UsNum = 0 Then TU vsock, "5*** I couldn't find this user on " & Param(Line, 3) & ".": Exit Sub
    OtherUserFlags = GetUserChanFlags(SearchUserFromHostmask(Channels(ChNum).User(UsNum).Hostmask), Channels(ChNum).Name)
    If MatchFlags(ChanFlags, "-n") And MatchFlags(OtherUserFlags, "+n") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is an owner.": Exit Sub
    If MatchFlags(ChanFlags, "-m") And MatchFlags(OtherUserFlags, "+m") Then TU vsock, "5*** Sorry, " & Channels(ChNum).User(UsNum).Nick & " is a master.": Exit Sub
    If Channels(ChNum).GotOPs Then SendLine "mode " & Param(Line, 3) & " -o " & Channels(ChNum).User(UsNum).Nick, 1
    If Not Channels(ChNum).GotOPs Then TU vsock, "5*** Couldn't deop " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & " - I don't have OPs.": Exit Sub
    TU vsock, "3*** Took op from " & Channels(ChNum).User(UsNum).Nick & " on " & Param(Line, 3) & "."
  End If
  SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 2
End Sub

Sub GUIKey(vsock As Long, Line As String)
  Dim ChNum As Long, Nick As String
  Nick = SocketItem(vsock).RegNick
  If Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".key <[" & ServerChannelPrefixes & "]channel>"): Exit Sub
  ChNum = FindChan(Param(Line, 2))
  If ChNum = 0 Then TU vsock, "5*** Sorry, I'm not in this channel.": Exit Sub
  If MatchFlags(GetUserChanFlags(Nick, Param(Line, 2)), "-o") Then TU vsock, "5*** Sorry, you need the +o flag to get the key for this channel.": FailCommand vsock, "+m", Line: Exit Sub
  SucceedCommand vsock, "+m", Line
  If GetChannelKey(ChNum) <> "" Then
    TU vsock, "10*** Key for " & Channels(ChNum).Name & ":3 " & GetChannelKey(ChNum) & ""
  Else
    TU vsock, "10*** Key for " & Channels(ChNum).Name & ":14 <none>"
  End If
End Sub

Sub GUITrace(vsock As Long, Line As String)
  Dim BotNum As Long
  Dim CurCheckBot As String
  Dim TargetBot As String
  Dim FoundOne As Boolean
  Dim u As Long
  'td <ttl>:<who>@<source> :<timestamp>:<bot1>:<bot2>...
  
  If Param(Line, 2) = "" Then
    TU vsock, "05*** Usage: .trace <bot>"
    Exit Sub
  Else
    If LCase(Param(Line, 2)) = LCase(BotNetNick) Then
      TU vsock, "05*** Trace: Yeah thats me. Reply in 1ms :)"
      Exit Sub
    End If
    CurCheckBot = Param(Line, 2)
    TargetBot = ""
    Do
      FoundOne = False
      For u = 1 To BotCount
        If LCase(Bots(u).Nick) = LCase(CurCheckBot) Then
          FoundOne = True
          If TargetBot = "" Then TargetBot = Bots(u).Nick
          If LCase(Bots(u).SubBotOf) = LCase(BotNetNick) Then
            CurCheckBot = Bots(u).Nick
            Exit Do
          End If
          CurCheckBot = Bots(u).SubBotOf
          Exit For
        End If
      Next u
      If Not FoundOne Then Exit Do
    Loop
    If FoundOne = False Then
      TU vsock, "05*** Sorry can´t find that bot."
      Exit Sub
    End If
    For u = 1 To SocketCount
      If SocketItem(u).RegNick = CurCheckBot Then
        RTU u, "t 9:" & SocketItem(vsock).RegNick & "@" & BotNetNick & " " & TargetBot & " :" & WinTickCount & ":" & BotNetNick
        Exit Sub
      End If
    Next u
  End If
End Sub

Sub GUIRelay(vsock As Long, Line As String)
  Dim u As Long
  Dim u2 As Long
  Dim Port As String
  Dim Host As String
  Dim NewSock As Long
  Dim strRemoteIP As String
  
  If Param(Line, 2) = "" Then TU vsock, "05*** Usage: .relay <bot>" & IIf(MatchFlags(SocketItem(vsock).Flags, "+t"), " (host:addr)", ""): Exit Sub
  If Param(Line, 3) <> "" And MatchFlags(SocketItem(vsock).Flags, "+t") = False Then TU vsock, "05*** Usage: .relay <bot>": Exit Sub
  u2 = -1
  For u = 1 To BotUserCount
    If LCase(Param(Line, 2)) = LCase(BotUsers(u).Name) And MatchFlags(BotUsers(u).Flags, "+b") Then u2 = u: Exit For
  Next u
  If u2 = -1 Then
    TU vsock, "05*** Sorry can´t find that bot."
  Else
    Port = GetUserData(u2, "UserPort", "<NOT SET>")
    Host = GetUserData(u2, "addr", "<NOT SET>")
    If Not Host = "<NOT SET>" Then
      If Port = "<NOT SET>" Then
        Port = Mid(Host, InStr(Host, ":") + 1)
        SetUserData u2, "UserPort", Port
      End If
      Host = Mid(Host, 1, InStr(Host, ":") - 1)
    Else
      Port = "<NOT SET>"
      TU vsock, "5*** No ConnectAddress set..."
      TU vsock, "5*** Use '.chaddr " & Param(Line, 2) & " Host:Port' or '.relay " & Param(Line, 2) & " <HOST:PORT>"
    End If
    If Param(Line, 3) <> "" Then
      If InStr(Param(Line, 3), ":") > 0 Then
        Port = Mid(Param(Line, 3), InStr(Param(Line, 3), ":") + 1)
        If Host = "<NOT SET>" Then
          SetUserData u2, "addr", Param(Line, 3)
        End If
        Host = Mid(Param(Line, 3), 1, InStr(Param(Line, 3), ":") - 1)
        SetUserData u2, "UserPort", Port
      Else
        Port = "<NOT SET>"
        TU vsock, "5*×* Usage: .relay <bot> <HOST:PORT>"
        Exit Sub
      End If
    End If
    
    If Port = "<NOT SET>" Then Exit Sub
    NewSock = AddSocket
    If ConnectTCP(NewSock, Host, CLng(Port)) = 0 Then
      TU vsock, "3*** Connecting to " & Host & ":" & Port & " ..."
      SetSockFlag NewSock, SF_Status, SF_Status_RelayCli
      SocketItem(NewSock).CurrentQuestion = vsock
      SocketItem(NewSock).SetupChan = BotUsers(u2).Name
      SocketItem(NewSock).UserNum = 0
      SocketItem(NewSock).OnBot = ""
      SocketItem(NewSock).IRCNick = ""
      SocketItem(NewSock).RegNick = ""
      SocketItem(NewSock).IsInternalSocket = True
      SetSockFlag vsock, SF_Status, SF_Status_RelaySrv
      SetSockFlag vsock, SF_LocalVisibleUser, SF_NO
      SocketItem(vsock).CurrentQuestion = NewSock
    Else
      TU vsock, "5*** Unable to connect..."
      RemoveSocket NewSock, 0, "", True
    End If
  End If
End Sub

Public Sub Netsetup(vsock As Long, Line As String) ' : AddStack "Setups_NetSetup(" & vsock & ", " & Line & ")"
Dim u As Long, SockFlag As Byte, InstantSet As Boolean
InstantSetIt:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving Network info..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
          Case "1"
              SocketItem(vsock).CurrentQuestion = "NetName"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Network name"
              TU vsock, "  14 Here's an example for this setting: 11QuakeNET"
              TU vsock, "  14 Specifies the Network name of the IRC-Network"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "NetHandle"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "22) Nicklength"
              TU vsock, "  14 Specifies the Nicklength of the IRC-Network"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "NetChan"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "23) Max. Channels"
              TU vsock, "  14 Specifies the count of Channels which can be joined"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "NetPrivMsg"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "24) LONGMESSAGE"
              TU vsock, "  14 Specifies whether the Network supports 'PRIVMSG NICK!IDENT@HOST.COM'."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "NetSERVERS"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "25) &SERVERS"
              TU vsock, "  14 Specifies whether the network supports '&SERVERS' (Split detection)."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "6"
              SocketItem(vsock).CurrentQuestion = "NetPrefix"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "26) ChanPrefix"
              TU vsock, "  14 Specifies the prefix of channels which can be joined"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "7"
              SocketItem(vsock).CurrentQuestion = "NetUserMode"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "27) User Prefixes"
              TU vsock, "  14 Specifies the mode and prefix for users which can be set"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "8"
              SocketItem(vsock).CurrentQuestion = "NetChanMode"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "26) Channel Modes"
              TU vsock, "  14 Specifies the modes of channels which can be set"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in NetSetup!"
              TU vsock, "5*** Please enter a valid number."
              Exit Sub
        End Select
    Case "NetName"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              ServerNetwork = Line
              WritePPString "NET", "NetworkName", ServerNetwork, NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "NetHandle"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If IsNumeric(Param(Line, 1)) Then
                ServerNickLen = Param(Line, 1)
                WritePPString "NET", "NickLength", ServerNickLen, NET_INI
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              Else
                TU vsock, "5*** Please enter a numeric value."
              End If
        End Select
    Case "NetChan"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              ServerMaxChannels = Param(Line, 1)
              WritePPString "NET", "MaxChan", ServerMaxChannels, NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "NetPrivMsg"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on", "off", "yes", "no", "an", "aus", "ja", "nein"
              ServerUseFullAdress = Switch(Line)
              WritePPString "NET", "UseFullAdress", IIf(ServerUseFullAdress, "1", "0"), NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              TU vsock, "5*** Please enter 'yes', 'no' or '0'."
        End Select
    Case "NetSERVERS"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on", "off", "yes", "no", "an", "aus", "ja", "nein"
              ServerSplitDetection = Switch(Line)
              WritePPString "NET", "SplitDetection", IIf(ServerSplitDetection, "1", "0"), NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              TU vsock, "5*** Please enter 'yes', 'no' or '0'."
        End Select
    Case "NetPrefix"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              ServerChannelPrefixes = Param(Line, 1)
              WritePPString "NET", "ChanPrefixes", ServerChannelPrefixes, NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "NetUserMode"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              ServerUserModes = Param(Line, 1)
              WritePPString "NET", "UserPrefixes", ServerUserModes, NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "NetChanMode"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              ServerChannelModes = Param(Line, 1)
              WritePPString "NET", "ChanModes", ServerChannelModes, NET_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case Else
        TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_NETSetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowNETActions vsock
  End If
End Sub

Public Sub AUTHsetup(vsock As Long, Line As String) ' : AddStack "Setups_NetSetup(" & vsock & ", " & Line & ")"
Dim u As Long, SockFlag As Byte, InstantSet As Boolean
InstantSetIt:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving AUTH info..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
          Case "1"
              SocketItem(vsock).CurrentQuestion = "AuthName"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Service name"
              TU vsock, "  14 Here's an example for this setting: 11NickServ"
              TU vsock, "  14 Specifies the name of the AUTH Service."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' to disable or '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "AuthCommand"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "22) Auth command"
              TU vsock, "  14 Specifies the command to auth."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "AuthUser"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "23) Parameter 1"
              TU vsock, "  14 Specifies the first parameter to auth."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "AuthPass"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "24) Parameter 2"
              TU vsock, "  14 Specifies the second parameter to auth."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' for none or '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "AuthRE"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "25) reAUTH"
              TU vsock, "  14 Specifies whether reAUTH if Service joins a channel."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in NetSetup!"
              TU vsock, "5*** Please enter a valid number."
              Exit Sub
        End Select
    Case "AuthName"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x":
              AuthTarget = ""
              DeletePPString "AUTH", "Target", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              AuthTarget = Param(Line, 1)
              WritePPString "AUTH", "Target", Param(Line, 1), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "AuthCommand"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              AuthCommand = Param(Line, 1)
              WritePPString "AUTH", "Command", Param(Line, 1), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "AuthUser"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x":
              AuthParam1 = ""
              DeletePPString "AUTH", "Username", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              AuthParam1 = Param(Line, 1)
              WritePPString "AUTH", "Username", Param(Line, 1), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "AuthPass"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x":
              AuthParam2 = ""
              DeletePPString "AUTH", "Password", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              AuthParam2 = Param(Line, 1)
              WritePPString "AUTH", "Password", Param(Line, 1), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "AuthRE"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on", "off", "yes", "no", "an", "aus", "ja", "nein"
              AuthReAuth = Switch(Line)
              WritePPString "AUTH", "ReAuth", IIf(AuthReAuth, "1", "0"), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              TU vsock, "5*** Please enter 'yes', 'no' or '0'."
        End Select
    Case Else
        TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_AUTHSetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowAUTHActions vsock
  End If
End Sub

' Personal Setup
'-——————————————————————————————————————————————————————- -- -  -
Public Sub PersonalSetup(vsock As Long, Line As String) ' : AddStack "Setups_PersonalSetup(" & vsock & ", " & Line & ")"
Dim u As Long, SockFlag As Byte, InstantSet As Boolean
InstantSetIt3:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving personal info..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
          Case "1"
              SocketItem(vsock).CurrentQuestion = "ExtraHelp"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "21) Additional help messages"
              TU vsock, "  14 If this setting is turned on, I'll show you things"
              TU vsock, "  14 like '10Type .help whois to get a list of all flags14'"
              TU vsock, "  14 to tell you how to find help on special topics. If"
              TU vsock, "  14 this is turned off, I'll not bother you with these"
              TU vsock, "  14 additional messages (good for advanced users)."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "Local_JP"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "22) Local user joins/parts"
              TU vsock, "  14 If turned on, I'll show a message when somebody"
              TU vsock, "  14 joins or leaves my local party line. Example:"
              TU vsock, EmptyLine
              TU vsock, "   " & MakeMsg(MSG_PLJoin, "Hippo")
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "Local_Talk"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "23) Local party line talk"
              TU vsock, "  14 If turned on, I'll show a message when somebody"
              TU vsock, "  14 says something on my local party line. Example:"
              TU vsock, EmptyLine
              TU vsock, "   " & MakeMsg(MSG_PLTalk, "Hippo", "Hi Kane!")
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "Local_Bot"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "24) Local bot connections"
              TU vsock, "  14 If turned on, I'll show a message when a bot"
              TU vsock, "  14 establishes / loses a link to me. Example:"
              TU vsock, EmptyLine
              TU vsock, "  3 *** Connected to |^AnGeL^|"
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "Botnet_JP"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "25) Botnet user joins/parts"
              TU vsock, "  14 If turned on, I'll show a message when somebody"
              TU vsock, "  14 joins or leaves the partyline of a bot in the"
              TU vsock, "  14 botnet. Example:"
              TU vsock, EmptyLine
              TU vsock, "   " & MakeMsg(MSG_PLBotNetJoin, "LameBot", "Franklin")
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "6"
              SocketItem(vsock).CurrentQuestion = "Botnet_Talk"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "26) Botnet party line talk"
              TU vsock, "  14 If turned on, I'll show a message when somebody"
              TU vsock, "  14 says something in the botnet. Example:"
              TU vsock, EmptyLine
              TU vsock, "   " & MakeMsg(MSG_PLBotNetTalk, "sChOBoT", "Kane", "moin Hippo :)")
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "7"
              SocketItem(vsock).CurrentQuestion = "Botnet_Bot"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "27) Botnet bot connections"
              TU vsock, "  14 If turned on, I'll show a message when a bot"
              TU vsock, "  14 establishes / loses a link to another bot in"
              TU vsock, "  14 the botnet. Example:"
              TU vsock, EmptyLine
              TU vsock, "   " & MakeMsg(MSG_PLBotTalk, "LameBot", MakeMsg(MSG_BNPingTimeout) & ": Blub (Lost 4 bots & 1 user)")
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "8"
              SocketItem(vsock).CurrentQuestion = "BotMSG_DefChan"
              If GetRest(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "28) Default chan for commands"
              TU vsock, EmptyLine
              TU vsock, "Type a channel, 'x' for none or '0' to cancel."
          Case "9"
              SocketItem(vsock).CurrentQuestion = "BotMSG_ChanJP"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "29) Joins/parts/Quits in channels"
              TU vsock, EmptyLine
              TU vsock, "Type a channel, 'x' for none or '0' to cancel."
          Case "10"
              SocketItem(vsock).CurrentQuestion = "BotMSG_ChanKB"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "210) Kicks/Modes in channels"
              TU vsock, EmptyLine
              TU vsock, "Type a channel, 'x' for none or '0' to cancel."
          Case "11"
              SocketItem(vsock).CurrentQuestion = "BotMSG_ChanPT"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "211) Public talk in channels"
              TU vsock, EmptyLine
              TU vsock, "Type a channel, 'x' for none or '0' to cancel."
          Case "12"
              SocketItem(vsock).CurrentQuestion = "BotMSG_FloodMSG"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "212) Flood protection messages"
              TU vsock, EmptyLine
              TU vsock, "Type a channel, 'x' for none or '0' to cancel."
          Case "13"
              SocketItem(vsock).CurrentQuestion = "BotMSG_ShowStatus"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "213) Status for default channel"
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "14"
              SocketItem(vsock).CurrentQuestion = "PrivToBot"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt3
              TU vsock, "214) PRIVMSGs to the bot"
              TU vsock, "   14 If turned on, I'll show you when somebody sends"
              TU vsock, "   14 a private message (via '/msg' or '/query') to"
              TU vsock, "   14 me. If the bot setup allows automatic answers,"
              TU vsock, "   14 you'll also see my reply to the user. Example:"
              TU vsock, EmptyLine
              TU vsock, "    " & "14[" & Time & "] PRIVMSG from Hippo: Hi, wie gehts?"
              TU vsock, "    " & "14[" & Time & "] I replied to Hippo: Hoi! Gut, danke... und dir?"
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in BotSetup!"
              TU vsock, "5*** Please enter a valid number."
              Exit Sub
        End Select
    Case "ExtraHelp", "Local_JP", "Local_Talk", "Local_Bot", "Botnet_JP", "Botnet_Talk", "Botnet_Bot", "PrivToBot"
        Select Case SocketItem(vsock).CurrentQuestion
          Case "ExtraHelp": SockFlag = SF_ExtraHelp
          Case "Local_JP": SockFlag = SF_Local_JP
          Case "Local_Talk": SockFlag = SF_Local_Talk
          Case "Local_Bot": SockFlag = SF_Local_Bot
          Case "Botnet_JP": SockFlag = SF_Botnet_JP
          Case "Botnet_Talk": SockFlag = SF_Botnet_Talk
          Case "Botnet_Bot": SockFlag = SF_Botnet_Bot
          Case "PrivToBot": SockFlag = SF_PrivToBot
        End Select
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              SetSockFlag vsock, SockFlag, SF_YES
              SaveSockFlags vsock
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              SetSockFlag vsock, SockFlag, SF_NO
              SaveSockFlags vsock
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "BotMSG_DefChan"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
            SetUserData SocketItem(vsock).UserNum, "BMDF", ""
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
            SetUserData SocketItem(vsock).UserNum, "BMDF", (GetRest(Line, 1))
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotMSG_ChanJP"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
            SetUserData SocketItem(vsock).UserNum, "BMJP", ""
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
            SetUserData SocketItem(vsock).UserNum, "BMJP", (GetRest(Line, 1))
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotMSG_ChanKB"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
            SetUserData SocketItem(vsock).UserNum, "BMKB", ""
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
            SetUserData SocketItem(vsock).UserNum, "BMKB", (GetRest(Line, 1))
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotMSG_ChanPT"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
            SetUserData SocketItem(vsock).UserNum, "BMPT", ""
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
            SetUserData SocketItem(vsock).UserNum, "BMPT", (GetRest(Line, 1))
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotMSG_FloodMSG"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
            SetUserData SocketItem(vsock).UserNum, "BMFP", ""
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
            SetUserData SocketItem(vsock).UserNum, "BMFP", (GetRest(Line, 1))
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotMSG_ShowStatus"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
            SetUserData SocketItem(vsock).UserNum, "BMSS", "1"
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
            SetUserData SocketItem(vsock).UserNum, "BMSS", ""
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case Else
        TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion: SocketItem(vsock).CurrentQuestion = "ChooseSetting"
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_PersonalSetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowPSActions vsock
  End If
End Sub

Private Sub ListPersonalSettings2(vsock As Long) ' : AddStack "Setups_ListPersonalSettings(" & vsock & ")"
Dim u As Long, TempStr As String, TempStr2 As String
' Messages for channel ops:
'    8. Joins/parts in channels..: #smof #animexx
'    9. Kicks/bans in channels...: ALL
'   10. Public talk in channels..: #animexx
'   11. Flood protection messages: ALL
  TU vsock, " Shown bot messages:"
  TU vsock, "   2 1) Additional help messages14..: " & IIf(GetSockFlag(vsock, SF_ExtraHelp) = SF_YES, "3ON ", "4OFF")
  TU vsock, "   2 2) Local user joins/parts14....: " & IIf(GetSockFlag(vsock, SF_Local_JP) = SF_YES, "3ON ", "4OFF")
  TU vsock, "   2 3) Local party line talk14.....: " & IIf(GetSockFlag(vsock, SF_Local_Talk) = SF_YES, "3ON ", "4OFF")
  TU vsock, "   2 4) Local bot connections14.....: " & IIf(GetSockFlag(vsock, SF_Local_Bot) = SF_YES, "3ON ", "4OFF")
  TU vsock, "   2 5) Botnet user joins/parts14...: " & IIf(GetSockFlag(vsock, SF_Botnet_JP) = SF_YES, "3ON ", "4OFF")
  TU vsock, "   2 6) Botnet party line talk14....: " & IIf(GetSockFlag(vsock, SF_Botnet_Talk) = SF_YES, "3ON ", "4OFF")
  TU vsock, "   2 7) Botnet bot connections14....: " & IIf(GetSockFlag(vsock, SF_Botnet_Bot) = SF_YES, "3ON ", "4OFF")
  TU vsock, " For channel ops:"
  TU vsock, "   2 8) Default channel14...........: " & IIf(GetUserData(SocketItem(vsock).UserNum, "BMDF", "") = "", "14<none>", GetUserData(SocketItem(vsock).UserNum, "BMDF", ""))
  TU vsock, "   2 9) Joins/parts in channel14....: " & IIf(GetUserData(SocketItem(vsock).UserNum, "BMJP", "") = "", "14<none>", GetUserData(SocketItem(vsock).UserNum, "BMJP", ""))
  TU vsock, "   210) Kicks/Modes in channel14....: " & IIf(GetUserData(SocketItem(vsock).UserNum, "BMKB", "") = "", "14<none>", GetUserData(SocketItem(vsock).UserNum, "BMKB", ""))
  TU vsock, "   211) Public talk in channel14....: " & IIf(GetUserData(SocketItem(vsock).UserNum, "BMPT", "") = "", "14<none>", GetUserData(SocketItem(vsock).UserNum, "BMPT", ""))
  TU vsock, "   212) Flood protection messages14.: " & IIf(GetUserData(SocketItem(vsock).UserNum, "BMFP", "") = "", "14<none>", GetUserData(SocketItem(vsock).UserNum, "BMFP", ""))
  TU vsock, "   213) Status for default channel: " & IIf(GetUserData(SocketItem(vsock).UserNum, "BMSS", "") = "", "4OFF", "3ON ")
  If MatchFlags(SocketItem(vsock).Flags, "+m") Then
    TU vsock, " For masters:"
    TU vsock, "   214) PRIVMSGs to the bot14......: " & IIf(GetSockFlag(vsock, SF_PrivToBot) = SF_YES, "3ON ", "4OFF")
    TU vsock, "   215) Other user's commands14....: " & IIf(GetSockFlag(vsock, SF_UserCommands) = SF_YES, "3ON ", "4OFF")
  End If
  TU vsock, EmptyLine
End Sub

Public Sub ShowPSActions(vsock As Long) ' : AddStack "Setups_ShowPSActions(" & vsock & ")"
  ListPersonalSettings vsock
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '1 on' or '5 off, 6 off, 7 off'."
  TU vsock, EmptyLine
End Sub

' Bot Setup
'-——————————————————————————————————————————————————————- -- -  -
Public Sub BotSetup(vsock As Long, Line As String)
Dim u As Long, InstantSet As Boolean, HostOnly As String, u2 As Long
Dim Attr As WIN32_FIND_DATA, FileName As String
InstantSetIt2:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving bot info..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
          Case "1"
              SocketItem(vsock).CurrentQuestion = "RealName"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "21) Real name"
              TU vsock, "  14 Here's an example for this setting: 10Just ask! :-)"
              TU vsock, "  14 Specifies the 'real name' shown in my /WHOIS info."
              TU vsock, "  14 You can use colors in the real name!"
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' to delete it or '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "QuitMSG"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "22) Quit message"
              TU vsock, "  14 Here's an example for this setting: 10Gotta go, bye!"
              TU vsock, "  14 Specifies the quit message that is shown in my IRC"
              TU vsock, "  14 channels when I'm leaving (due to a bot winsock2_shutdown)."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' to delete it or '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "IdentCMD"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "23) IDENT command"
              TU vsock, "  14 Here's an example for this setting: 10ITSME"
              TU vsock, "  14 This setting changes my 'IDENT' command, the"
              TU vsock, "  14 command that is used to make me recognize one of"
              TU vsock, "  14 my users from a new host. If you change the /MSG"
              TU vsock, "  14 command, it's harder to abuse a user's password"
              TU vsock, "  14 without knowing my telnet port. If you don't"
              TU vsock, "  14 like the IDENT command at all, just turn it off"
              TU vsock, "  14 by typing '*'."
              TU vsock, EmptyLine
              TU vsock, "  14 Please don't forget to tell your users about"
              TU vsock, "  14 IDENT command changes."
              TU vsock, EmptyLine
              TU vsock, "Choose a new command, type 'x' to set it to 'IDENT',"
              TU vsock, "type '*' to disable IDENT or type '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "VersionReply"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "24) VERSION reply"
              TU vsock, "  14 This setting changes my 'VERSION' reply which"
              TU vsock, "  14 everybods gets who requests my version on IRC"
              TU vsock, "  14 via '/ctcp <mynick> VERSION'."
              TU vsock, "  14 The default reply is 'AnGeL Bot Vx.y.z'. If"
              TU vsock, "  14 you don't want a CTCP VERSION reply at all,"
              TU vsock, "  14 type '*' to turn it off."
              TU vsock, EmptyLine
              TU vsock, "Choose a new version reply, type 'x' to set the default,"
              TU vsock, "type '*' to disable it or type '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "BotAdmin"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "25) Bot admin"
              TU vsock, "  14 This setting changes the admin information"
              TU vsock, "  14 which is shown when somebody in the botnet"
              TU vsock, "  14 enters '.who " & BotNetNick & "'."
              TU vsock, "  14 You should enter your nickname and e-mail"
              TU vsock, "  14 address here so that people can contact you"
              TU vsock, "  14 when there are problems with your bot."
              TU vsock, EmptyLine
              TU vsock, "Choose a new admin line or type '0' to cancel."
          Case "6"
              SocketItem(vsock).CurrentQuestion = "1stNick"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "26) 1st Nick"
              TU vsock, "  14 This setting changes the nickname I always try"
              TU vsock, "  14 to use on IRC (as long as it's free). Please"
              TU vsock, "  14 choose a nick that is not used by another bot -"
              TU vsock, "  14 otherwise I'll always try to gain this nick"
              TU vsock, "  14 without success."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "7"
              SocketItem(vsock).CurrentQuestion = "2ndNick"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "27) 2nd Nick"
              TU vsock, "  14 If my 1st Nick is not available, I'll use this"
              TU vsock, "  14 one. If this nick is not free, too, I'll use"
              TU vsock, "  14 a nick like 'AnGeL3992'... so please specify"
              TU vsock, "  14 nicks that should be available."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "8"
              SocketItem(vsock).CurrentQuestion = "Username"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "28) Username"
              TU vsock, "  14 Specifies my username or so-called 'ident'"
              TU vsock, "  14 on IRC. This is the *!username@* part of my"
              TU vsock, "  14 hostmask. You'll need to '.reconnect' me to"
              TU vsock, "  14 make changes to my username visible."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting or type '0' to cancel."
          Case "9"
              If BotCount > 1 Then TU vsock, "5*** You can't change the botnet nick while I'm linked to other bots.": TU vsock, "5*** Leave BotSetup now, use '.unlink *' to unlink all bots and try again.": Exit Sub
              SocketItem(vsock).CurrentQuestion = "BotnetNick"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "29) Botnet nick"
              TU vsock, "  14 Specifies the nickname I use in the botnet."
              TU vsock, "  14 This nick has got nothing to do with my IRC"
              TU vsock, "  14 nick. If there is no botnet nick set, I'll"
              TU vsock, "  14 use my 1st Nick."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' to delete it or '0' to cancel."
          Case "10"
              SocketItem(vsock).CurrentQuestion = "UserPort"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "210) User port"
              TU vsock, "  14 Specifies the port on which I can receive telnet"
              TU vsock, "  14 connections initiated by users. This is a way to"
              TU vsock, "  14 connect to me even if I'm not on IRC - just type"
              TU vsock, "  14 'telnet <host> <port>' at your command line."
              TU vsock, EmptyLine
              TU vsock, "  14 By disabling the user port, users can't connect"
              TU vsock, "  14 to me via telnet anymore. If you disable both"
              TU vsock, "  14 the IDENT command and the user port, it's almost"
              TU vsock, "  14 impossible to gain illegal access to me (e.g. by"
              TU vsock, "  14 using a stolen password)."
              TU vsock, EmptyLine
              TU vsock, "Choose a new user port, type 'x' to set it to '23',"
              TU vsock, "type '*' to disable it or type '0' to cancel."
          Case "11"
              SocketItem(vsock).CurrentQuestion = "BotPort"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "211) Bot port"
              TU vsock, "  14  Specifies the port on which I can receive telnet"
              TU vsock, "  14  connections initiated by other bots. When two"
              TU vsock, "  14  bots should be linked together, at least one bot"
              TU vsock, "  14  has to know the other bot's bot port."
              TU vsock, EmptyLine
              TU vsock, "  14  If you don't want other bots to connect to me,"
              TU vsock, "  14  you can disable the bot port by typing '*'."
              TU vsock, EmptyLine
              TU vsock, "Choose a new bot port, type 'x' to set it to '3333',"
              TU vsock, "type '*' to disable it or type '0' to cancel."
          Case "12"
              SocketItem(vsock).CurrentQuestion = "FileArea"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "212) File area"
              TU vsock, "  14  If this setting is turned on, users with the"
              TU vsock, "  14  flag '+i' can join the file area by typing"
              TU vsock, "  14  '.files'. My logs and other files can be"
              TU vsock, "  14  downloaded there."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "13"
              SocketItem(vsock).CurrentQuestion = "LocalIP"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "213) Local IP"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  '1': Get local IP by resolving it myself."
              TU vsock, "  14  '2': Get local IP through my IRC server."
              TU vsock, "  14  or : Enter an IP/Hostname yourself."
              TU vsock, EmptyLine
              TU vsock, "Type '1', '2', an IP address or '0' to cancel."
          Case "14"
              SocketItem(vsock).CurrentQuestion = "PumpDCC"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "214) Pump DCC"
              TU vsock, "  14  Specifies the size of data packets sent by"
              TU vsock, "  14  me via DCC (i.e. when somebody requests a"
              TU vsock, "  14  file from the file area). You can speed up"
              TU vsock, "  14  my DCC Sends by specifying higher values."
              TU vsock, EmptyLine
              TU vsock, "  14  Attention: If you specify values that are"
              TU vsock, "  14  too high, I'll get ping timeouts on IRC."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' to set it to '4096'"
              TU vsock, "or type '0' to cancel."
          Case "15"
              If Not WinNTOS Then
                TU vsock, "5*** Sorry, this setting is only available under Windows NT."
                Exit Sub
              Else
                SocketItem(vsock).CurrentQuestion = "NTService"
                If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
                TU vsock, "215) NT service"
                TU vsock, "  14  Here are the possible values for this setting:"
                TU vsock, "  14  'on' : Registers this bot as a Windows NT"
                TU vsock, "  14         service. As a service, the bot will be"
                TU vsock, "  14         automatically started as soon as the OS"
                TU vsock, "  14         is ready, even before a user has logged"
                TU vsock, "  14         on. The bot will be only visible in the"
                TU vsock, "  14         task manager then."
                TU vsock, "  14  'off': Unregisters the NT service. You will"
                TU vsock, "  14         have to start the bot manually after"
                TU vsock, "  14         your next reboot."
                TU vsock, EmptyLine
                TU vsock, "Type 'on', 'off' or '0' to cancel."
              End If
          Case "16"
              SocketItem(vsock).CurrentQuestion = "BanLimit"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "216) Ban Limit"
              TU vsock, "  14  Specifies the maximum number of bans that"
              TU vsock, "  14  should be set in an IRC channel. If the"
              TU vsock, "  14  number of bans in a channel raises above"
              TU vsock, "  14  this limit, no further bans will be set."
              TU vsock, "  14  The maximum ban number on IRCnet is '20'."
              TU vsock, EmptyLine
              TU vsock, "Choose a new setting, type 'x' to set it to '20'"
              TU vsock, "or type '0' to cancel."
          Case "17"
              SocketItem(vsock).CurrentQuestion = "FakeIDKick"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "217) Bogus kick"
              TU vsock, "  14  If this setting is turned on, I'll kick"
              TU vsock, "  14  and ban users joining with bogus idents."
              TU vsock, "  14  A bogus ident can contain colors, control"
              TU vsock, "  14  chars, '@' chars and so on."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "18"
              SocketItem(vsock).CurrentQuestion = "HideBot"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "218) Hide Bot"
              TU vsock, "  14  If this setting is turned on, I'll no respond"
              TU vsock, "  14  to unknown people."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
         Case "19"
              SocketItem(vsock).CurrentQuestion = "Language"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "219) Language"
              TU vsock, "  14  This setting allows you to choose a language"
              TU vsock, "  14  file. Most bot messages can be changed by"
              TU vsock, "  14  using language files (*.lng). You can get a"
              TU vsock, "  14  sample at www.angel-bot.de and customize"
              TU vsock, "  14  it to your needs. Just DCC send your file"
              TU vsock, "  14  to me and you'll be able to select it here."
              TU vsock, EmptyLine
              u = kernel32_FindFirstFileA(FileAreaHome & "" & "*.lng", Attr)
              If u <> -1 Then
                TU vsock, "  2  *** Available language files:"
                Do
                  If InStr(Attr.cFileName, Chr(0)) > 0 Then FileName = Left(Attr.cFileName, InStr(Attr.cFileName, Chr(0)) - 1) Else FileName = Attr.cFileName
                  TU vsock, "      " & ParamX(FileName, ".", 1)
                  If kernel32_FindNextFileA(u, Attr) = 0 Then Exit Do
                Loop
                kernel32_FindClose u
              Else
                TU vsock, "  2  *** Available language files:"
                TU vsock, "      <none>"
              End If
              TU vsock, EmptyLine
              TU vsock, "Enter the name of one of the language files listed above,"
              TU vsock, "type 'x' to select English or type '0' to cancel."
              TU vsock, EmptyLine
          Case "20"
              SocketItem(vsock).CurrentQuestion = "StrictHost"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "220) StrictHost"
              TU vsock, "  14  If this setting is turned on, the leading"
              TU vsock, "  14  chars '~-+^=' in idents will be regarded"
              TU vsock, "  14  when matching hostmasks. You'll have to"
              TU vsock, "  14  add your users with a star in front of"
              TU vsock, "  14  the ident ('*!10*14ident@host.com') then."
              TU vsock, "  14  If this setting is turned off (default),"
              TU vsock, "  14  you can add them as '*!ident@host.com'"
              TU vsock, "  14  and they will match even if their host"
              TU vsock, "  14  is 'xyz!10~14ident@host.com'."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "21"
              SocketItem(vsock).CurrentQuestion = "CMDPrefix"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "221) CommandPrefix"
              TU vsock, "  14  Current prefix is '10" & CommandPrefix & "14'."
              TU vsock, EmptyLine
              TU vsock, "Choose a new prefix or type '0' to cancel."
          Case "22"
              SocketItem(vsock).CurrentQuestion = "Range"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "222) Port Range"
              TU vsock, "  14  Current DCC-PortRange is '10" & IIf(PortRange = "0", "1024-3677", PortRange) & "14'."
              TU vsock, "  14  You can change it by entering a new range of Ports."
              TU vsock, EmptyLine
              TU vsock, "Choose a new range (like '4000-4500'), type 'x' to set default or type '0' to cancel."
          Case "23"
              SocketItem(vsock).CurrentQuestion = "MaxLogAge"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "223) Log History"
              TU vsock, "  14  Current Log History is '10" & LogMaxAge & "14' days."
              TU vsock, "  14  You can change it by entering a new count of days."
              TU vsock, EmptyLine
              TU vsock, "Choose a new range (like '30') or type '0' to cancel."
          Case "24"
              SocketItem(vsock).CurrentQuestion = "RestrictCycle"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "224) Restricted-Cycle"
              TU vsock, "  14  If turned on, the bot would try to cycle servers if"
              TU vsock, "  14  the connection is restricted (+r)."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "25"
              SocketItem(vsock).CurrentQuestion = "AutoNetSetup"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "225) Auto-NetSetup"
              TU vsock, "  14  If turned on, the bot would change its netsettings"
              TU vsock, "  14  each time it connects. If turned of you should specify"
              TU vsock, "  14  all items in the NetSetup."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "26"
              SocketItem(vsock).CurrentQuestion = "UseIDENTD"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "226) Use IdentD"
              TU vsock, "  14  If turned on, the bot will open an IDENT Server"
              TU vsock, "  14  each time it connects."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "27"
              SocketItem(vsock).CurrentQuestion = "RouterWorkAround"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "227) Router Workaround"
              TU vsock, "  14  If turned on, the bot will initialize DCC Chats"
              TU vsock, "  14  to people with same Host on IRC with the local"
              TU vsock, "  14  Networks IP Adress."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "28"
              SocketItem(vsock).CurrentQuestion = "BaseFlags"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "228) BaseFlags"
              TU vsock, "  14  Changes the default flags for new Users (default: +fp)"
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in BotSetup!"
              TU vsock, "5*** Please enter a valid number."
              Exit Sub
        End Select
    Case "RealName"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              RealName = "Frag doch! :)"
              DeletePPString "Identification", "RealName", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              RealName = Line
              WritePPString "Identification", "RealName", RealName, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "QuitMSG"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString "Server", "QuitMessage", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Len(Line) > 80 Then
                TU vsock, "5*** Sorry, this quit message is " & CStr(Len(Line) - 80) & " characters too long. Try again."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              Else
                WritePPString "Server", "QuitMessage", Line, AnGeL_INI
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              End If
        End Select
    Case "IdentCMD"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              IdentCommand = "IDENT"
              DeletePPString "Identification", "IdentCommand", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "*"
              IdentCommand = "°NONE"
              WritePPString "Identification", "IdentCommand", IdentCommand, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              IdentCommand = Param(Line, 1)
              WritePPString "Identification", "IdentCommand", IdentCommand, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "VersionReply"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              VersionReply = "AnGeL Bot " & BotVersion + IIf(ServerNetwork <> "", "+" & ServerNetwork, "") & " - Copyright " & CpString
              DeletePPString "Server", "VersionReply", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "*"
              VersionReply = "°NONE"
              WritePPString "Server", "VersionReply", VersionReply, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              VersionReply = Line
              WritePPString "Server", "VersionReply", VersionReply, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotAdmin"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString "Identification", "Admin", Line, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "1stNick"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Len(Param(Line, 1)) > ServerNickLen Then
                TU vsock, "5*** A nick can't be longer than " & CStr(ServerNickLen) & " characters!"
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              If IsValidNick(Param(Line, 1)) = False Then
                TU vsock, "5*** """ & Param(Line, 1) & """ Erroneous Nickname"
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              PrimaryNick = Param(Line, 1)
              WritePPString "Identification", "PrimaryNick", PrimaryNick, AnGeL_INI
              SendLine "NICK " & Param(Line, 1), 2
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "CMDPrefix"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              CommandPrefix = LCase(Param(Line, 1))
              WritePPString "Others", "CMDPrefix", CommandPrefix, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "2ndNick"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Len(Param(Line, 1)) > ServerNickLen Then
                TU vsock, "5*** A nick can't be longer than " & CStr(ServerNickLen) & " characters!"
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              If IsValidNick(Param(Line, 1)) = False Then
                TU vsock, "5*** """ & Param(Line, 1) & """ Erroneous Nickname"
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              SecondaryNick = Param(Line, 1)
              WritePPString "Identification", "SecondaryNick", SecondaryNick, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "Username"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If IsValidIdent(Param(Line, 1)) = False Then TU vsock, "5*** """ & Param(Line, 1) & """ Erroneous Username": Exit Sub
              If Len(Param(Line, 1)) > 8 Then TU vsock, "5*** A username can't be longer than 8 characters!": Exit Sub
              WritePPString "Identification", "Identd", Param(Line, 1), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "BotnetNick"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              If BotCount > 1 Then
                TU vsock, "5*** You can't change the botnet nick while I'm linked to other bots.": TU vsock, "5*** Leave BotSetup now, use '.unlink *' to unlink all bots and try again."
              Else
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then If LCase(SocketItem(u).OnBot) = LCase(BotNetNick) Then SocketItem(u).OnBot = PrimaryNick
                Next u
                For u = 1 To BotCount
                  If LCase(Bots(u).Nick) = LCase(BotNetNick) Then Bots(u).Nick = PrimaryNick
                Next u
                BotNetNick = PrimaryNick
                DeletePPString "Identification", "BotNetNick", AnGeL_INI
              End If
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If BotCount > 1 Then
                TU vsock, "5*** You can't change the botnet nick while I'm linked to other bots.": TU vsock, "5*** Leave BotSetup now, use '.unlink *' to unlink all bots and try again."
              Else
                If Len(Param(Line, 1)) > ServerNickLen Then TU vsock, "5*** A nick can't be longer than " & CStr(ServerNickLen) & " characters!": Exit Sub
                If IsValidNick(Param(Line, 1)) = False Then TU vsock, "5*** """ & Param(Line, 1) & """ Erroneous Nickname": Exit Sub
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then If LCase(SocketItem(u).OnBot) = LCase(BotNetNick) Then SocketItem(u).OnBot = Param(Line, 1)
                Next u
                For u = 1 To BotCount
                  If LCase(Bots(u).Nick) = LCase(BotNetNick) Then Bots(u).Nick = Param(Line, 1)
                Next u
                BotNetNick = Param(Line, 1)
                WritePPString "Identification", "BotNetNick", BotNetNick, AnGeL_INI
              End If
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "UserPort"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "*"
              TelnetPort = 0
              WritePPString "TelNet", "UserPort", "0", AnGeL_INI
              If TelnetSocket > 0 Then
                RemoveSocket TelnetSocket, 0, "", True
              End If
              TelnetSocket = -1
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString "TelNet", "UserPort", AnGeL_INI
              If TelnetSocket > 0 Then
                RemoveSocket TelnetSocket, 0, "", True
              End If
              TelnetSocket = AddSocket
              u = ListenTCP(TelnetSocket, TelnetPort)
              If u = 0 Then
                SocketItem(TelnetSocket).RegNick = "<TELNET>"
                SetSockFlag TelnetSocket, SF_Status, SF_Status_TelnetListen
              Else
                RemoveSocket TelnetSocket, 0, "", True
                TelnetSocket = -1
              End If
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Not IsNumeric(Param(Line, 1)) Then
                TU vsock, "5*** Please enter a valid port number."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              
              u = CLng(Param(Line, 1))
              If IsListenPortInUse(u) Then
                TU vsock, "5*** This port is already in use! Please choose another one."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting": Exit Sub
              Else
                TelnetPort = u
                If TelnetSocket > 0 Then
                  RemoveSocket TelnetSocket, 0, "", True
                End If
                TelnetSocket = AddSocket
                If ListenTCP(TelnetSocket, u) = 0 Then
                  SocketItem(TelnetSocket).RegNick = "<TELNET>"
                  SetSockFlag TelnetSocket, SF_Status, SF_Status_TelnetListen
                Else
                  RemoveSocket TelnetSocket, 0, "", True
                  TelnetSocket = -1
                End If
                WritePPString "TelNet", "UserPort", CStr(TelnetPort), AnGeL_INI
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              End If
        End Select
    Case "BotPort"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "*"
              BotnetPort = 0
              WritePPString "TelNet", "Port", "0", AnGeL_INI
              If BotnetSocket > 0 Then
                RemoveSocket BotnetSocket, 0, "", True
              End If
              BotnetSocket = -1
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              BotnetPort = 3333
              DeletePPString "TelNet", "Port", AnGeL_INI
              If BotnetSocket > 0 Then
                RemoveSocket BotnetSocket, 0, "", True
              End If
              BotnetSocket = AddSocket
              If ListenTCP(BotnetSocket, BotnetPort) = 0 Then
                SocketItem(BotnetSocket).RegNick = "<BOTNET>"
                SetSockFlag BotnetSocket, SF_Status, SF_Status_BotnetListen
              Else
                RemoveSocket BotnetSocket, 0, "", True
                BotnetSocket = -1
              End If
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Not IsNumeric(Param(Line, 1)) Then
                TU vsock, "5*** Please enter a valid port number."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              u = CLng(Param(Line, 1))
              If IsListenPortInUse(u) Then
                TU vsock, "5*** This port is already in use! Please choose another one."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting": Exit Sub
              Else
                BotnetPort = u
                If BotnetSocket > 0 Then
                  RemoveSocket BotnetSocket, 0, "", True
                End If
                BotnetSocket = AddSocket
                If ListenTCP(BotnetSocket, BotnetPort) = 0 Then
                  SocketItem(BotnetSocket).RegNick = "<BOTNET>"
                  SetSockFlag BotnetSocket, SF_Status, SF_Status_BotnetListen
                Else
                  RemoveSocket BotnetSocket, 0, "", True
                  BotnetSocket = -1
                End If
                WritePPString "TelNet", "Port", CStr(BotnetPort), AnGeL_INI
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              End If
        End Select
    Case "FileArea"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              FileAreaEnabled = True
              WritePPString "Others", "FileArea", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              FileAreaEnabled = False
              WritePPString "Others", "FileArea", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "LocalIP"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "1"
              WritePPString "Others", "ResolveIP", Param(Line, 1), AnGeL_INI
              ResolveIP = Param(Line, 1)
              MyIP = IrcGetLongIp(GetLastLocalIP)
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "2"
              WritePPString "Others", "ResolveIP", Param(Line, 1), AnGeL_INI
              ResolveIP = Param(Line, 1)
              MyIP = IrcGetLongIp(GetLastLocalIP)
              HostOnly = Mask(MyHostmask, 11)
              HostOnly = IrcGetLongIp(GetCacheIP(HostOnly, True))
              If HostOnly <> MyIP Then
                If HostOnly <> "4294967295" Then MyIP = HostOnly
              End If
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              MyIP = IrcGetLongIp(GetCacheIP(Param(Line, 1), True))
              If IrcGetAscIp(MyIP) <> GetCacheIP(Param(Line, 1), True) Then
                TU vsock, "5*** Invalid IP/Hostname address."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              WritePPString "Others", "ResolveIP", Param(Line, 1), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "PumpDCC"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              PumpDCC = 4096
              DeletePPString "Others", "PumpDCC", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Not IsNumeric(Param(Line, 1)) Then
                TU vsock, "5*** Please enter a valid block size."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              If CLng(Param(Line, 1)) > 4096 Then
                TU vsock, "5*** Value to large."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              PumpDCC = CLng(Param(Line, 1))
              WritePPString "Others", "PumpDCC", CStr(PumpDCC), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "NTService"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              NTServiceName = GetPPString("Others", "NTService", "", AnGeL_INI)
              If NTServiceName <> "" Then UninstallService
              NTServiceName = "AnGeL-"
              For u = 1 To 8
                NTServiceName = NTServiceName + Chr(Asc("0") + Int(Rnd * 10))
              Next u
              WritePPString "Others", "NTService", NTServiceName, AnGeL_INI
              InstallService
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              NTServiceName = GetPPString("Others", "NTService", "", AnGeL_INI)
              If NTServiceName <> "" Then UninstallService
              DeletePPString "Others", "NTService", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "BanLimit"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              BanLimit = 20
              DeletePPString "Others", "BanLimit", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Not IsNumeric(Param(Line, 1)) Then
                TU vsock, "5*** Please enter a valid ban limit."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              ElseIf CLng(Param(Line, 1)) <= 5 Then
                TU vsock, "5*** This value is too low. Please enter a value higher than 5."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              BanLimit = CLng(Param(Line, 1))
              WritePPString "Others", "BanLimit", CStr(BanLimit), AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "FakeIDKick"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              FakeIDKick = True
              WritePPString "Others", "FakeIDKick", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              FakeIDKick = False
              WritePPString "Others", "FakeIDKick", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "HideBot"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              HideBot = True
              WritePPString "Others", "HideBot", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              HideBot = False
              WritePPString "Others", "HideBot", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "BaseFlags"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "+fp", "x"
            BaseFlags = "+fp"
            DeletePPString "Others", "BaseFlags", AnGeL_INI
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
            BaseFlags = CombineFlags("", Param(Line, 1))
            WritePPString "Others", "BaseFlags", BaseFlags, AnGeL_INI
            If BaseFlags = "" Then
              TU vsock, "3*** Base Flags deleted! New Users won't have any flags."
            Else
              TU vsock, "3*** Base Flags for new users now: 10" & BaseFlags
            End If
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "Language"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "english", "x"
              HideBot = True
              DeletePPString "Others", "Language", AnGeL_INI
              InitLanguage
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              On Local Error Resume Next
              FileName = ParamX(GetFileName(Dir(FileAreaHome & "" & Param(Line, 1) & ".lng")), ".", 1)
              If FileName = "" Then
                TU vsock, "5*** Sorry, I couldn't find this language file."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              WritePPString "Others", "Language", FileName, AnGeL_INI
              If InitLanguage = False Then
                TU vsock, "5*** Failed to load language '" & Param(Line, 1) & "'."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              Else
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              End If
        End Select
    Case "Range"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              PortRange = "0"
              DeletePPString "Others", "PortRange", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InStr(Line, "-") <> 0 And IsNumeric(ParamX(Line, "-", 1)) = True And IsNumeric(ParamX(Line, "-", 2)) = True And ParamX(Line, "-", 1) < ParamX(Line, "-", 2) Then
                WritePPString "Others", "PortRange", Line, AnGeL_INI
                PortRange = Line
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              Else
                TU vsock, "5*** Please enter a Range, 'x' for defaults or '0' to get back to menu."
              End If
        End Select
    Case "MaxLogAge"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If IsNumeric(Param(Line, 1)) Then
                WritePPString "Others", "MaxLogAge", Param(Line, 1), AnGeL_INI
                LogMaxAge = Param(Line, 1)
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              Else
                TU vsock, "5*** Please enter a numeric value."
              End If
        End Select
    Case "StrictHost"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              StrictHost = True
              WritePPString "Others", "StrictHost", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              StrictHost = False
              WritePPString "Others", "StrictHost", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "FakeIDKick"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              FakeIDKick = True
              WritePPString "Others", "FakeIDKick", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              FakeIDKick = False
              WritePPString "Others", "FakeIDKick", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "RestrictCycle"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              RestrictCycle = True
              WritePPString "Others", "RestrictCycle", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              RestrictCycle = False
              WritePPString "Others", "RestrictCycle", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "AutoNetSetup"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              AutoNetSetup = True
              WritePPString "Others", "AutoNetSetup", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              AutoNetSetup = False
              WritePPString "Others", "AutoNetSetup", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "UseIDENTD"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              UseIDENTD = True
              WritePPString "Others", "UseIDENTD", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              UseIDENTD = False
              WritePPString "Others", "UseIDENTD", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "RouterWorkAround"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              RouterWorkAround = True
              WritePPString "Others", "RouterWorkAround", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              RouterWorkAround = False
              WritePPString "Others", "RouterWorkAround", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case Else
        TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_BotSetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowBSActions vsock
  End If
End Sub
Sub KISetup(vsock As Long, Line As String) ' : AddStack "Setups_KISetup(" & vsock & ", " & Line & ")"
Dim InstantSet As Boolean, u As Long
On Error GoTo KIErr
InstantSetIt:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving KI info..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
          Case "1"
              SocketItem(vsock).CurrentQuestion = "FirstName"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) First Name"
              TU vsock, "  14 Here you can define the first name of the Bot"
              TU vsock, "  14  which whould be used in KI-Answers."
              TU vsock, EmptyLine
              TU vsock, "2Which name should I use?"
              TU vsock, "Choose a new setting, or '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "LastName"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Last Name"
              TU vsock, "  14 Here you can define the last name of the Bot"
              TU vsock, "  14 which whould be used in KI-Answers."
              TU vsock, EmptyLine
              TU vsock, "2Which name should I use?"
              TU vsock, "Choose a new setting, or '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "Gender"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Gender"
              TU vsock, "  14 Please specify the Gender of the Bot"
              TU vsock, "  14 if you set up a female 'First Name' you"
              TU vsock, "  14 whould use 'f' ;-)"
              TU vsock, EmptyLine
              TU vsock, "2Please type 'f' for 'female', 'm' for 'male' or '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "Age"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Age"
              TU vsock, "  14 Please enter the Age of these Bot."
              TU vsock, "  14 15 to 20 is a good age for female Bots ;-)"
              TU vsock, EmptyLine
              TU vsock, "Choose a new age, or '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "City"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) City/Town"
              TU vsock, "  14 Please enter the City/Town of the Bot"
              TU vsock, EmptyLine
              TU vsock, "Choose a new City/Town, or '0' to cancel."
          Case "6"
              SocketItem(vsock).CurrentQuestion = "Country"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Country"
              TU vsock, "  14 Please enter the Country of the Bot"
              TU vsock, EmptyLine
              TU vsock, "Choose a new Country, or '0' to cancel."
          Case "7"
              SocketItem(vsock).CurrentQuestion = "HideBot"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "218) Hide Bot"
              TU vsock, "  14  If this setting is turned on, I'll no longer"
              TU vsock, "  14  reply to msg commands, dcc chats etc. from"
              TU vsock, "  14  unknown users. You can use this setting in"
              TU vsock, "  14  combination with a different version reply"
              TU vsock, "  14  to 'hide' the bot -> nobody will be able to"
              TU vsock, "  14  see that I'm a bot."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in KISetup!"
              TU vsock, "5*** Please enter a KISetup number, or leave by typing '0'."
              Exit Sub
        End Select
    Case "FirstName"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString "KI", "FirstName", Line, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              KIFName = Line
        End Select
    Case "LastName"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString "KI", "LastName", Line, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              KILName = Line
        End Select
    Case "Gender"
        Select Case Param(Line, 1)
          Case "0"
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "f", "female", "w", "weiblich", "frau", "woman", "schlampe", "bitch", "maedchen", "mädchen", "g", "girl"
            WritePPString "KI", "Gender", "f", AnGeL_INI
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
            KIGender = "f"
          Case "m", "male", "männlich", "maennlich", "junge", "mann", "wichser", "pisser", "j"
            WritePPString "KI", "Gender", "m", AnGeL_INI
            SocketItem(vsock).CurrentQuestion = "ChooseSetting"
            KIGender = "m"
          Case Else
            TU vsock, "5*** Please enter 'f', 'm' or '0'."
        End Select
    Case "HideBot"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              HideBot = True
              WritePPString "Others", "HideBot", "yes", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              HideBot = False
              WritePPString "Others", "HideBot", "no", AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "Age"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString "KI", "Age", Line, AnGeL_INI
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              KIAge = Line
        End Select
    Case "City"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString "KI", "City", Line, AnGeL_INI
              KICity = Line
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "Country"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString "KI", "Country", Line, AnGeL_INI
              KICountry = Line
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case Else
        TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_KISetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowKIActions vsock
  End If
Exit Sub
KIErr:
  SendNote "AnGeL KISetup ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & SocketItem(vsock).CurrentQuestion & "] - ist aufgetreten."
  SendNote "AnGeL KISetup ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & SocketItem(vsock).CurrentQuestion & "] - ist aufgetreten."
  Err.Clear
End Sub
Private Sub ListPolicieSettings(vsock As Long)
Dim u As Long, TempStr As String, TempStr2 As String
  TU vsock, " 2 1) UserManagement14......: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("UserManagement") = True, "3YES", "4NO")) & _
            " 2 2) DataBase14............: " & IIf(AnGeLFiles.CommandAllowed("DataBase") = True, "3YES", "4NO")
  TU vsock, " 2 3) Resolves14............: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("Resolves") = True, "3YES", "4NO")) & _
            " 2 4) CreateObject14........: " & IIf(AnGeLFiles.CommandAllowed("Objects") = True, "3YES", "4NO")
  TU vsock, " 2 5) FileOperations14......: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("FileOperations") = True, "3YES", "4NO")) & _
            " 2 6) ExecuteCommand14......: " & IIf(AnGeLFiles.CommandAllowed("ExecuteCommand") = True, "3YES", "4NO")
  TU vsock, " 2 7) TimeOperations14......: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("TimeOperations") = True, "3YES", "4NO")) & _
            " 2 8) SocketOperations14....: " & IIf(AnGeLFiles.CommandAllowed("SocketOperations") = True, "3YES", "4NO")
  TU vsock, " 2 9) INIOperations14.......: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("INIOperations") = True, "3YES", "4NO")) & _
            " 210) Hooks14...............: " & IIf(AnGeLFiles.CommandAllowed("NotificationHooks") = True, "3YES", "4NO")
  TU vsock, " 211) WMI14.................: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("WMI") = True, "3YES", "4NO")) & _
            " 212) Native FileSys14......: " & IIf(AnGeLFiles.CommandAllowed("NativeFS") = True, "3YES", "4NO")
  TU vsock, " 213) BotInteractions14.....: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("BotInteractions") = True, "3YES", "4NO")) & _
            " 214) ChannelInteractions14.: " & IIf(AnGeLFiles.CommandAllowed("ChannelInteractions") = True, "3YES", "4NO")
  TU vsock, " 215) SessionChange14.......: " & Spaces2(7, IIf(AnGeLFiles.CommandAllowed("SessionChange") = True, "3YES", "4NO")) & _
            " 216) AddInModules14........: " & IIf(AnGeLFiles.CommandAllowed("AddInModules") = True, "3YES", "4NO")
  TU vsock, EmptyLine
End Sub
Private Sub ListBotSettings(vsock As Long) ' : AddStack "Setups_ListBotSettings(" & vsock & ")"
Dim u As Long, TempStr As String, TempStr2 As String
  TU vsock, " 2 1) Real name14......: " & GetPPString("Identification", "RealName", "Frag doch! :)      14(default)", AnGeL_INI)
  TU vsock, " 2 2) Quit message14...: " & GetPPString("Server", "QuitMessage", "AnGeL leaving...   14(default)", AnGeL_INI)
  TempStr = GetPPString("Identification", "IdentCommand", "IDENT              14(default)", AnGeL_INI)
  TU vsock, " 2 3) IDENT command14..: " & IIf(TempStr <> "°NONE", TempStr, "14<disabled>")
  TempStr = GetPPString("Server", "VersionReply", "AnGeL Bot " & BotVersion + IIf(ServerNetwork <> "", "+" & ServerNetwork, "") & " - Copyright " & CpString, AnGeL_INI)
  If Len(TempStr) > 33 Then TempStr = Left(TempStr, 30) & "14..."
  TU vsock, " 2 4) VERSION reply14..: " & IIf(TempStr <> "°NONE", TempStr, "14<disabled>")
  TU vsock, " 2 5) Bot admin14......: " & GetPPString("Identification", "Admin", "Nobody <no@mail.address.com>", AnGeL_INI)
  TU vsock, EmptyLine
  TU vsock, " 2 6) 1st Nick14...: " & Spaces2(10, GetPPString("Identification", "PrimaryNick", "<none>", AnGeL_INI)) & " 2 7) 2nd Nick14......: " & GetPPString("Identification", "SecondaryNick", "<none>", AnGeL_INI)
  TU vsock, " 2 8) Username14...: " & Spaces2(10, GetPPString("Identification", "Identd", "<none>", AnGeL_INI)) & " 2 9) Botnet Nick14...: " & GetPPString("Identification", "BotNetNick", "14<1st Nick>", AnGeL_INI)
  TempStr = Trim(GetPPString("TelNet", "UserPort", "23", AnGeL_INI)): If TempStr = "0" Then TempStr = "14<disabled>"
  TempStr2 = Trim(GetPPString("TelNet", "Port", "3333", AnGeL_INI)): If TempStr2 = "0" Then TempStr2 = "14<disabled>"
  TU vsock, " 210) User port14..: " & Spaces2(10, TempStr) & " 211) Bot port14......: " & TempStr2
  Select Case GetPPString("Others", "ResolveIP", "2", AnGeL_INI)
    Case "1": TempStr = "14<resolve>"
    Case "2": TempStr = "14<server>"
    Case Else: TempStr = GetPPString("Others", "ResolveIP", "2", AnGeL_INI)
  End Select
  TU vsock, " 212) File area14..: " & IIf(LCase(GetPPString("Others", "FileArea", "yes", AnGeL_INI)) = "yes", "3ON ", "4OFF") & "        213) Local IP14......: " & TempStr
  TU vsock, " 214) Pump DCC14...: " & Spaces2(10, GetPPString("Others", "PumpDCC", "4096", AnGeL_INI)) & " 215) NT service14....: " & IIf(GetPPString("Others", "NTService", "", AnGeL_INI) <> "", "3ON ", "4OFF")
  TU vsock, " 216) Ban limit14..: " & Spaces2(10, GetPPString("Others", "BanLimit", "20", AnGeL_INI)) & " 217) Bogus kick14....: " & IIf(LCase(GetPPString("Others", "FakeIDKick", "yes", AnGeL_INI)) = "yes", "3ON ", "4OFF")
  TU vsock, " 218) Hide bot14...: " & IIf(LCase(GetPPString("Others", "HideBot", "no", AnGeL_INI)) = "yes", "3ON ", "4OFF") & "        219) Language14......: " & GetPPString("Others", "Language", "English", AnGeL_INI)
  TU vsock, " 220) StrictHost14.: " & IIf(LCase(GetPPString("Others", "StrictHost", "no", AnGeL_INI)) = "yes", "3ON ", "4OFF") & "        221) Command-Prefix: " & GetPPString("Others", "CMDPrefix", "!", AnGeL_INI)
  TU vsock, " 222) PortRange14..: " & Spaces2(10, IIf(LCase(GetPPString("Others", "PortRange", "0", AnGeL_INI)) = "0", "14<default> ", GetPPString("Others", "PortRange", "0", AnGeL_INI))) & " 223) Log History14...: " & GetPPString("Others", "MaxLogAge", "30", AnGeL_INI)
  TU vsock, " 224) +r Cycle14...: " & Spaces2(10, IIf(LCase(GetPPString("Others", "RestrictCycle", "yes", AnGeL_INI)) = "yes", "3ON ", "4OFF")) & "    225) Auto-NetSetup14.: " & IIf(LCase(GetPPString("Others", "AutoNETSETUP", "yes", AnGeL_INI)) = "yes", "3ON ", "4OFF")
  TU vsock, " 226) Use IDENTD14.: " & Spaces2(10, IIf(LCase(GetPPString("Others", "UseIDENTD", "yes", AnGeL_INI)) = "yes", "3ON ", "4OFF")) & "    227) Router Fix14....: " & IIf(LCase(GetPPString("Others", "RouterWorkAround", "yes", AnGeL_INI)) = "yes", "3ON ", "4OFF")
  TU vsock, " 228) BaseFlags.14.: " & BaseFlags
  TU vsock, EmptyLine
End Sub
Private Sub ListNETSettings(vsock As Long) ' : AddStack "Setups_ListNETSettings(" & vsock & ")"
Dim u As Long, TempStr As String, TempStr2 As String
  TU vsock, " 2 1) " & MakeSettingText("Network name", 25) + GetPPString("NET", "NetworkName", "IRCNet" & vbTab & "14(default)", NET_INI)
  TU vsock, " 2 2) " & MakeSettingText("Nicklength", 25) + GetPPString("NET", "NickLength", "9" & vbTab & "14(default)", NET_INI)
  TU vsock, " 2 3) " & MakeSettingText("Max. Channel", 25) + GetPPString("NET", "MaxChan", "10" & vbTab & "14(default)", NET_INI)
  TU vsock, " 2 4) " & MakeSettingText("Supports x!x@x-PRIVMSG", 25) + IIf(Switch(GetPPString("NET", "UseFullAdress", "1", NET_INI)), "3YES", "4NO")
  TU vsock, " 2 5) " & MakeSettingText("Supports &servers", 25) + IIf(Switch(GetPPString("NET", "SplitDetection", "1", NET_INI)), "3YES", "4NO")
  TU vsock, " 2 6) " & MakeSettingText("Channel prefixes", 25) + GetPPString("NET", "ChanPrefixes", "#&!+" & vbTab & "14(default)", NET_INI)
  TU vsock, " 2 7) " & MakeSettingText("User prefixes", 25) + GetPPString("NET", "UserPrefixes", "(ov)@+" & vbTab & "14(default)", NET_INI)
  TU vsock, " 2 8) " & MakeSettingText("Channel modes", 25) + GetPPString("NET", "ChanModes", "imntl,k,b,sp" & vbTab & "14(default)", NET_INI)
  TU vsock, EmptyLine
End Sub
Private Sub ListAUTHSettings(vsock As Long) ' : AddStack "Setups_ListNETSettings(" & vsock & ")"
Dim u As Long, TempStr As String, TempStr2 As String
  TU vsock, " 2 1) " & MakeSettingText("Service name", 15) + IIf(GetPPString("AUTH", "Target", "", AnGeL_INI) = "", "14<disabled>", GetPPString("AUTH", "Target", "", AnGeL_INI))
  TU vsock, " 2 2) " & MakeSettingText("Auth command", 15) + GetPPString("AUTH", "Command", "IDENTIFY" & vbTab & "14(default)", AnGeL_INI)
  TU vsock, " 2 3) " & MakeSettingText("Parameter 1", 15) + GetPPString("AUTH", "Username", "14<none>", AnGeL_INI)
  TU vsock, " 2 4) " & MakeSettingText("Parameter 2", 15) + GetPPString("AUTH", "Password", "14<none>", AnGeL_INI)
  TU vsock, " 2 5) " & MakeSettingText("ReAUTH on join", 15) + IIf(Switch(GetPPString("AUTH", "ReAUTH", "1", AnGeL_INI)), "3YES", "4NO")
  TU vsock, EmptyLine
End Sub
Public Sub ShowBSActions(vsock As Long) ' : AddStack "Setups_ShowBSActions(" & vsock & ")"
  ListBotSettings vsock
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '10 on' or '4 FunnyBot'."
  TU vsock, EmptyLine
End Sub
Public Sub ShowNETActions(vsock As Long) ' : AddStack "Setups_ShowNETActions(" & vsock & ")"
  ListNETSettings vsock
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '1 IRCNet' or '6 #+!&'."
  TU vsock, EmptyLine
End Sub
Public Sub ShowKIActions(vsock As Long) ' : AddStack "Setups_ShowKIActions(" & vsock & ")"
  ListKISettings vsock
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '1 Michael' or '3 female'."
  TU vsock, EmptyLine
End Sub
Public Sub ShowAUTHActions(vsock As Long) ' : AddStack "Setups_ShowNETActions(" & vsock & ")"
  ListAUTHSettings vsock
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '1 AuthServ' or '3 secretpassword'."
  TU vsock, EmptyLine
End Sub
Public Sub ShowPolActions(vsock As Long)
  ListPolicieSettings vsock
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '1 on' or '3 off'."
  TU vsock, EmptyLine
End Sub

' Channel Setup
'-——————————————————————————————————————————————————————- -- -  -
Public Sub ChanSetup(vsock As Long, Line As String) ' : AddStack "Setups_ChanSetup(" & vsock & ", " & Line & ")"
Dim InstantSet As Boolean, u As Long, ChNum As Long, ChangMode As String
On Error GoTo ChanSErr
InstantSetIt:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving channel info..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
          Case "1"
              SocketItem(vsock).CurrentQuestion = "EnforceModes"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "21) Enforce Modes"
              TU vsock, "  14 Here's an example for this setting: 10+nt-iskl"
              TU vsock, "  14 This means that I'll always set the modes +nt and that"
              TU vsock, "  14 I'll try to prevent that the modes +i, +s, +k or +l are"
              TU vsock, "  14 set in the channel."
              TU vsock, EmptyLine
              TU vsock, "2Which modes should I enforce in this channel?"
              TU vsock, "Choose a new setting, type 'x' to delete it or '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "DefaultTopic"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "22) Default Topic"
              TU vsock, "  14 Here's an example for this setting: 10No topic set yet ;-)"
              TU vsock, "  14 This is the topic I will set if there's no topic set when"
              TU vsock, "  14 I join the channel. I will also protect this topic if"
              TU vsock, "  14 the 'Protect Topic' channel flag is set."
              TU vsock, EmptyLine
              TU vsock, "2What should be the default topic in this channel?"
              TU vsock, "Choose a new setting, type 'x' to delete it or '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "NewbieGreeting"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "23) Greet Newcomers"
              TU vsock, "  14 Here's an example for this setting: 10Welcome to #blubb!"
              TU vsock, "  14 If I see somebody for the first time in this channel,"
              TU vsock, "  14 I'll send this greeting via NOTICE to him/her."
              TU vsock, EmptyLine
              TU vsock, "2What should be the greeting message in this channel?"
              TU vsock, "Choose a new setting, type 'x' to delete it or '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "ProtectTopic"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "24) Protect Topic"
              TU vsock, "  14 If the topic protection is turned on, I'll always"
              TU vsock, "  14 try to set the default topic in this channel."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "ProtectFriends"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "25) Protect Friends"
              TU vsock, "  14 If the friend protection is turned on, I'll kickban"
              TU vsock, "  14 unknown users who kick my friends. I'll also deop"
              TU vsock, "  14 users who deop my friends."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "6"
              SocketItem(vsock).CurrentQuestion = "ColorKick"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "26) Color Kick"
              TU vsock, "  14 If the color kick is turned on, I'll kick unknown users"
              TU vsock, "  14 who use color codes in the channel."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "7"
              SocketItem(vsock).CurrentQuestion = "CloneKick"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "27) Clone Kick"
              TU vsock, "  14 If the clone kick is turned on, I'll kickban users with"
              TU vsock, "  14 more than one clone. For example, if a user loads ONE"
              TU vsock, "  14 clone, I'll do nothing. But if another clone joins the"
              TU vsock, "  14 channel, I'll kickban this user and all clones."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "8"
              SocketItem(vsock).CurrentQuestion = "AutoVoice"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "28) Autovoice if +m"
              TU vsock, "  14 Here are the possible values for this setting:"
              TU vsock, "  14 'OFF': Nobody will be automatically voiced."
              TU vsock, "  14 'ON' : All users who join the channel will be voiced"
              TU vsock, "  14        if the channel is +m (moderated)."
              TU vsock, "  14 'EXT': Only *registered* users will be voiced when"
              TU vsock, "  14        they join this channel while it's moderated."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'ext', 'off' or '0' to cancel."
          Case "9"
              SocketItem(vsock).CurrentQuestion = "DeopUnknownUsers"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "29) Deop unknown users"
              TU vsock, "  14 Here are the possible values for this setting:"
              TU vsock, "  14 'OFF': Unknown users won't be deopped at all."
              TU vsock, "  14 'ON' : Deops only unknown users opped by other users."
              TU vsock, "  14        Unknown users opped by bots won't be deopped."
              TU vsock, "  14 'EXT': *Always* deops unknown users, even if they are"
              TU vsock, "  14        opped by bots or bot owners. Very secure."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'ext', 'off' or '0' to cancel."
          Case "10"
              SocketItem(vsock).CurrentQuestion = "ReactToSeen"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "210) React to " & CommandPrefix & "seen"
              TU vsock, "   14 This setting determines my reaction to 2" & CommandPrefix & "seen"
              TU vsock, "   14 messages in the channel. Here are the possible"
              TU vsock, "   14 values for this setting:"
              TU vsock, "   14 'OFF': No reaction to " & CommandPrefix & "seen channel messages."
              TU vsock, "   14 '1'  : Default answer in a channel message."
              TU vsock, "   14 '2'  : Answer in a channel notice."
              TU vsock, "   14 '3'  : Private message to the requesting user."
              TU vsock, "   14 '4'  : Private notice to the requesting user."
              TU vsock, EmptyLine
              TU vsock, "Type 'off', '1'-'4' or '0' to cancel."
          Case "11"
              SocketItem(vsock).CurrentQuestion = "ReactToWhois"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "211) React to !whois"
              TU vsock, "   14 If this setting is turned on, I'll react to 2!whois"
              TU vsock, "   14 messages in the channel."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "12"
              SocketItem(vsock).CurrentQuestion = "ReactToWhatis"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "212) React to " & CommandPrefix & "whatis"
              TU vsock, "   14 If this setting is turned on, I'll react to"
              TU vsock, "   14 2" & CommandPrefix & "whatis14 messages in the channel."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "13"
              SocketItem(vsock).CurrentQuestion = "AllowVoiceControl"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "213) Voice control"
              TU vsock, "   14 Allows voiced users in a channel to make me kick"
              TU vsock, "   14 somebody by typing '2" & CommandPrefix & "k <nick> <reason>14',"
              TU vsock, "   14 kickban with '2" & CommandPrefix & "kb14' and change the topic with '2" & CommandPrefix & "t14'."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "14"
              SocketItem(vsock).CurrentQuestion = "Secret"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "214) Secret mode"
              TU vsock, "   14 If this setting is turned on, I'll not provide"
              TU vsock, "   14 information about the channel over the botnet"
              TU vsock, "   14 and I'll not mention it in seen replies. I'll"
              TU vsock, "   14 only show it to users on the secret channel or"
              TU vsock, "   14 to owners using '2.seen14' on the party line."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "15"
              SocketItem(vsock).CurrentQuestion = "EnforceBans"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "215) Enforce Bans"
              TU vsock, "   14 If this setting is turned on, I'll kick unknown"
              TU vsock, "   14 users who are matching any bans in the channel."
              TU vsock, "   14 Channel ops and users matching ban exceptions"
              TU vsock, "   14 (mode +e) won't be kicked."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "16"
              SocketItem(vsock).CurrentQuestion = "FloodSettings"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "216) Flood settings"
              TU vsock, "   14 Please enter 3 numbers, seperated by spaces,"
              TU vsock, "   14 standing for the following flood settings:"
              TU vsock, "   14 1) Maximum lines per user"
              TU vsock, "   14 2) Maximum characters per user"
              TU vsock, "   14 3) Allowed repetitions per user"
              TU vsock, "   14"
              TU vsock, "   14 Example: A setting of '210 900 314' means that a"
              TU vsock, "   14 user may write up to2 1014 lines with up to2 90014"
              TU vsock, "   14 characters in total and is allowed to repeat"
              TU vsock, "   14 a sentence up to2 314 times in 10 seconds. Users"
              TU vsock, "   14 exceeding these limits will be kicked."
              TU vsock, EmptyLine
              TU vsock, "   14 (Owners and bots won't be kicked, of course!)"
              TU vsock, EmptyLine
              TU vsock, "   3 Current settings: " & UCase(GetChannelSetting(SocketItem(vsock).SetupChan, "FloodSettings", "10 900 3"))
              TU vsock, EmptyLine
              TU vsock, "Enter a flood setting, type 'x' to set the default, 'OFF'"
              TU vsock, "to disable the text flood protection or '0' to cancel."
          Case "17"
              SocketItem(vsock).CurrentQuestion = "BanMask"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt
              TU vsock, "217) Ban Mask"
              TU vsock, "   14 Please select the desired ban masks number"
              TU vsock, "   14 from the list below:"
              TU vsock, "   14"
              TU vsock, "   14  1 - *!~ident@host.domain"
              TU vsock, "   14  2 - *!*ident@host.domain"
              TU vsock, "   14  3 - *!*@host.domain"
              TU vsock, "   14  4 - *!*ident@*.domain"
              TU vsock, "   14  5 - *!*@*.domain"
              TU vsock, "   14  6 - nick!~ident@host.domain"
              TU vsock, "   14  7 - nick!*ident@host.domain"
              TU vsock, "   14  8 - nick!*@host.domain"
              TU vsock, "   14  9 - nick!*ident@*.domain"
              TU vsock, "   14 10 - nick!*@*.domain"
              TU vsock, "   14"
              TU vsock, EmptyLine
              TU vsock, "   3 Current settings: " & (CLng(GetChannelSetting(SocketItem(vsock).SetupChan, "BanMask", "3")) + 1)
              TU vsock, EmptyLine
              TU vsock, "Enter a flood setting, type 'x' to set the default or '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in ChanSetup!"
              TU vsock, "5*** Please enter a ChanSetup number, or leave by typing '0'."
              Exit Sub
        End Select
    Case "EnforceModes"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString SocketItem(vsock).SetupChan, "EnforceModes", HomeDir & "Channels.ini"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InStr("+-", Left(Line, 1)) = 0 Then
                TU vsock, "5*** You have to write '+' or '-' in front of the first flag."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              For u = 1 To Len(Line)
                If Mid(Line, u, 1) = " " Then Exit For
                If InStr(ServerChannelModes & "t+-", Mid(Line, u, 1)) = 0 Then
                  TU vsock, "5*** Sorry, there are invalid flags in your setting. Try again."
                  If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                  Exit Sub
                End If
              Next u
              WritePPString SocketItem(vsock).SetupChan, "EnforceModes", Line, HomeDir & "Channels.ini"
              ChNum = FindChan(SocketItem(vsock).SetupChan)
              If ChNum > 0 And (Channels(ChNum).GotOPs And Channels(ChNum).GotHOPs) Then
                ChangMode = ChangeMode(GetChannelSetting(Channels(ChNum).Name, "EnforceModes", ""), Channels(ChNum).Mode)
                If ChangMode <> "" Then
                  SendLine "mode " & SocketItem(vsock).SetupChan & " " & ChangMode, 2
                End If
              End If
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "DefaultTopic"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString SocketItem(vsock).SetupChan, "DefaultTopic", HomeDir & "Channels.ini"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If Len(Line) > ServerTopicLen Then
                TU vsock, "5*** Sorry, this topic is " & CStr(Len(Line) - ServerTopicLen) & " characters too long. Try again."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              WritePPString SocketItem(vsock).SetupChan, "DefaultTopic", Line, HomeDir & "Channels.ini"
              SendLine "topic " & SocketItem(vsock).SetupChan & " :" & Line, 2
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "NewbieGreeting"
        Select Case Param(Line, 1)
          Case "0"
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString SocketItem(vsock).SetupChan, "NewbieGreeting", HomeDir & "Channels.ini"
              DeletePPString SocketItem(vsock).SetupChan, "", HomeDir & "Newbie.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then Channels(u).NewbieGreeting = ""
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              WritePPString SocketItem(vsock).SetupChan, "NewbieGreeting", Line, HomeDir & "Channels.ini"
              DeletePPString SocketItem(vsock).SetupChan, "", HomeDir & "Newbie.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then Channels(u).NewbieGreeting = Line
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "AutoVoice"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "off", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).AutoVoice = 0
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "on", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).AutoVoice = 1
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "ext"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "ext", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).AutoVoice = 2
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on', 'ext' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'ext', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "DeopUnknownUsers"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "off", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).DeopUnknownUsers = 0
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "on", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).DeopUnknownUsers = 1
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "ext"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "ext", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).DeopUnknownUsers = 2
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on', 'ext' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'ext', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "ReactToSeen"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "0", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).ReactToSeen = 0
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "1", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).ReactToSeen = 1
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "1" To "4"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, Param(Line, 1), HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).ReactToSeen = CByte(Param(Line, 1))
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'off' or '1'-'4' after the setting number."
              Else
                TU vsock, "5*** Please enter 'off', '1'-'4' or '0'."
              End If
              Exit Sub
        End Select
    Case "BanMask"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).BanMask = 3
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If IsNumeric(Param(Line, 1)) = False Or CLng(Param(Line, 1)) < 1 Or CLng(Param(Line, 1)) > 10 Then
                TU vsock, "5*** Please enter a valid number, 'x' or '0'."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              End If
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, (CLng(Param(Line, 1)) - 1), HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).BanMask = (CLng(Param(Line, 1)) - 1)
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          End Select
    Case "FloodSettings"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "x"
              DeletePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).MaxLines = 8
                  Channels(u).MaxChars = 800
                  Channels(u).MaxRepeats = 3
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "off", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).MaxLines = 0
                  Channels(u).MaxChars = 0
                  Channels(u).MaxRepeats = 0
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If (ParamCount(Line) > 3) Or (IsNumeric(Param(Line, 1)) = False) Or (IsNumeric(Param(Line, 2)) = False) Or (IsNumeric(Param(Line, 3)) = False) Then
                TU vsock, "5*** Please enter 3 numbers, 'x', 'off', or '0'."
                If InstantSet Then SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                Exit Sub
              ElseIf ((CLng(Param(Line, 1)) > 0) And (CLng(Param(Line, 1)) < 3)) Or ((CLng(Param(Line, 2)) > 0) And (CLng(Param(Line, 2)) < 100)) Then
                TU vsock, "4*** Attention! Your settings are VERY strict! I'll kick for almost nothing now."
              End If
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, Line, HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Channels(u).MaxLines = CLng(Param(Line, 1))
                  Channels(u).MaxChars = CLng(Param(Line, 2))
                  Channels(u).MaxRepeats = CLng(Param(Line, 3))
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
        End Select
    Case "ProtectTopic", "ProtectFriends", "ColorKick", "CloneKick", "ReactToWhois", "AllowVoiceControl", "Secret", "ReactToWhatis", "EnforceBans"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "on", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Select Case SocketItem(vsock).CurrentQuestion
                    Case "ProtectFriends": Channels(u).ProtectFriends = True
                    Case "CloneKick": Channels(u).CloneKick = True
                    Case "ColorKick": Channels(u).ColorKick = True
                    Case "ReactToWhois": Channels(u).ReactToWhois = True
                    Case "AllowVoiceControl": Channels(u).AllowVoiceControl = True
                    Case "Secret": Channels(u).Secret = True
                    Case "ReactToWhatis": Channels(u).ReactToWhatis = True
                    Case "EnforceBans": Channels(u).EnforceBans = True
                  End Select
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
              WritePPString SocketItem(vsock).SetupChan, SocketItem(vsock).CurrentQuestion, "off", HomeDir & "Channels.ini"
              For u = 1 To ChanCount
                If LCase(Channels(u).Name) = LCase(SocketItem(vsock).SetupChan) Then
                  Select Case SocketItem(vsock).CurrentQuestion
                    Case "ProtectFriends": Channels(u).ProtectFriends = False
                    Case "CloneKick": Channels(u).CloneKick = False
                    Case "ColorKick": Channels(u).ColorKick = False
                    Case "ReactToWhois": Channels(u).ReactToWhois = False
                    Case "AllowVoiceControl": Channels(u).AllowVoiceControl = False
                    Case "Secret": Channels(u).Secret = False
                    Case "ReactToWhatis": Channels(u).ReactToWhatis = False
                    Case "EnforceBans": Channels(u).EnforceBans = False
                  End Select
                End If
              Next u
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case Else
      TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_ChanSetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowCSActions vsock
  End If
Exit Sub
ChanSErr:
  SendNote "AnGeL ChanSetup ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & SocketItem(vsock).CurrentQuestion & "] - ist aufgetreten."
  SendNote "AnGeL ChanSetup ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & SocketItem(vsock).CurrentQuestion & "] - ist aufgetreten."
  Err.Clear
End Sub
Public Sub PolSetup(vsock As Long, Line As String)
Dim u As Long, InstantSet As Boolean
On Error GoTo PolSErr
InstantSetIt2:
  Select Case SocketItem(vsock).CurrentQuestion
    Case "ChooseSetting"
        Select Case Param(Line, 1)
          Case "0"
              TU vsock, "10*** Saving Policies..."
              SetSockFlag vsock, SF_Status, SF_Status_Party
              SetAway vsock, ""
              AnGeLFiles.WritePolicies
          Case "1"
              SocketItem(vsock).CurrentQuestion = "UserManagement"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "21) User Management"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to add or remove users,"
              TU vsock, "  14         manage hosts, use of 'chattr'-command"
              TU vsock, "  14         and refreshing of the userfile."
              TU vsock, "  14  'off': Disables usermanagement-related"
              TU vsock, "  14         commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "2"
              SocketItem(vsock).CurrentQuestion = "DataBase"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "22) DataBase"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts using ODBC to connect"
              TU vsock, "  14         to local databases or SQL-Servers"
              TU vsock, "  14  'off': Disables ODBC usebility for scripts"
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "3"
              SocketItem(vsock).CurrentQuestion = "Resolve"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "23) Resolve"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to resolve Hostnames."
              TU vsock, "  14  'off': Disables Hostname resolving for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "4"
              SocketItem(vsock).CurrentQuestion = "Objects"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "24) CreateObject"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts using 'CreateObject' and"
              TU vsock, "  14         'GetObject' to create ActiveX objects."
              TU vsock, EmptyLine
              TU vsock, "  14         ActiveX is a powerfull windows component"
              TU vsock, "  14         but its as dangerous as powerfull,"
              TU vsock, "  14         because you can cotrol (and destroy)"
              TU vsock, "  14         a whole system with it"
              TU vsock, "  14  'off': Disables CreateOject and GetObject."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "5"
              SocketItem(vsock).CurrentQuestion = "FileOperations"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "25) FileOperations"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to modify Files in the,"
              TU vsock, "  14         FileArea-Directory"
              TU vsock, "  14  'off': Disables FileArea access for scripts"
              TU vsock, "  14         commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "6"
              SocketItem(vsock).CurrentQuestion = "ExecuteCommand"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "26) ExecuteCommand"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to execute partyline commands"
              TU vsock, "  14  'off': Disables partyline commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "7"
              SocketItem(vsock).CurrentQuestion = "TimeOperations"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "27) TimeOperations"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to read or modify"
              TU vsock, "  14         the systemtime"
              TU vsock, "  14  'off': Disables time-related"
              TU vsock, "  14         commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "8"
              SocketItem(vsock).CurrentQuestion = "SocketOperations"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "28) SocketOperations"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to communicate with the"
              TU vsock, "  14         Internet using Socket-Commands."
              TU vsock, "  14  'off': Disables Socket-commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "9"
              SocketItem(vsock).CurrentQuestion = "INIOperations"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "29) INI operations"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to modify ini-files"
              TU vsock, "  14  'off': Disables ini-file access for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "10"
              SocketItem(vsock).CurrentQuestion = "Hooks"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "210) Hooks"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to Hook events"
              TU vsock, "  14  'off': Disables event hooking."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "11"
              SocketItem(vsock).CurrentQuestion = "WMI"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "211) WMI"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to use the"
              TU vsock, "  14         Windows Management Instrument"
              TU vsock, "  14         (if available)."
              TU vsock, "  14  'off': Disables WMI for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "12"
              SocketItem(vsock).CurrentQuestion = "NativeFileSys"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "212) Native FileSys"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to En-/Decrypt, or (De)Compress."
              TU vsock, "  14         files on NTFS drives."
              TU vsock, "  14  'off': Disables native FileSystem access for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "13"
              SocketItem(vsock).CurrentQuestion = "BotInteractions"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "213) BotInteractions"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to interact the the bot"
              TU vsock, "  14         and the partyline."
              TU vsock, "  14  'off': Disables bot interaction-"
              TU vsock, "  14         commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "14"
              SocketItem(vsock).CurrentQuestion = "ChannelInteractions"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "214) ChannelInteractions"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to interact with channels"
              TU vsock, "  14         (i.e. kicking etc.)"
              TU vsock, "  14  'off': Disables interactions with channels for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "15"
              SocketItem(vsock).CurrentQuestion = "SessionChange"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "215) SessionChange"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to modify the"
              TU vsock, "  14         usersession-system (make valid/invalid)."
              TU vsock, "  14  'off': Disables usersession commands for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case "16"
              SocketItem(vsock).CurrentQuestion = "AddInModules"
              If Param(Line, 2) <> "" Then Line = Right(Line, Len(Line) - Len(Param(Line, 1)) - 1): InstantSet = True: GoTo InstantSetIt2
              TU vsock, "216) AddInModules"
              TU vsock, "  14  Here are the possible values for this setting:"
              TU vsock, "  14  'on' : Allows scripts to load AddIn-Modules"
              TU vsock, "  14  'off': Disables AddIn-Modules for scripts."
              TU vsock, EmptyLine
              TU vsock, "Type 'on', 'off' or '0' to cancel."
          Case Else
              If Left(Trim(Line), 1) = "." Then TU vsock, "5*** You can't use bot commands in POLICYSetup!"
              TU vsock, "5*** Please enter a POLICYSetup number, or leave by typing '0'."
              Exit Sub
        End Select
    Case "UserManagement"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "UserManagement", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "UserManagement", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "DataBase"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "DataBase", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "DataBase", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "Resolve"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "Resolves", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "Resolves", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "Objects"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "Objects", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "Objects", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "FileOperations"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "FileOperations", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "FileOperations", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "ExecuteCommand"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "ExecuteCommand", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "ExecuteCommand", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "TimeOperations"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "TimeOperations", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "TimeOperations", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "SocketOperations"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "SocketOperations", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "SocketOperations", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "INIOperations"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "INIOperations", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "INIOperations", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "Hooks"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "NotificationHooks", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "NotificationHooks", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "WMI"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "WMI", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "WMI", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "BotInteractions"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "BotInteractions", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "BotInteractions", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "ChannelInteractions"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "ChannelInteractions", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "ChannelInteractions", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "SessionChange"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "SessionChange", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "SessionChange", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "AddInModules"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "AddInModules", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "AddInModules", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case "NativeFileSys"
        Select Case LCase(Param(Line, 1))
          Case "0": SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "on"
             AnGeLFiles.AllowCommand "NativeFS", True
             SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case "off"
             AnGeLFiles.AllowCommand "NativeFS", False
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
          Case Else
              If InstantSet Then
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                TU vsock, "5*** Please enter 'on' or 'off' after the setting number."
              Else
                TU vsock, "5*** Please enter 'on', 'off' or '0'."
              End If
              Exit Sub
        End Select
    Case Else
      TU vsock, "4*** ERROR. Invalid Question: " & SocketItem(vsock).CurrentQuestion
  End Select
  If GetSockFlag(vsock, SF_Status) = SF_Status_POLSetup And SocketItem(vsock).CurrentQuestion = "ChooseSetting" Then
    ShowPolActions vsock
  End If
Exit Sub
PolSErr:
  SendNote "AnGeL ChanSetup ERROR", "SailorCM", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & SocketItem(vsock).CurrentQuestion & "] - ist aufgetreten."
  Err.Clear
End Sub
Public Sub ListChannelSettings(vsock As Long, Chan As String) ' : AddStack "Setups_ListChannelSettings(" & vsock & ", " & Chan & ")"
Dim u As Long, CheckHost As String, CheckDesc As String, ToldOneHostMask As Boolean
  TU vsock, " 2 1) Enforce Modes14...: " & GetChannelSetting(Chan, "EnforceModes", "<none>")
  ToldOneHostMask = False
  CheckHost = ""
  CheckDesc = GetChannelSetting(Chan, "DefaultTopic", "<none>")
  For u = 1 To Len(CheckDesc)
    If Mid(CheckDesc, u, 1) <> " " Then
      CheckHost = CheckHost + Mid(CheckDesc, u, 1)
    Else
      If InStr(u + 1, CheckDesc & " ", " ") - u > 40 - Len(CheckHost) Then
        If Not ToldOneHostMask Then
          TU vsock, " 2 2) Default Topic14...: " & CheckHost
          ToldOneHostMask = True
        Else
          TU vsock, " 2                     " & CheckHost
        End If
        CheckHost = ""
      Else
        CheckHost = CheckHost & " "
      End If
    End If
  Next u
  If Not ToldOneHostMask Then
    TU vsock, " 2 2) Default Topic14...: " & CheckHost
    ToldOneHostMask = True
  Else
    TU vsock, " 2                     " & CheckHost
  End If
  ToldOneHostMask = False
  
  CheckHost = ""
  CheckDesc = GetChannelSetting(Chan, "NewbieGreeting", "<no message set>")
  For u = 1 To Len(CheckDesc)
    If Mid(CheckDesc, u, 1) <> " " Then
      CheckHost = CheckHost + Mid(CheckDesc, u, 1)
    Else
      If InStr(u + 1, CheckDesc & " ", " ") - u > 40 - Len(CheckHost) Then
        If Not ToldOneHostMask Then
          TU vsock, " 2 3) Greet Newcomers14.: " & CheckHost
          ToldOneHostMask = True
        Else
          TU vsock, " 2                     " & CheckHost
        End If
        CheckHost = ""
      Else
        CheckHost = CheckHost & " "
      End If
    End If
  Next u
  If Not ToldOneHostMask Then
    TU vsock, " 2 3) Greet Newcomers14.: " & CheckHost
    ToldOneHostMask = True
  Else
    TU vsock, " 2                     " & CheckHost
  End If
  TU vsock, " 2 4) Protect topic14...: " & IIf(LCase(GetChannelSetting(Chan, "ProtectTopic", "off")) = "on", "3ON ", "4OFF") & "  2 5) Protect friends14...: " & IIf(LCase(GetChannelSetting(Chan, "ProtectFriends", "off")) = "on", "3ON ", "4OFF")
  TU vsock, " 2 6) Color kick14......: " & IIf(LCase(GetChannelSetting(Chan, "ColorKick", "off")) = "on", "3ON ", "4OFF") & "  2 7) Clone kick14........: " & IIf(LCase(GetChannelSetting(Chan, "CloneKick", "off")) = "on", "3ON ", "4OFF")
  Select Case LCase(GetChannelSetting(Chan, "AutoVoice", "off"))
    Case "off": CheckDesc = "4OFF"
    Case "on":  CheckDesc = "3ON "
    Case "ext": CheckDesc = "10EXT"
  End Select
  Select Case LCase(GetChannelSetting(Chan, "DeopUnknownUsers", "off"))
    Case "off": CheckHost = "4OFF"
    Case "on":  CheckHost = "3ON "
    Case "ext": CheckHost = "10EXT"
  End Select
  TU vsock, " 2 8) Autovoice if +m14.: " & CheckDesc & "  2 9) Deop unknown users: " & CheckHost
  TU vsock, " 210) React to " & CommandPrefix & "seen14..: " & IIf(LCase(GetChannelSetting(Chan, "ReactToSeen", "1")) = "0", "4OFF", "10" & Spaces2(3, UCase(GetChannelSetting(Chan, "ReactToSeen", "1"))) & "") & "  211) React to !whois14...: " & IIf(LCase(GetChannelSetting(Chan, "ReactToWhois", "on")) = "on", "3ON ", "4OFF")
  TU vsock, " 212) React to " & CommandPrefix & "whatis: " & IIf(LCase(GetChannelSetting(Chan, "ReactToWhatis", "on")) = "on", "3ON ", "4OFF") & "  213) Voice control14.....: " & IIf(LCase(GetChannelSetting(Chan, "AllowVoiceControl", "on")) = "on", "3ON ", "4OFF")
  TU vsock, " 214) Secret mode14.....: " & IIf(LCase(GetChannelSetting(Chan, "Secret", "off")) = "on", "3ON ", "4OFF") & "  215) Enforce bans14......: " & IIf(LCase(GetChannelSetting(Chan, "EnforceBans", "on")) = "on", "3ON ", "4OFF")
  Select Case LCase(GetChannelSetting(Chan, "FloodSettings", "10 900 3"))
    Case "off": CheckHost = "4OFF"
    Case "10 900 3": CheckHost = "3DEF"
    Case Else: CheckHost = "10EXT"
  End Select
  TU vsock, " 216) Flood settings14..: " & CheckHost & "  217) Ban Mask..........:10 " & (CLng(GetChannelSetting(Chan, "BanMask", "3")) + 1)
  TU vsock, EmptyLine
End Sub
Sub ListKISettings(vsock As Long) ' : AddStack "Setups_ListKISettings(" & vsock & ")"
  TU vsock, " 2 1) First Name14....: " & MakeLength(IIf((KIFName <> ""), KIFName, "4NOT SET"), 10) & "  2 2) Last Name14.....: " & MakeLength(IIf((KILName <> ""), KILName, "4NOT SET"), 10)
  TU vsock, " 2 3) Gender14........: " & MakeLength(IIf(KIGender <> "", KIGender, "4NOT SET"), 10) & "  2 4) Age14...........: " & MakeLength(IIf(CStr(KIAge) <> "", CStr(KIAge), "4NOT SET"), 10)
  TU vsock, " 2 5) City/Town14.....: " & MakeLength(IIf(KICity <> "", KICity, "4NOT SET"), 10) & "  2 6) Country14.......: " & MakeLength(IIf(KICountry <> "", KICountry, "4NOT SET"), 10)
  TU vsock, EmptyLine
End Sub
Public Sub ShowCSActions(vsock As Long) ' : AddStack "Setups_ShowCSActions(" & vsock & ")"
  ListChannelSettings vsock, SocketItem(vsock).SetupChan
  TU vsock, "Enter a setting number or enter '0' to leave the setup."
  TU vsock, "You can use shortcuts like '10 on' or '1 +tn-i'."
  TU vsock, EmptyLine
End Sub
Public Sub ListPersonalSettings(vsock As Long) ' : AddStack "Routines_ListPersonalSettings(" & vsock & ")"
Dim u As Long, CheckHost As String, CheckDesc As String, ToldOneHostMask As Boolean
  TU vsock, " 2 1) Show Colors14....: " & IIf(GetSockFlag(vsock, SF_Colors) = SF_YES, "3ON ", "4OFF") & "  2 2) Auto-WHO on join14..: " & IIf(GetSockFlag(vsock, SF_AutoWHO) = SF_YES, "3ON ", "4OFF")
  'TU vsock, " 2 3) Console Modes14..: " & SocketItem(vsock).Console
  TU vsock, EmptyLine
End Sub



