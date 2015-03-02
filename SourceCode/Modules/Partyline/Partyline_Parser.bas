Attribute VB_Name = "Partyline_Parser"
',-======================- ==-- -  -
'|   AnGeL - Partyline - Parser
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Sub PartyLineLogon(vsock As Long, Nick As String, FirstLogon As Boolean) ' : AddStack "Routines_PartyLineLogon(" & vsock & ", " & Nick & ", " & FirstLogon & ")"
Dim NoteCount As Long, VersionString As String, UserLevelString As String
Dim ScNum As Long
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    VersionString = Right(BotVersionEx, Len(BotVersionEx) - 1)
    TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'1,1              0,1" & String(Len(VersionString), " ") & "1,1             2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
    TU vsock, "11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'1,1  0,1AnGeL Party Line Version " & VersionString & "1,1  2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
    TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'1,1              0,1" & String(Len(VersionString), " ") & "1,1             2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
    TU vsock, EmptyLine
    TU vsock, MakeMsg(MSG_ConnectedTo, BotNetNick) & " " & CpString
  Else
    TU vsock, MakeMsg(MSG_ConnectedTo, BotNetNick) & " " & BotVersion + IIf(ServerNetwork <> "", "+" & ServerNetwork, "") & "  " & CpString
  End If
  TU vsock, EmptyLine
  TU vsock, "News: INVITE/EXCEPT support, enhanced scripting, .whatsnew"
  TU vsock, "  Visit the new AnGeL homepage: 2http://www.angel-bot.de"
  TU vsock, EmptyLine
  ReadSockFlags vsock
  ShowMOTD vsock
  If FirstLogon Then
    TU vsock, MakeMsg(MSG_FirstLogin, Nick, BotNetNick)
    If LCase(IdentCommand) <> "ident" Then
      TU vsock, MakeMsg(MSG_SecurityInfo, IdentCommand)
    End If
    TU vsock, EmptyLine
  End If
  TUEx vsock, SF_ExtraHelp, MakeMsg(MSG_ShortIntro)
  TUEx vsock, SF_ExtraHelp, EmptyLine
  NoteCount = NotesCount(Nick)
  If NoteCount > 0 Then
    If NoteCount = 1 Then
      TU vsock, MakeMsg(MSG_PLNote, Nick)
    Else
      TU vsock, MakeMsg(MSG_PLNotes, CStr(NoteCount), Nick)
    End If
  End If
  'Set user to SF_Status_Party and tell the botnet about the join
  SetSockFlag vsock, SF_Status, SF_Status_Party
  SetSockFlag vsock, SF_LocalVisibleUser, SF_YES
  SocketItem(vsock).NumOfServerEvents = 0
  SocketItem(vsock).OrderSign = FindFreeOrderSign
  ToBotNet 0, "j " & BotNetNick & " " & Nick & " " & LongToBase64(SocketItem(vsock).PLChannel) & " " & GetLevelSign(SocketItem(vsock).Flags) + SocketItem(vsock).OrderSign & " " & Mask(SocketItem(vsock).Hostmask, 10)
  'Check script Hooks
  HaltDefault = False
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.PLJoin Then
      RunScriptX ScNum, "PLJoin", vsock, Nick, SocketItem(vsock).Flags
    End If
  Next ScNum
  If HaltDefault = False Then
    UserLevelString = LevelString(SocketItem(vsock).Flags, AllFlags(Nick))
    If SocketItem(vsock).PLChannel = 0 Then
      TU vsock, MakeMsg(MSG_PLJoin, Nick, UserLevelString, Strip(UserLevelString))
      SpreadMessageEx vsock, SocketItem(vsock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLJoin, Nick, UserLevelString, Strip(UserLevelString))
    Else
      TU vsock, MakeMsg(MSG_PLJoinChan, Nick, "#" & CStr(SocketItem(vsock).PLChannel), UserLevelString, Strip(UserLevelString))
      SpreadMessageEx vsock, SocketItem(vsock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLJoinChan, Nick, "#" & CStr(SocketItem(vsock).PLChannel), UserLevelString, Strip(UserLevelString))
    End If
    TU vsock, EmptyLine
  End If
  'Write party line seen entry for this user
  WriteSeenEntry Nick, "", Now, "*mine*", "*partyline*", Mask(SocketItem(vsock).Hostmask, 10)
End Sub

Public Sub FailCommand(ByVal vsock As Long, SpreadTo As String, Line As String) ' : AddStack "GUI_FailCommand(" & vsock & ", " & SpreadTo & ", " & Line & ")"
  SpreadFlagMessage vsock, SpreadTo, MakeMsg(MSG_PLNickFailed, SocketItem(vsock).RegNick, Line)
End Sub

Public Sub SucceedCommand(ByVal vsock As Long, SpreadTo As String, Line As String) ' : AddStack "GUI_SucceedCommand(" & vsock & ", " & SpreadTo & ", " & Line & ")"
  SpreadFlagMessage vsock, SpreadTo, MakeMsg(MSG_PLNickDid, SocketItem(vsock).RegNick, Line)
End Sub

Sub PartySort(vsock As Long, SockNum As Long, ByVal InLine As String) ' : AddStack "PartyHandler_PartySort(" & vsock & ", " & SockNum & ", " & InLine & ")"
Dim chPos As Long, SPos As Long, Part As String, u As Long
Dim Line As String, StringToEcho As String
  Output InLine
  Line = SocketItem(vsock).InputBuffer
  u = 1
    If Left(InLine, 4) = "100 " And SocketItem(vsock).RegNick = "" Then
        InLine = ""
        SetSockFlag vsock, SF_Echo, SF_NO
        SetSockFlag vsock, SF_DCC, SF_YES
        SetSockFlag vsock, SF_LF_ONLY, SF_YES
        SetSockFlag vsock, SF_Telnet, SF_NO
        TU vsock, "101 " & BotNetNick
        Exit Sub
    End If
  Do
    Select Case Asc(Mid(InLine, u, 1))
      Case 27
        Select Case Mid(InLine, u + 1, 2)
          Case "[A", "[B", "[C", "[D"
            InLine = ""
            StringToEcho = ""
        End Select
      Case 8, 127 ' BackSpace?
        If Len(Line) > 0 Then
          Line = Left(Line, Len(Line) - 1)
          StringToEcho = Chr(8) & " " & Chr(8) ' StringToEcho & Mid(Line, u, 1)
        End If
      Case 255 'IAC
        Select Case Mid(InLine, u + 1, 2)
          Case T_DO & T_SPGA
            Output "Received DO SPGA -> Sent WILL SPGA" & vbCrLf
            SendTCP vsock, T_IAC & T_WILL & T_SPGA: u = u + 2
          Case T_DO & T_ECHO
            Output "Received DO ECHO -> Sent WILL ECHO -> echo on" & vbCrLf
            SetSockFlag vsock, SF_Echo, SF_YES
            SendTCP vsock, T_IAC & T_WILL & T_ECHO: u = u + 2
          Case T_DONT & T_SPGA
            Output "Received DONT SPGA -> Sent WONT SPGA" & vbCrLf
            SendTCP vsock, T_IAC & T_WONT & T_SPGA: u = u + 2
          Case T_DONT & T_ECHO
            Output "Received DONT ECHO -> Sent WONT ECHO -> echo off" & vbCrLf
            SetSockFlag vsock, SF_Echo, SF_NO
            SendTCP vsock, T_IAC & T_WONT & T_ECHO: u = u + 2
          Case T_WILL & T_SPGA
            Output "Received WILL SPGA" & vbCrLf
            u = u + 2
          Case T_WILL & T_ECHO
            Output "Received WILL ECHO" & vbCrLf
            u = u + 2

            Output "Received WONT SPGA" & vbCrLf
            u = u + 2
          Case T_WONT & T_ECHO
            Output "Received WONT ECHO" & vbCrLf
            u = u + 2
          Case Else
            Output "UNKNOWN! -> " & Mid(InLine, u + 1, 2)
            'Trace Asc(Mid(InLine, u + 1, 1)), Asc(Mid(InLine, u + 2, 1))
            u = u + 2
        End Select
      Case Else: Line = Line + Mid(InLine, u, 1): StringToEcho = StringToEcho + Mid(InLine, u, 1)
    End Select
    u = u + 1
  Loop Until u > Len(InLine)
  If GetSockFlag(vsock, SF_Echo) = SF_YES Then SendTCP vsock, StringToEcho
  SPos = 1
  Do
    chPos = InStr(SPos, Line, Chr(13))
    If chPos = 0 Then chPos = InStr(SPos, Line, Chr(10))
    If chPos = 0 Then Exit Do
    Part = Mid(Line, SPos, chPos - SPos)
    Party vsock, SockNum, Part
    If Not IsValidSocket(vsock) Then Exit Sub
    SPos = chPos + 1
    If Mid(Line, SPos, 1) = Chr(10) Then SPos = SPos + 1
  Loop
  Part = Mid(Line, SPos, Len(Line) - SPos + 1)
  SocketItem(vsock).InputBuffer = Part
  If Len(SocketItem(vsock).InputBuffer) > 5000 Then
    If GetSockFlag(vsock, SF_LocalVisibleUser) = SF_YES Then
      SpreadMessage vsock, SocketItem(vsock).PLChannel, "3*** " & SocketItem(vsock).RegNick & " was killed for character flooding"
      ToBotNet 0, "pt " & BotNetNick & " " & SocketItem(vsock).RegNick & " " & SocketItem(vsock).OrderSign & " Killed for character flooding"
      SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Closed connection from " & SocketItem(vsock).Hostmask & " (" & SocketItem(vsock).RegNick & "): Character flood"
    Else
      SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Telnet: Closed connection from " & SocketItem(vsock).Hostmask & ": Character flood"
    End If
    RemoveSocket vsock, 0, "", True
  End If
End Sub

Sub Party(ByVal vsock As Long, ByVal SockNum As Long, Line As String)
  Dim Nick As String, Flags As String, ChNum As Long, Messg As String, RegNick As String
  Dim u As Long, u2 As Long, ChanFlags As String, UsNum As Long, u4 As Single
  Dim FoundOne As Boolean, oppedone As Boolean, CheckURL As String, CheckDesc As String
  Dim ChangMode As String, HostNum As Long, t As Long, Delivered As Long, NoteCount As Long
  Dim OtherUserFlags As String, FileNum As Integer, RealNick As String, u3 As Long
  Dim Rest As String, CheckLine As String, ToldOneHostMask As Boolean
  Dim SearchIn As String, TheFlag As String, Result As Long
  Dim TargetNick As String, ScNum As Long, UserNum As Long, CheckHost As String
  Nick = SocketItem(vsock).RegNick
  Flags = SocketItem(vsock).Flags
  UserNum = SocketItem(vsock).UserNum
  If (SocketItem(vsock).AwayMessage = "") And (Line <> "") Then SocketItem(vsock).LastEvent = Now
  Select Case GetSockFlag(vsock, SF_Status)
    Case SF_Status_UserGetName
        SocketItem(vsock).UserNum = GetUserNum(Line)
        UserNum = SocketItem(vsock).UserNum
        If SocketItem(vsock).UserNum = 0 Then
          TU vsock, MakeMsg(ERR_Login_Unknown)
          RemoveSocket vsock, 0, "", False
          Exit Sub
        End If
        SocketItem(vsock).RegNick = BotUsers(SocketItem(vsock).UserNum).Name
        SocketItem(vsock).Flags = BotUsers(SocketItem(vsock).UserNum).Flags
        If MatchFlags(SocketItem(vsock).Flags, "+b") Then
          TU vsock, "ERROR Please use the botnet port! (" & CStr(BotnetPort) & ")."
          RemoveSocket vsock, 0, "", False
          Exit Sub
        End If
        If MatchFlags(SocketItem(vsock).Flags, "-p") Then
          TU vsock, MakeMsg(ERR_Login_NoChat)
          SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLTelNetNoAcc, SocketItem(vsock).RegNick)
          RemoveSocket vsock, 0, "", False
          Exit Sub
        End If
        'SetSockFlag vsock, SF_Colors, SF_NO
        Nick = SocketItem(vsock).RegNick
        SetSockFlag vsock, SF_Colors, GetUserData(SocketItem(vsock).UserNum, "colors", SF_NO)
        Flags = SocketItem(vsock).Flags
        If BotUsers(SocketItem(vsock).UserNum).Password <> "" Then
          SetSockFlag vsock, SF_Echo, SF_NO
          TU vsock, MakeMsg(MSG_EnterPWD, SocketItem(vsock).RegNick)
          SetSockFlag vsock, SF_Status, SF_Status_UserGetPass
          SpreadFlagMessageEx u, "+m", SF_Local_JP, MakeMsg(MSG_PLTelNetOpened, SocketItem(vsock).RegNick)
        Else
          TU vsock, MakeMsg(MSG_ChoosePWD, SocketItem(vsock).RegNick)
          SetSockFlag vsock, SF_Status, SF_Status_UserChoosePass
          SpreadFlagMessageEx u, "+m", SF_Local_JP, MakeMsg(MSG_PLTelNetFirst, SocketItem(vsock).RegNick)
        End If
    Case SF_Status_UserChoosePass
        If Len(Line) < 6 Then TU vsock, MakeMsg(ERR_Pass_TooShort): Exit Sub
        If InStr(Line, " ") > 0 Then TU vsock, MakeMsg(ERR_Pass_NoSpaces): Exit Sub
        If WeakPass(Line, Nick) Then TU vsock, MakeMsg(ERR_Pass_TooWeak): Exit Sub
        BotUsers(UserNum).Password = EncryptIt(Line)
        BotUsers(UserNum).ValidSession = True
        TU vsock, MakeMsg(MSG_ThanksPWD, Line, Nick)
        TU vsock, EmptyLine
        RTU vsock, MakeMsg(MSG_CanYouSeeColors)
        SetSockFlag vsock, SF_Status, SF_Status_UserChooseColors
    Case SF_Status_UserChooseColors
        Select Case LCase(Line)
          Case "yes", "y", "ja", "j", "oui", "si", "sì", "sí", "sim", "hai", "1"
            Line = "yes"
          Case "no", "n", "nein", "non", "ningún", "ningun", "nenhum", "nessun", "0"
            Line = "no"
          Case Else
            Line = ""
        End Select
        If Line <> "" Then
          SetSockFlag vsock, SF_Colors, IIf(LCase(Line) = "yes", SF_YES, SF_NO)
          SetUserData UserNum, "colors", IIf(LCase(Line) = "yes", SF_YES, SF_NO)
          PartyLineLogon vsock, Nick, True
        Else
          TU vsock, MakeMsg(MSG_CanYouSeeColTwo)
        End If
    Case SF_Status_UserGetPass
        If Len(Line) = 0 Then Exit Sub
        If BotUsers(UserNum).Password <> EncryptIt(Line) Then
          TU vsock, MakeMsg(ERR_Login_WrongPass, Nick)
          If GetSockFlag(vsock, SF_DCC) = SF_NO Then
            SpreadFlagMessage vsock, "+m", MakeMsg(MSG_PLTelNetWrongPW, Nick)
          Else
            SpreadFlagMessage vsock, "+m", MakeMsg(MSG_PLDCCWrongPW, Nick)
          End If
          RemoveSocket vsock, 0, "", True
          Exit Sub
        End If
        BotUsers(UserNum).ValidSession = True
        If GetSockFlag(vsock, SF_DCC) = SF_NO Then SetSockFlag vsock, SF_Echo, SF_YES
        If GetUserData(UserNum, "colors", "") = "" Then
          TU vsock, "Thank you."
          TU vsock, EmptyLine
          RTU vsock, MakeMsg(MSG_CanYouSeeColors)
          SetSockFlag vsock, SF_Status, SF_Status_UserChooseColors
        Else
          PartyLineLogon vsock, Nick, False
        End If
    Case SF_Status_ScriptUser
        If (SocketItem(vsock).CurrentQuestion <> "") Then
          For ScNum = 1 To ScriptCount
            If Scripts(ScNum).Name = SocketItem(vsock).SetupChan Then
              RunScriptX ScNum, SocketItem(vsock).CurrentQuestion, vsock, Nick, SocketItem(vsock).Flags, Line
            End If
          Next ScNum
        End If
    Case SF_Status_Party
      If Left(Line, 1) = "." And Mid(Line, 2, 1) <> "." Then
        'Check commands
        If MatchCommand(Mid(Param(Line, 1), 2)) <> "" Then Line = CStr("." & MatchCommand(Mid(Param(Line, 1), 2)) & " " & GetRest(Line, 2))
        Result = CheckCommand(vsock, Mid(Param(Line, 1), 2))
        If Result = CC_NoCommand Then
          TU vsock, MakeMsg(MSG_PL_LookHelp, Line)
          SpreadFlagMessage vsock, "+m", MakeMsg(MSG_PLNickTried, Nick, Line)
          Exit Sub
        ElseIf Result = CC_NotAllowed Then
          If SocketItem(vsock).IRCNick <> "²*SCRIPT*²" Then
            TU vsock, MakeMsg(ERR_NotAllowed, Nick)
            SpreadFlagMessage vsock, "+m", MakeMsg(MSG_PLNickFailed, Nick, Line)
            Exit Sub
          End If
        Else
          If Commands(Result).SpreadTo <> "" Then
            SpreadFlagMessage vsock, Commands(Result).SpreadTo, MakeMsg(MSG_PLNickDid, Nick, Line)
          End If
          If SocketItem(vsock).IRCNick <> "²*SCRIPT*²" Then
            'Check script Hooks
            HaltDefault = False
            For ScNum = 1 To ScriptCount
              If Scripts(ScNum).Hooks.Commands Then
                RunScriptX ScNum, "Commands", vsock, Nick, SocketItem(vsock).Flags, Line
              End If
            Next ScNum
            If HaltDefault Then Exit Sub
          End If
        End If
        
        'Execute them
        Select Case LCase(Param(Line, 1))
          'User commands
          Case ".dccstat", ".ds"
              ShowDCStat vsock
          Case ".who"
              If Param(Line, 2) <> "" Then
                If GetNextBot(Param(Line, 2)) = 0 Then TU vsock, MakeMsg(ERR_BotNotFound, Param(Line, 2)): Exit Sub
                SendToBot Param(Line, 2), "w " & CStr(vsock) & ":" & Nick & "@" & BotNetNick & " " & Param(Line, 2) & " A"
              Else
                ListBotUsers vsock
              End If
          Case ".whom"
              ListBotNetUsers vsock
          Case ".whois"
              Whois vsock, Line, False
          Case ".colors", ".color", ".colour"
              Colors vsock, Line
          Case ".newpass"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".newpass <password>"): Exit Sub
              If Len(Param(Line, 2)) > 5 Then
                If WeakPass(Param(Line, 2), Nick) Then TU vsock, "5*** " & MakeMsg(ERR_Pass_TooWeak): Exit Sub
                BotUsers(UserNum).Password = EncryptIt(Param(Line, 2))
                TU vsock, "3*** Password changed to '" & Param(Line, 2) & "'."
                SpreadFlagMessage vsock, "+m", MakeMsg(MSG_PLNickDid, Nick, ".newpass ...")
              Else
                TU vsock, "5*** " & MakeMsg(ERR_Pass_TooShort)
              End If
          Case ".chat"
              If Trim(Param(Line, 2)) = "" Then
                SocketItem(vsock).PLChannel = 0
              Else
                On Local Error Resume Next
                SocketItem(vsock).PLChannel = CLng(Param(Line, 2))
                If Err.Number > 0 Then
                  TU vsock, "5*** Sorry, invalid channel number."
                Else
                  SpreadMessage 0, SocketItem(vsock).PLChannel, MakeMsg(MSG_PLJoinChan, SocketItem(vsock).RegNick, "#" & CStr(SocketItem(vsock).PLChannel))
                End If
                ToBotNet 0, "j " & BotNetNick & " " & SocketItem(u).RegNick & " " & LongToBase64(SocketItem(u).PLChannel) & " " & GetLevelSign(Flags) + SocketItem(u).OrderSign & " " & Right(Mask(SocketItem(u).Hostmask, 2), Len(Mask(SocketItem(u).Hostmask, 2)) - 4)
                On Error GoTo 0
                If Err.Number > 0 Then Err.Clear
              End If
          Case ".nick"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".nick <new nick>"): Exit Sub
              If Len(Param(Line, 2)) > ServerNickLen Then TU vsock, MakeMsg(ERR_Nick_TooLong, CStr(ServerNickLen)): Exit Sub
              If IsValidNick(Param(Line, 2)) = False Then TU vsock, MakeMsg(ERR_Nick_Erroneous, Param(Line, 2)): Exit Sub
              If UserExist(Param(Line, 2)) And Not (LCase(Param(Line, 2)) = LCase(Nick) And Param(Line, 2) <> Nick) Then TU vsock, MakeMsg(ERR_Nick_InUse, Param(Line, 2)): Exit Sub
              ChangeNick Nick, Param(Line, 2)
              SharingSpreadMessage Nick, ".chnick " & Nick & " " & Param(Line, 2)
          Case ".me"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".me <action>"): Exit Sub
              ToBotNet 0, "a " & Nick & "@" & BotNetNick & " A " & GetRest(Line, 2)
              SpreadMessage vsock, SocketItem(vsock).PLChannel, MakeMsg(MSG_PLAct, Nick, GetRest(Line, 2))
          Case ".msg"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".msg <nick> <message>"): Exit Sub
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If LCase(SocketItem(u).RegNick) = LCase(Param(Line, 2)) Then
                    If GetSockFlag(u, SF_LocalVisibleUser) = SF_YES Then TU u, "14*" & Nick & "* " & Right(Line, Len(Line) - Len(Param(Line, 2)) - 6): FoundOne = True
                  End If
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, there's no user called '" & Param(Line, 2) & "' on the party line.": Exit Sub
          Case ".away"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If (SocketItem(vsock).AwayMessage = "") And (Param(Line, 2) = "") Then TU vsock, MakeMsg(ERR_CommandUsage, ".away <reason>"): Exit Sub
              SetAway vsock, GetRest(Line, 2)
          Case ".back"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If (SocketItem(vsock).AwayMessage = "") Then TU vsock, "5*** You are not marked as away!": Exit Sub
              SetAway vsock, ""
          Case ".files"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Not FileAreaEnabled Then TU vsock, "5*** Sorry, the file area is disabled!": Exit Sub
              If MatchFlags(Flags, "-i") Then
                TU vsock, "5*** Sorry, you don't have file area access (+i needed).": Exit Sub
              Else
                HaltDefault = False
                For ScNum = 1 To ScriptCount
                  If Scripts(ScNum).Hooks.fa_userjoin Then
                    RunScriptX ScNum, "fa_userjoin", vsock, SocketItem(vsock).RegNick
                  End If
                Next ScNum
                If HaltDefault = True Then Exit Sub
                SetAway vsock, "File area"
                SetSockFlag vsock, SF_Status, SF_Status_FileArea
                If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                  TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1           2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                  TU vsock, "11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1  File area  2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                  TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1           2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                End If
                SocketItem(vsock).FileAreaDir = "\"
                TU vsock, EmptyLine
                TU vsock, "Welcome to the file area! If you don't know what"
                TU vsock, "to do, type 'help' to get a list of all commands."
                TU vsock, EmptyLine
                TU vsock, "2*** Your current directory is: " & SocketItem(vsock).FileAreaDir
                TU vsock, EmptyLine
              End If
          Case ".invite"
              If Not (IsValidChannel(LCase(Left(Param(Line, 3), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".invite <nick> <[" & ServerChannelPrefixes & "]channel>"): Exit Sub
              If MatchFlags(GetUserChanFlags(Nick, Param(Line, 3)), "-o") Then TU vsock, "5*** Sorry, you don't have op rights for this channel.": FailCommand vsock, "+m", Line: Exit Sub
              SucceedCommand vsock, "+m", Line
              SendLine "invite " & Param(Line, 2) & " " & Param(Line, 3), 1
              TU vsock, "3*** Trying to invite " & Param(Line, 2) & " to " & Param(Line, 3) & "..."
              SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 2
          Case ".info"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              Rest = GetUserData(UserNum, "info", "")
              If Param(Line, 2) = "" And Rest = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".info <your info line>"): Exit Sub
              If Param(Line, 2) = "" Then
                TU vsock, "3*** Your current info line: " & Rest
                TU vsock, "3*** Type '.info <new info line>' to change it."
              Else
                If Len(GetRest(Line, 2)) > 440 Then TU vsock, "5*** Sorry, this info line is " & CStr(Len(GetRest(Line, 2)) - 440) & " characters too long.": Exit Sub
                SetUserData UserNum, "info", GetRest(Line, 2)
                TU vsock, "3*** Your info line was set."
              End If
          Case ".chinfo"
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".chinfo <nick> <info line>"): Exit Sub
              UsNum = GetUserNum(Param(Line, 2))
              If UsNum = 0 Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
              If Len(GetRest(Line, 3)) > 440 Then TU vsock, "5*** Sorry, this info line is " & CStr(Len(GetRest(Line, 3)) - 440) & " characters too long.": Exit Sub
              SetUserData UsNum, "info", GetRest(Line, 3)
              TU vsock, "3*** The info line of " & BotUsers(UsNum).Name & " was set."
          Case ".channel", ".ch"
              Channel vsock, Line
          Case ".chanbans", ".cb"
              ChanBans vsock, Line
          Case ".seen"
              seen vsock, Line
          Case ".kick", ".k"
              Kick vsock, Line
          Case ".kickban", ".kb"
              KickBan vsock, Line
          Case ".wset"
              If MatchFlags(Flags, "-w") Then
                TU vsock, "5*** You need the +w flag to use '.wset'.": Exit Sub
              Else
                If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".wset <item> (description)"): Exit Sub
                If Left(Param(Line, 2), 1) = """" Then
                  If InStr(InStr(Line, """") + 1, Line, """") = 0 Then TU vsock, "5*** If you want to add items with spaces, use '.wset ""<Item>"" (description)'": Exit Sub
                  Rest = Mid(Line, InStr(Line, """") + 1, InStr(InStr(Line, """") + 1, Line, """") - InStr(Line, """") - 1)
                  SearchIn = Trim(Right(Line, Len(Line) - InStr(InStr(Line, """") + 1, Line, """")))
                  If SearchIn <> "" Then
                    WriteWhatis Rest, SearchIn
                    TU vsock, "3*** Description for '" & Rest & "' was set to '" & SearchIn & "'."
                  Else
                    If ReadWhatis(Rest, True, Nick) = "" Then
                      TU vsock, "5*** There's no description for '" & Rest & "'."
                    Else
                      Rest = WriteWhatis(Rest, "")
                      TU vsock, "3*** Description for '" & Rest & "' was deleted."
                    End If
                  End If
                Else
                  Rest = Param(Line, 2)
                  SearchIn = GetRest(Line, 3)
                  If SearchIn <> "" Then
                    WriteWhatis Rest, SearchIn
                    TU vsock, "3*** Description for '" & Rest & "' was set to '" & SearchIn & "'."
                  Else
                    If ReadWhatis(Rest, True, Nick) = "" Then
                      TU vsock, "5*** There's no description for '" & Rest & "'."
                    Else
                      Rest = WriteWhatis(Rest, "")
                      TU vsock, "3*** Description for '" & Rest & "' was deleted."
                    End If
                  End If
                End If
              End If
          Case ".wlist"
              If MatchFlags(Flags, "-w") Then
                TU vsock, "5*** You need the +w flag to use '.wlist'.": Exit Sub
              Else
                WList vsock, Line
              End If
          Case ".whatis"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".whatis <item>"): Exit Sub
              Rest = Trim(Replace(GetRest(Line, 2), """", " "))
              SearchIn = ReadWhatis(Rest, False, Nick)
              If SearchIn = "" Then
                TU vsock, "5*** There's no description for '" & Rest & "'."
              Else
                TU vsock, "3*** " & SearchIn
              End If
          Case ".voice"
              GUIVoice vsock, Line
          Case ".devoice"
              GUIDeVoice vsock, Line
          Case ".hop"
              GUIHop vsock, Line
          Case ".dehop"
              GUIDeHop vsock, Line
          Case ".op"
              GUIOp vsock, Line
          Case ".deop"
              GUIDeOp vsock, Line
          Case ".quit", ".q"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Param(Line, 2) = "" Then
                ToBotNet 0, "pt " & BotNetNick & " " & Nick & " " & SocketItem(vsock).OrderSign
                SpreadMessageEx 0, SocketItem(vsock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLLeave, Nick)
              Else
                ToBotNet 0, "pt " & BotNetNick & " " & Nick & " " & SocketItem(vsock).OrderSign & " " & Right(Line, Len(Line) - Len(Param(Line, 1)) - 1)
                SpreadMessageEx 0, SocketItem(vsock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLLeaveMsg, Nick, Right(Line, Len(Line) - Len(Param(Line, 1)) - 1))
              End If
              WriteSeenEntry Nick, "", Now, "*mine*", "*partyline*", Mask(SocketItem(vsock).Hostmask, 10)
              TU vsock, MSG_ThankYou
              
              RemoveSocket vsock, 0, "", True
              Exit Sub
          Case ".su", ".switchuser"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              If Param(Line, 2) = "" Then
                ToBotNet 0, "pt " & BotNetNick & " " & Nick & " " & SocketItem(vsock).OrderSign
                SpreadMessageEx 0, SocketItem(vsock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLLeave, Nick)
              Else
                ToBotNet 0, "pt " & BotNetNick & " " & Nick & " " & SocketItem(vsock).OrderSign & " " & Right(Line, Len(Line) - Len(Param(Line, 1)) - 1)
                SpreadMessageEx 0, SocketItem(vsock).PLChannel, SF_Local_JP, MakeMsg(MSG_PLLeaveMsg, Nick, Right(Line, Len(Line) - Len(Param(Line, 1)) - 1))
              End If
              WriteSeenEntry Nick, "", Now, "*mine*", "*partyline*", Mask(SocketItem(vsock).Hostmask, 10)
              TU vsock, MSG_ThankYou
              SocketItem(vsock).SendQLines = 0
              SocketItem(vsock).SendQTries = 0
              SocketItem(vsock).AwayMessage = ""
              SocketItem(vsock).SockTag = ""
              SocketItem(vsock).InputBuffer = ""
              SocketItem(vsock).LastEvent = Now
              SocketItem(vsock).CurrentQuestion = ""
              SocketItem(vsock).IRCNick = ""
              SetSockFlag vsock, SF_Status, SF_Status_UserGetName
              SocketItem(vsock).RegNick = ""
              TU vsock, " "
              TU vsock, "Please enter your user name:"
              SpreadFlagMessageEx 0, "+m", SF_Local_JP, MakeMsg(MSG_PLTelNetIncoming, SocketItem(vsock).Hostmask)
          Case ".relay"
              GUIRelay vsock, Line
          Case ".key"
              GUIKey vsock, Line
          Case ".say"
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".say <[" & ServerChannelPrefixes & "]channel>/<nick> <message>"): Exit Sub
              If IsValidChannel(Left(Param(Line, 2), 1)) Then
                If MatchFlags(GetUserChanFlags(Nick, Param(Line, 2)), "-m") Then TU vsock, "5*** Sorry, you don't have +m for this channel.": Exit Sub
                SendLine "privmsg " & Param(Line, 2) & " :" & GetRest(Line, 3), 2
                TU vsock, "3*** Said that in " & Param(Line, 2) & "."
              Else
                If MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you need global +m to make me talk to users.": Exit Sub
                SendLine "privmsg " & Param(Line, 2) & " :" & GetRest(Line, 3), 2
                TU vsock, "3*** Said that to " & Param(Line, 2) & "."
              End If
              SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 3
          Case ".action", ".act"
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".action <[" & ServerChannelPrefixes & "]channel>/<nick> <message>"): Exit Sub
              If Left(Param(Line, 2), 1) = "#" Then
                If MatchFlags(GetUserChanFlags(Nick, Param(Line, 2)), "-m") Then TU vsock, "5*** Sorry, you don't have +m for this channel.": Exit Sub
                SendLine "privmsg " & Param(Line, 2) & " :ACTION " & GetRest(Line, 3) & "", 2
                TU vsock, "3*** Did that in " & Param(Line, 2) & "."
              Else
                If MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you need global +m to make me talk to users.": Exit Sub
                SendLine "privmsg " & Param(Line, 2) & " :ACTION " & GetRest(Line, 3) & "", 2
                TU vsock, "3*** Did that to " & Param(Line, 2) & "."
              End If
              SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 3
          Case ".join"
              If Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".join <[" & ServerChannelPrefixes & "]channel> (key)"): Exit Sub
              If MatchFlags(Flags, "-n") Then
                If MatchFlags(GetUserChanFlags(Nick, Param(Line, 2)), "-n") Then SpreadFlagMessage vsock, "+n", "14[" & Time & "] *** " & Nick & " failed " & Line: PutLog "||| *** " & Nick & " failed " & Line: TU vsock, "5*** Sorry, you don't have +n for this channel.": Exit Sub
              End If
              If FindChan(Param(Line, 2)) <> 0 Then TU vsock, "5*** I'm already on this channel.": Exit Sub
              SpreadFlagMessage vsock, "+n", "14[" & Time & "] *** " & Nick & " did " & Line: PutLog "||| *** " & Nick & " did " & Line
              If Param(Line, 3) = "" Then SendLine "join " & Param(Line, 2), 1 Else SendLine "join " & Param(Line, 2) & " " & Param(Line, 3), 1
              TU vsock, "3*** Trying to join " & Param(Line, 2) & "..."
              SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
          Case ".part"
              If Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".part <[" & ServerChannelPrefixes & "]channel>"): Exit Sub
              If MatchFlags(Flags, "-n") Then
                If MatchFlags(GetUserChanFlags(Nick, Param(Line, 2)), "-n") Then SpreadFlagMessage vsock, "+n", MakeMsg(MSG_PLNickFailed, Nick, Line): PutLog "||| *** " & Nick & " failed " & Line: TU vsock, "5*** Sorry, you don't have +n for this channel.": Exit Sub
              End If
              If FindChan(Param(Line, 2)) = 0 Then TU vsock, "5*** I'm not on this channel.": Exit Sub
              SpreadFlagMessage vsock, "+n", MakeMsg(MSG_PLNickDid, Nick, Line): PutLog "||| *** " & Nick & " did " & Line
              SendLine "part " & Param(Line, 2) & " :C-Ya!", 1
              TU vsock, "3*** Leaving " & Param(Line, 2) & "..."
              SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
          Case ".boot"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".boot <nick> (reason)"): Exit Sub
              If Param(Line, 3) <> "" Then Messg = Right(Line, Len(Line) - 7 - Len(Param(Line, 2)))
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If LCase(SocketItem(u).RegNick) = LCase(Param(Line, 2)) Then
                    If GetSockFlag(u, SF_LocalVisibleUser) = SF_YES Then
                      If Messg = "" Then
                        TU u, MakeMsg(MSG_PLYouWereBooted, Nick)
                        SpreadMessage u, SocketItem(u).PLChannel, MakeMsg(MSG_PLNickWasBooted, SocketItem(u).RegNick, Nick)
                        ToBotNet 0, "pt " & BotNetNick & " " & SocketItem(u).RegNick & " " & SocketItem(u).OrderSign & " " & MakeMsg(MSG_BNBooted, Nick)
                      Else
                        TU u, MakeMsg(MSG_PLYouWereBooted2, Nick, Messg)
                        SpreadMessage u, SocketItem(u).PLChannel, MakeMsg(MSG_PLNickWasBooted2, SocketItem(u).RegNick, Nick, Messg)
                        ToBotNet 0, "pt " & BotNetNick & " " & SocketItem(u).RegNick & " " & SocketItem(u).OrderSign & " " & MakeMsg(MSG_BNBooted2, Nick, Messg)
                      End If
                      WriteSeenEntry SocketItem(u).RegNick, "", Now, "*mine*", "*partyline*", Mask(SocketItem(u).Hostmask, 10)
                      RemoveSocket u, 0, "", True
                      FoundOne = True
                    End If
                  End If
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** I couldn't find this user on the party line."
          Case ".adduser"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".adduser <nick> (handle)"): Exit Sub
              Rest = IIf(Param(Line, 3) = "", Param(Line, 2), Param(Line, 3))
              If Len(Rest) > ServerNickLen Then TU vsock, "5*** A nick can't be longer than " & CStr(ServerNickLen) & " characters!": Exit Sub
              If IsValidNick(Rest) = False Then TU vsock, "5*** """ & Rest & """ Erroneous Nickname": Exit Sub
              CheckLine = "": CheckDesc = ""
              For u = 1 To ChanCount
                UsNum = FindUser(Param(Line, 2), u)
                If UsNum > 0 Then
                  CheckLine = Channels(u).User(UsNum).Hostmask
                  CheckDesc = Channels(u).Name
                  If CheckLine <> "" Then
                    If Param(Line, 3) = "" Then Rest = Channels(u).User(UsNum).Nick: Exit For
                  End If
                End If
              Next u
              If CheckDesc = "" Then TU vsock, MakeMsg(ERR_NotOnLocalChans, Rest): Exit Sub
              If CheckLine = "" Then TU vsock, "5*** I couldn't get this user's hostmask. Please try again in a few seconds.": Exit Sub
              UsNum = GetUserNum(Rest)
              If UsNum > 0 Then
                RealNick = BotUsers(UsNum).Name
                Select Case AddHost(0, RealNick, Mask(CheckLine, 23))
                  Case AH_Success
                    TU vsock, "3*** Added the hostmask '" & Mask(CheckLine, 23) & "' to " & RealNick & "."
                    SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .adduser " & GetRest(Line, 2)
                    UpdateRegUsers "A " & Mask(CheckLine, 23)
                  Case AH_AlreadyThere
                    TU vsock, "5*** I already know this user and the current hostmask."
                  Case AH_MatchingUser
                    TU vsock, "5*** Can't add this hostmask - it matches " & ExtReply & "."
                  Case AH_TooManyHosts
                    TU vsock, "5*** Sorry, maximum number of hostmasks reached (20)."
                  Case Else
                    TU vsock, "5*** Some kind of strange error occurred. Please notify Hippo@animexx.de on how you did that. ;)"
                End Select
              Else
                Messg = SearchUserFromHostmask2(CheckLine)
                If Messg <> "" Then TU vsock, "5*** Can't add this user - hostmask matches " & Messg & ".": Exit Sub
                If MatchFlags(Flags, "-m") Then Messg = "p" Else Messg = BaseFlags
                Select Case AddUser(Rest, Messg)
                  Case AU_Success
                    If MatchFlags(Flags, "-m") Then
                      Messg = ""
                      For u2 = 1 To BotUsers(UserNum).ChannelFlagCount
                        If MatchFlags(BotUsers(UserNum).ChannelFlags(u2).Flags, "+m") Then
                          If Messg = "" Then Messg = BotUsers(UserNum).ChannelFlags(u2).Channel Else Messg = Messg & ", " & BotUsers(UserNum).ChannelFlags(u2).Channel
                          Chattr Rest, "+f " & BotUsers(UserNum).ChannelFlags(u2).Channel
                        End If
                      Next u2
                      TU vsock, "3*** Added user " & Rest & " with flag ""p"" (and local ""f"" for " & Messg & ")."
                    Else
                      TU vsock, "3*** Added user " & Rest & " with Flags """ & BaseFlags & """."
                    End If
                    Select Case AddHost(0, Rest, Mask(CheckLine, 23))
                      Case AH_Success
                        TU vsock, "3*** Hostmask: " & Mask(CheckLine, 23)
                        UpdateRegUsers "A " & Mask(CheckLine, 23)
                      Case Else
                        TU vsock, "3*** Hostmask: none (couldn't be added!)"
                    End Select
                    SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .adduser " & GetRest(Line, 2)
                  Case Else
                    TU vsock, "5*** The user couldn't be added. Please notify Hippo@animexx.de on how you did that. ;)"
                End Select
              End If
          Case ".+user"
              TargetNick = Param(Line, 2)
              If TargetNick = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+user <nick> (hostmask)"): Exit Sub
              If Param(Line, 3) <> "" Then
                If IsValidHostmask(Param(Line, 3)) = False Then TU vsock, "5*** A valid hostmask must look like this: nick!identd@host.domain": Exit Sub
                Messg = SearchUserFromHostmask3(Param(Line, 3))
                If Messg <> "" Then TU vsock, "5*** Can't add this user - hostmask belongs to " & Messg & ".": Exit Sub
              End If
              If MatchFlags(Flags, "-m") Then Messg = "p" Else Messg = BaseFlags
              Select Case AddUser(TargetNick, Messg)
                Case AU_TooLong
                  TU vsock, "5*** A nick can't be longer than " & ExtReply & " characters!"
                Case AU_InvalidNick
                  TU vsock, "5*** """ & Param(Line, 2) & """ Erroneous Nickname"
                Case AU_UserExists
                  TU vsock, "5*** Can't add this user - nickname already in use."
                Case AU_Success
                  If MatchFlags(Flags, "-m") Then
                    Messg = ""
                    For u2 = 1 To BotUsers(UserNum).ChannelFlagCount
                      If MatchFlags(BotUsers(UserNum).ChannelFlags(u2).Flags, "+m") Then
                        If Messg = "" Then Messg = BotUsers(UserNum).ChannelFlags(u2).Channel Else Messg = Messg & ", " & BotUsers(UserNum).ChannelFlags(u2).Channel
                        Chattr TargetNick, "+f " & BotUsers(UserNum).ChannelFlags(u2).Channel
                      End If
                    Next u2
                    TU vsock, "3*** Added user " & TargetNick & " with flag ""p"" (and local ""f"" for " & Messg & ")."
                  Else
                    TU vsock, "3*** Added user " & TargetNick & " with flags """ & BaseFlags & """."
                  End If
                  If Param(Line, 3) <> "" Then
                    Select Case AddHost(SocketItem(vsock).UserNum, TargetNick, Param(Line, 3))
                      Case AH_Success
                        TU vsock, "3*** Hostmask: " & Param(Line, 3)
                        UpdateRegUsers "A " & Param(Line, 3)
                      Case AH_MatchingUser
                        TU vsock, "3*** Hostmask: none (couldn't be added - it's matching " & ExtReply & "!)"
                      Case Else
                        TU vsock, "3*** Hostmask: none (couldn't be added!)"
                    End Select
                  Else
                    TU vsock, "3*** Hostmask: none"
                  End If
                  SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .+user " & GetRest(Line, 2)
              End Select
          Case ".trace"
              GUITrace vsock, Line
          Case ".+bot"
              TargetNick = Param(Line, 2)
              If TargetNick = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+bot <nick> (address:port) (+botflags)"): Exit Sub
              If Param(Line, 3) <> "" Then
                If Left(Param(Line, 3), 1) = "+" Then Rest = "": TheFlag = Param(Line, 3) Else Rest = Param(Line, 3): TheFlag = ""
                If Left(Param(Line, 4), 1) = "+" Then Rest = Param(Line, 3): TheFlag = Param(Line, 4)
                If TheFlag <> "" Then
                  If InStr(TheFlag, "h") > 0 Then TheFlag = TheFlag & "-a"
                  TheFlag = GetBotattrResult("", TheFlag)
                  If TheFlag = "" Then TU vsock, "5*** You specified invalid bot flags. Please take a look at '.help whois'.": Exit Sub
                End If
                If Rest <> "" Then
                  If InStr(Rest, ":") = 0 Then TU vsock, "5*** Please specify the new address like this: 'Host.Domain:Port'": Exit Sub
                  If (InStr(Rest, "*") > 0) Or (InStr(Rest, "!") > 0) Or (InStr(Rest, "?") > 0) Or (InStr(Rest, "@") > 0) Then TU vsock, "5*** The connect address may not contain the following characters: *, !, ?, @": Exit Sub
                End If
              Else
                TheFlag = ""
                Rest = ""
              End If
              CheckLine = "": CheckDesc = ""
              For u = 1 To ChanCount
                UsNum = FindUser(Param(Line, 2), u)
                If UsNum > 0 Then
                  FoundOne = True
                  TargetNick = Channels(u).User(UsNum).Nick
                  If Channels(u).User(UsNum).Hostmask <> "" Then
                    CheckLine = Channels(u).User(UsNum).Hostmask
                    CheckDesc = Channels(u).Name
                    Exit For
                  End If
                End If
              Next u
              If (FoundOne = True) And (CheckLine = "") Then TU vsock, "5*** I couldn't get this bot's hostmask. Please try again in a few seconds.": Exit Sub
              
              Select Case AddUser(TargetNick, "bf")
                Case AU_InvalidNick
                  TU vsock, "5*** """ & TargetNick & """ Erroneous Nickname"
                Case AU_TooLong
                  TU vsock, "5*** A nick can't be longer than " & ExtReply & " characters!"
                Case AU_UserExists
                  TU vsock, "5*** Can't add this bot - nickname is already in use."
                Case AU_Success
                  SucceedCommand vsock, "+t", Line
                  TU vsock, "3*** Added bot " & TargetNick & " with Flags ""bf""."
                  If CheckLine <> "" Then
                    Select Case AddHost(0, TargetNick, Mask(CheckLine, 21))
                      Case AH_MatchingUser
                        TU vsock, "3*** Hostmask: none ('" & Mask(CheckLine, 21) & "' couldn't be added - it's matching " & ExtReply & "!)"
                      Case AH_Success
                        TU vsock, "3*** Hostmask: " & Mask(CheckLine, 21) & " (from " & CheckDesc & ")."
                        UpdateRegUsers "A " & Mask(CheckLine, 21)
                      Case Else
                        TU vsock, "3*** Hostmask: none (couldn't be added!)"
                    End Select
                  End If
                  If Rest <> "" Then
                    SetUserData BotUserCount, UD_LinkAddr, Rest
                    TU vsock, "3*** Address : " & Rest
                  Else
                    TU vsock, "3*** Address : none"
                  End If
                  If TheFlag <> "" Then
                    BotUsers(BotUserCount).BotFlags = TheFlag
                    TU vsock, "3*** Botflags: +" & TheFlag
                  End If
                  SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .+bot " & GetRest(Line, 2)
              End Select
          Case ".remuser", ".-user"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, "'.remuser <nick>' or '.-user <nick>'"): Exit Sub
              UsNum = GetUserNum(Param(Line, 2))
              If UsNum = 0 Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
              RealNick = BotUsers(UsNum).Name
              If MatchFlags(Flags, "-n") And MatchFlags(GetUserFlags(RealNick), "+n") Then TU vsock, "5*** You can't remove an owner!!!": FailCommand vsock, "+m", Line: Exit Sub
              If MatchFlags(Flags, "-s") And MatchFlags(GetUserFlags(RealNick), "+s") Then TU vsock, "5*** You can't remove a super owner!!!": FailCommand vsock, "+m", Line: Exit Sub
              Select Case RemUser(RealNick)
                Case RU_UserNotFound
                  TU vsock, MakeMsg(ERR_UserNotFound, RealNick)
                Case RU_Success
                  SucceedCommand vsock, "+m", Line
                  'Boot removed user from the party line (if user didn't remove himself)
                  For u = 1 To SocketCount
                    If IsValidSocket(u) Then
                      If SocketItem(u).RegNick = RealNick Then
                        If GetSockFlag(u, SF_LocalVisibleUser) = SF_YES Then
                          If u = vsock Then
                            TU u, "4*** You just removed yourself from my userlist!"
                          Else
                            TU u, "4*** You were removed from my userlist. Bye!"
                            SpreadMessage u, SocketItem(u).PLChannel, "3*** " & SocketItem(u).RegNick & " was booted off the party line (removed from userlist)"
                            ToBotNet 0, "pt " & BotNetNick & " " & SocketItem(u).RegNick & " " & SocketItem(u).OrderSign & " removed from userlist"
                            RemoveSocket u, 0, "", True
                          End If
                        End If
                      End If
                    End If
                  Next u
                  TU vsock, "3*** Removed " & RealNick & ". All data lost."
                  UpdateRegUsers "R " & RealNick
                  SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .-user " & GetRest(Line, 2)
              End Select
          Case ".-bot"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-bot <nick>"): Exit Sub
              RealNick = GetRealNick(Param(Line, 2))
              If RealNick <> "" Then If MatchFlags(GetUserFlags(RealNick), "-b") Then TU vsock, "5*** Error: This is not a bot.": FailCommand vsock, "+t", Line:  Exit Sub
              Select Case RemUser(Param(Line, 2))
                Case RU_UserNotFound
                  TU vsock, "5*** I couldn't find this bot."
                Case RU_Success
                  SucceedCommand vsock, "+t", Line
                  TU vsock, "3*** Removed " & RealNick & ". All data lost."
                  UpdateRegUsers "R " & RealNick
                  SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .-bot " & GetRest(Line, 2)
              End Select
          Case ".+host"
              GUIAddHost vsock, Line
          Case ".-host"
              GUIRemHost vsock, Line
          Case ".chattr"
            For u = 1 To ParamXCount(GetRest(Line, 3), "|")
              GUIChattr vsock, ".chattr " & Param(Line, 2) & " " & ParamX(GetRest(Line, 3), "|", u)
            Next u
          Case ".botattr"
              GUIBotattr vsock, Line
          Case ".whatsnew"
              ShowNews vsock
          Case ".sharecmd"
              If Param(Line, 2) = "" Then TU vsock, "5*** Usage: .sharecmd <command>": Exit Sub
              SharingSpreadMessage SocketItem(vsock).RegNick, "cmd " & GetRest(Line, 2)
              Party vsock, SockNum, GetRest(Line, 2)
          Case ".check"
              Select Case Param(Line, 2)
              Case "1"
                  TU vsock, "2*** Searching for users without passwords..."
                  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                    TU vsock, "0,1 User:       | Flags:                  "
                  Else
                    TU vsock, " User:         Flags:                  "
                    TU vsock, " ------------  ------------------------"
                  End If
                  For u = 1 To BotUserCount
                    If (BotUsers(u).Password = "") And MatchFlags(BotUsers(u).Flags, "-b") Then
                      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                        TU vsock, " " & Spaces2(11, BotUsers(u).Name) & "14 | +" & BotUsers(u).Flags
                      Else
                        TU vsock, " " & Spaces2(11, BotUsers(u).Name) & "   +" & BotUsers(u).Flags
                      End If
                      u2 = u2 + 1
                    End If
                  Next u
                  TU vsock, "2*** Found " & CStr(u2) & " match" & IIf(u2 = 1, ".", "es.")
                  TU vsock, EmptyLine
              Case "2"
                  TU vsock, "2*** Searching for bots without passwords..."
                  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                    TU vsock, "0,1 User:       | Flags:                  "
                  Else
                    TU vsock, " User:         Flags:                  "
                    TU vsock, " ------------  ------------------------"
                  End If
                  For u = 1 To BotUserCount
                    If (BotUsers(u).Password = "") And MatchFlags(BotUsers(u).Flags, "+b") Then
                      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                        TU vsock, " " & Spaces2(11, BotUsers(u).Name) & "14 | +" & BotUsers(u).Flags
                      Else
                        TU vsock, " " & Spaces2(11, BotUsers(u).Name) & "   +" & BotUsers(u).Flags
                      End If
                      u2 = u2 + 1
                    End If
                  Next u
                  TU vsock, "2*** Found " & CStr(u2) & " match" & IIf(u2 = 1, ".", "es.")
                  TU vsock, EmptyLine
              Case "3"
                  TU vsock, "2*** Removing outdated channel flags..."
                  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                    TU vsock, "0,1 User:       | Channel:                "
                  Else
                    TU vsock, " User:         Channel:                "
                    TU vsock, " ------------  ------------------------"
                  End If
                  For u = 1 To BotUserCount
                    Do
                      FoundOne = False
                      For u2 = 1 To BotUsers(u).ChannelFlagCount
                        If Not InAutoJoinChannels(BotUsers(u).ChannelFlags(u2).Channel) Then
                          If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                            TU vsock, " " & Spaces2(11, BotUsers(u).Name) & "14 | " & BotUsers(u).ChannelFlags(u2).Channel
                          Else
                            TU vsock, " " & Spaces2(11, BotUsers(u).Name) & "   " & BotUsers(u).ChannelFlags(u2).Channel
                          End If
                          For u3 = u2 To BotUsers(u).ChannelFlagCount - 1
                            BotUsers(u).ChannelFlags(u3) = BotUsers(u).ChannelFlags(u3 + 1)
                          Next u3
                          BotUsers(u).ChannelFlagCount = BotUsers(u).ChannelFlagCount - 1
                          FoundOne = True
                          Exit For
                        End If
                      Next u2
                      If Not FoundOne Then Exit Do
                    Loop
                  Next u
                  TU vsock, "2*** Finished."
                  TU vsock, EmptyLine
              Case "4"
                u2 = 0
                TU vsock, "2*** Searching for users without hostmask..."
                TU vsock, " User:       14| Flags:                  "
                TU vsock, "14 ------------+-----------------------------"
                For u = 1 To BotUserCount
                  If BotUser(u, 5, 1) = "" Then
                    If Not InStr(BotUser(u, 2), "b") > 0 Then TU vsock, " " & Spaces2(12, BotUser(u, 1)) & Spaces2(24, "14| +" & BotUser(u, 2)): u2 = u2 + 1
                  End If
                Next u
                If u2 = 0 Then TU vsock, "14 none"
                TU vsock, "14 ------------+-----------------------------"
                TU vsock, "2*** Found " & u2 & " matches."
                Exit Sub
              Case "5"
                u2 = 0
                TU vsock, "2*** Searching for bots without hostmask..."
                TU vsock, " Bot:        14| Flags:                  "
                TU vsock, "14 ------------+----------------------------"
                For u = 1 To BotUserCount
                  If BotUser(u, 5, 1) = "" Then
                    If InStr(BotUser(u, 2), "b") > 0 Then TU vsock, " " & Spaces2(12, BotUser(u, 1)) & Spaces2(24, "14| +" & BotUser(u, 2)): u2 = u2 + 1
                  End If
                Next
                If u2 = 0 Then TU vsock, "14 none"
                TU vsock, "14 ------------+----------------------------"
                TU vsock, "2*** Found " & u2 & " matches."
                Exit Sub
              Case "6"
                u2 = 0
                TU vsock, "2*** Searching for users with +a ..."
                TU vsock, " User:       14| Flags:                  "
                TU vsock, "14 ------------+----------------------------"
                For u = 1 To BotUserCount
                  If InStr(BotUser(u, 2), "a") > 0 Then TU vsock, " " & Spaces2(12, BotUser(u, 1)) & Spaces2(24, "14| +" & BotUser(u, 2)): u2 = u2 + 1
                Next
                If u2 = 0 Then TU vsock, "14 none"
                TU vsock, "14 ------------+----------------------------"
                TU vsock, "2*** Found " & u2 & " matches."
                Exit Sub
              Case Else
                  TU vsock, MakeMsg(ERR_CommandUsage, ".check <1, 2, 3, 4, 5 or 6>")
                  TU vsock, "5           1 = list users without passwords"
                  TU vsock, "5           2 = list bots without passwords"
                  TU vsock, "5           3 = remove outdated channel flags"
                  TU vsock, "5           4 = list users without hostmask"
                  TU vsock, "5           5 = list bots without hostask"
                  TU vsock, "5           6 = list users with autoop"
                  Exit Sub
              End Select
          Case ".match", ".lmatch"
              Match vsock, Line
          Case ".status", ".stat", ".st"
              TU vsock, "2*** STATUS report for " & BotNetNick & ", AnGeL " & BotVersionEx
              If Connected Then
                TU vsock, "3    Bot Type : " & IIf(IsNTService, "WindowsNT Service", IIf(WinNTOS, "WindowsNT Application", "Windows Application")) & " (" & WinVersionName & ")"
                TU vsock, "3    Uptime   : " & TimeSince(CStr(StartUpTime))
                TU vsock, "3    Online as: " & MyNick & " 14(" & Mask(MyHostmask, 10) & ")" & IIf(RestrictedIndex > 0, " 4<RESTRICTED>", "")
                TU vsock, "3    Real name: " & RealName
                TU vsock, "3    Server   : " & StripDP(ServerName) & " 14(connected for " & TimeSince(CStr(ConnectTime)) & ")"
                TU vsock, "3    Network  : " & IIf(ServerNetwork <> "", ServerNetwork, "14unknown") & " 14(NickLen=" & ServerNickLen & " MaxChans=" & ServerMaxChannels & ")"
 
              Else
                TU vsock, "3    Bot Type : " & IIf(IsNTService, "WindowsNT Service", IIf(WinNTOS, "WindowsNT Application", "Windows Application")) & " (" & WinVersionName & ")"
                TU vsock, "3    Uptime   : " & TimeSince(CStr(StartUpTime))
                TU vsock, "3    Online as: 4<not online!>"
              End If
              TU vsock, EmptyLine
              ChanList vsock, "4"
              TU vsock, "2*** End of STATUS report"
          Case ".comment"
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".comment <nick> <comment / ban reason>"): TU vsock, "5*** To remove a comment, type '.comment <nick> none'.": Exit Sub
              UsNum = GetUserNum(Param(Line, 2))
              If UsNum = 0 Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
              If LCase(Param(Line, 3)) <> "none" Then
                If Len(GetRest(Line, 3)) > 80 Then TU vsock, "5*** Sorry, this info line is " & CStr(Len(GetRest(Line, 3)) - 80) & " characters too long.": Exit Sub
                SetUserData UsNum, "comment", GetRest(Line, 3)
                TU vsock, "3*** The comment of " & BotUsers(UsNum).Name & " was set."
              Else
                SetUserData UsNum, "comment", ""
                TU vsock, "3*** The comment of " & BotUsers(UsNum).Name & " was erased."
              End If
              SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .comment " & GetRest(Line, 2)
          Case ".urls"
              If Dir(HomeDir & "urls.txt") = "" Then
                TU vsock, "5*** Sorry there are no known URL's"
                Exit Sub
              End If
              FileNum = FreeFile: Open HomeDir & "URLS.txt" For Input As #FileNum
              Do While Not EOF(FileNum)
                Line Input #FileNum, CheckLine
                CheckURL = Left(CheckLine, InStr(CheckLine, " ") - 1)
                CheckDesc = Right(CheckLine, Len(CheckLine) - Len(CheckURL) - 1)
                CheckDesc = CheckDesc & " "
                If Not ToldOneHostMask Then TU vsock, "2*** Listing known URL's"
                ToldOneHostMask = True
                CheckHost = ""
                TU vsock, " 7-4=5>1 " & CheckURL
                For u = 1 To Len(CheckDesc)
                  If Mid(CheckDesc, u, 1) <> " " Then
                    CheckHost = CheckHost + Mid(CheckDesc, u, 1)
                  Else
                    If InStr(u + 1, CheckDesc, " ") - u > 61 - Len(CheckHost) Then
                      TU vsock, "14     " & CheckHost
                      CheckHost = ""
                    Else
                      CheckHost = CheckHost & " "
                    End If
                  End If
                Next u
                TU vsock, "14     " & CheckHost
              Loop
              Close #FileNum
          Case ".chansetup", ".chanset", ".cs"
              If Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then
                If Not LCase(Param(Line, 2)) = "default" Then
                  TU vsock, MakeMsg(ERR_CommandUsage, ".chansetup <[" & ServerChannelPrefixes & "]channel>")
                  Exit Sub
                End If
              End If
              If MatchFlags(Flags, "-n") Then
                If MatchFlags(GetUserChanFlags(Nick, Param(Line, 2)), "-n") Then TU vsock, "5*** Sorry, you're not allowed to setup this channel.": Exit Sub
              End If
              If Param(Line, 3) <> "" Then
                Messg = GetRest(Line, 3)
                SocketItem(vsock).SetupChan = Param(Line, 2)
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                For u = 1 To ParamXCount(Messg, ",")
                  ChanSetup vsock, Trim(ParamX(Messg, ",", u))
                Next u
                TU vsock, "3*** Done."
              Else
                If UserNum = 0 Then Exit Sub  'Don't allow scripts to enter the chansetup menu
                SetAway vsock, "Channel setup"
                SetSockFlag vsock, SF_Status, SF_Status_ChanSetup
                If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                  TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1               5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                  TU vsock, "8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1  Channel setup  5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                  TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1               5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                End If
                TU vsock, EmptyLine
                TU vsock, "Welcome to the channel setup for 2" & Param(Line, 2) & "."
                TU vsock, "Here are the current settings of this channel:"
                TU vsock, EmptyLine
                SocketItem(vsock).SetupChan = Param(Line, 2)
                SocketItem(vsock).CurrentQuestion = "ChooseSetting"
                ShowCSActions vsock
              End If
          Case ".talk"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".talk <message>"): Exit Sub
              If (InStr(LCase(Line), "password") > 0) Or (InStr(LCase(Line), "enter your") > 0) Then TU vsock, "5*** Sorry, no abusive messages allowed.": Exit Sub
              ToBotNet 0, "c " & BotNetNick & " A " & Right(Line, Len(Line) - Len(Param(Line, 1)) - 1)
              SpreadMessage vsock, -1, Right(Line, Len(Line) - Len(Param(Line, 1)) - 1)
              TU vsock, "3*** Said that in the botnet."
          Case ".botsetup", ".bs"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              SetAway vsock, "Bot setup"
              SetSockFlag vsock, SF_Status, SF_Status_BotSetup
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1           5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                TU vsock, "8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1  Bot setup  5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1           5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
              End If
              TU vsock, EmptyLine
              TU vsock, "Welcome to the bot setup. Here are the current settings:"
              TU vsock, EmptyLine
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              ShowBSActions vsock
          Case ".netsetup", ".ns"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              SetAway vsock, "Network setup"
              SetSockFlag vsock, SF_Status, SF_Status_NETSetup
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1               2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                TU vsock, "11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1  Network setup  2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1               2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
              End If
              TU vsock, EmptyLine
              TU vsock, "Welcome to the network setup. Here are the current settings:"
              TU vsock, EmptyLine
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              ShowNETActions vsock
          Case ".polsetup", ".ps", ".policysetup"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              SetAway vsock, "Policy setup"
              SetSockFlag vsock, SF_Status, SF_Status_POLSetup
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, " 8,0,%0,8%'9,8,%8,9%'3,9,%9,3%'1,3,%3,1%'0,1              3,1'%1,3%,9,3'%3,9%,8,9'%9,8%,0,8'%8,0%,"
                TU vsock, "8,0,%0,8%'9,8,%8,9%'3,9,%9,3%'1,3,%3,1%'0,1  Policy setup  3,1'%1,3%,9,3'%3,9%,8,9'%9,8%,0,8'%8,0%,"
                TU vsock, " 8,0,%0,8%'9,8,%8,9%'3,9,%9,3%'1,3,%3,1%'0,1              3,1'%1,3%,9,3'%3,9%,8,9'%9,8%,0,8'%8,0%,"
              End If
              TU vsock, EmptyLine
              TU vsock, "Welcome to the policy setup. Here are the current settings:"
              TU vsock, EmptyLine
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              ShowPolActions vsock
          Case ".authsetup", ".as"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              SetAway vsock, "AUTH setup"
              SetSockFlag vsock, SF_Status, SF_Status_AUTHSetup
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1            2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                TU vsock, "11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1  AUTH setup  2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
                TU vsock, " 11,0,%0,11%'12,11,%11,12%'2,12,%12,2%'1,2,%2,1%'0,1            2,1'%1,2%,12,2'%2,12%,11,12'%12,11%,0,11'%11,0%,"
              End If
              TU vsock, EmptyLine
              TU vsock, "Welcome to the AUTH setup. Here are the current settings:"
              TU vsock, EmptyLine
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              ShowAUTHActions vsock
          Case ".auth"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              CommandAuth vsock, Line
          Case ".kisetup", ".ks"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              SetAway vsock, "KI setup"
              SetSockFlag vsock, SF_Status, SF_Status_KISetup
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1          5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                TU vsock, "8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1  KI setup  5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1          5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
              End If
              TU vsock, EmptyLine
              TU vsock, "Welcome to the KI setup. Here are the current settings:"
              TU vsock, EmptyLine
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              ShowKIActions vsock
          Case ".setup", ".console"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              SetAway vsock, "Personal setup"
              SetSockFlag vsock, SF_Status, SF_Status_PersonalSetup
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1                5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                TU vsock, "8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1  Personal setup  5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
                TU vsock, " 8,0,%0,8%'4,8,%8,4%'5,4,%4,5%'1,5,%5,1%'0,1                5,1'%1,5%,4,5'%5,4%,8,4'%4,8%,0,8'%8,0%,"
              End If
              TU vsock, EmptyLine
              TU vsock, "Welcome to the personal setup of 2" & SocketItem(vsock).RegNick & "."
              TU vsock, "Here are the current settings:"
              TU vsock, EmptyLine
              SocketItem(vsock).CurrentQuestion = "ChooseSetting"
              ShowPSActions vsock
          Case ".update"
              If Dir(HomeDir & "Update.exe") = "" Then TU vsock, "5*** You must DCC send the file 'AnGeL.exe' first to run AutoUpdate.": Exit Sub
              SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Calculating AutoUpdate file checksum..."
              If ValidChecksum(HomeDir & "Update.exe") Then
                DontConnect = True
                AutoUpdate = True
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** AutoUpdate file is valid, proceeding..."
                Status "*** AutoUpdate requested..." & vbCrLf
                SpreadMessage 0, -1, "4*** AUTO UPDATE sequence started"
                SpreadMessage 0, -1, "7*** Sorry, I have to log you off! Bye bye..."
                SpreadMessage 0, -1, ""
                Shell HomeDir & "Update.exe"
                TimedEvent "winsock2_shutdown", 0
              Else
                SpreadFlagMessage 0, "+m", "14[" & Time & "]4 *** ERROR: Corrupt AutoUpdate file - bad checksum!"
                SpreadFlagMessage 0, "+m", "14[" & Time & "]4     Please check the file and your system for viruses."
                SpreadFlagMessage 0, "+m", "14[" & Time & "]4     Only use update files downloaded directly from the"
                SpreadFlagMessage 0, "+m", "14[" & Time & "]4     official AnGeL Homepage: www.angel-bot.de"
                Kill HomeDir & "Update.exe"
              End If
          Case ".reconnect"
              DontConnect = False
              Status "*** Reconnect requested..." & vbCrLf
              SpreadMessage 0, -1, "7*** RECONNECT requested by " & Nick & ""
              SendLine "quit :Reconnecting...", 1
              Disconnect
              Output vbCrLf
              Output "*** Socket Closed (me)" & vbCrLf
              ConnectServer JumpConnectDelay, ""
          Case ".jump"
              GUIJump vsock, Line
          Case ".restart"
              If Dir(HomeDir & App.EXEName & ".exe") = "" Then TU vsock, "5*** AnGeL Binary not found!": Exit Sub
              Status "*** Restart requested..." & vbCrLf
              SpreadMessage 0, -1, "7*** RESTART requested by " & Nick & ""
              TimedEvent "RESTART", 0
          Case ".die"
              If Param(Line, 2) <> "sure" Then TU vsock, "5*** Are you sure?! Type '.die sure' if you really want to shut me down.": Exit Sub
              Status "*** Die requested..." & vbCrLf
              SpreadMessage 0, -1, "7*** DIE requested by " & Nick & ""
              BotRestart = False
              TimedEvent "winsock2_shutdown", 0
          Case ".chnick"
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".chnick <old nick> <new nick>"): Exit Sub
              If Len(Param(Line, 3)) > ServerNickLen Then TU vsock, "5*** A nick can't be longer than " & CStr(ServerNickLen) & " characters!": Exit Sub
              If IsValidNick(Param(Line, 3)) = False Then TU vsock, "5*** """ & Param(Line, 3) & """ Erroneous Nickname": Exit Sub
              RealNick = GetRealNick(Param(Line, 2))
              If Not UserExist(RealNick) Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
              If MatchFlags(GetUserFlags(RealNick), "+s") And MatchFlags(Flags, "-s") Then TU vsock, "5*** You can't change a super owner's nick!": Exit Sub
              If MatchFlags(GetUserFlags(RealNick), "+n") And MatchFlags(Flags, "-n") Then TU vsock, "5*** You can't change an owner's nick!": Exit Sub
              If UserExist(Param(Line, 3)) And Not (LCase(Param(Line, 3)) = LCase(RealNick) And Param(Line, 3) <> RealNick) Then TU vsock, "5*** """ & Param(Line, 3) & """ nickname is already in use.": Exit Sub
              ChangeNick RealNick, Param(Line, 3)
              SharingSpreadMessage Nick, ".chnick " & RealNick & " " & Param(Line, 3)
              FoundOne = False
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If LCase(SocketItem(u).RegNick) = LCase(RealNick) And SocketItem(u).OnBot = BotNetNick Then
                    'Don't change the nick of bots
                    If Not ((GetSockFlag(u, SF_Status) = SF_Status_Bot) Or (GetSockFlag(u, SF_Status) = SF_Status_BotLinking)) Then
                      ToBotNet 0, "nc " & BotNetNick & " " & SocketItem(u).OrderSign & " " & Param(Line, 3)
                      If Not FoundOne Then SpreadMessage 0, SocketItem(u).PLChannel, "3*** " & RealNick & " is now known as " & Param(Line, 3)
                      SocketItem(u).RegNick = Param(Line, 3)
                      FoundOne = True
                    End If
                  End If
                End If
              Next u
              If Not FoundOne Then TU vsock, "3*** Okay, done. " & RealNick & " is now known as " & Param(Line, 3) Else TU vsock, "3*** Done."
          Case ".note"
              If Param(Line, 3) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".note <nick(@bot)> <message>"): Exit Sub
              If InStr(Param(Line, 2), "@") = 0 Then
                'Local note
                RealNick = GetRealNick(Param(Line, 2))
                If Not UserExist(RealNick) Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
                If MatchFlags(GetUserFlags(RealNick), "+b") Then TU vsock, "5*** Sorry, you can't send notes to bots.": Exit Sub
                If RealNick <> Nick Then
                  For u = 1 To SocketCount
                    If IsValidSocket(u) Then
                      If (SocketItem(u).RegNick = RealNick) And (GetSockFlag(u, SF_Status) = SF_Status_Party) And (SocketItem(u).AwayMessage = "") Then TU u, "5*" & Nick & "* " & GetRest(Line, 3): FoundOne = True
                    End If
                  Next u
                End If
                If Not FoundOne Then
                  If SendNote(Nick, RealNick, "", GetRest(Line, 3)) Then
                    TU vsock, "3*** Your note was stored."
                    CheckNotes RealNick
                  Else
                    TU vsock, "5*** Sorry, this user's mailbox is full."
                  End If
                Else
                  TU vsock, "3*** Your note was delivered."
                End If
              Else
                'Botnet note
                Rest = GetPartBot(Param(Line, 2))
                RealNick = GetPartNick(Param(Line, 2))
                If GetBotPos(Rest) = 0 Then TU vsock, MakeMsg(ERR_BotNotFound, Param(Line, 2)): Exit Sub
                SendToBot Rest, "p " & CStr(vsock) & ":" & Nick & "@" & BotNetNick & " " & Param(Line, 2) & " " & GetRest(Line, 3)
              End If
          Case ".notes"
              If UserNum = 0 Then Exit Sub   'Don't execute when a script is calling this command
              NoteCount = NotesCount(Nick)
              If NoteCount > 0 Then
                Select Case LCase(Param(Line, 2))
                  Case "read", ""
                      TU vsock, "2*** Listing " & IIf(Param(Line, 2) = "", "and erasing ", "") & "your notes:"
                      For u = 1 To NoteCount
                        TheFlag = NotesFlag(Nick, u)
                        Rest = NotesDate(Nick, u)
                        If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                          If TheFlag = "" Then
                            TU vsock, "2" & CStr(u) & ".3 " & NotesFrom(Nick, u) & " 14(" & Format(GetDate(Rest), "dd.mm.yy, hh:nn") & "): " & NotesText(Nick, u)
                          Else
                            TU vsock, "2" & CStr(u) & ".3 " & NotesFrom(Nick, u) & " 10[" & TheFlag & "] 14(" & Format(GetDate(Rest), "dd.mm.yy, hh:nn") & "): " & NotesText(Nick, u)
                          End If
                        Else
                          If TheFlag = "" Then
                            TU vsock, CStr(u) & ". " & NotesFrom(Nick, u) & " (" & Format(GetDate(Rest), "dd.mm.yy, hh:nn") & "): " & NotesText(Nick, u)
                          Else
                            TU vsock, CStr(u) & ". " & NotesFrom(Nick, u) & " [" & TheFlag & "] (" & Format(GetDate(Rest), "dd.mm.yy, hh:nn") & "): " & NotesText(Nick, u)
                          End If
                        End If
                      Next u
                      TU vsock, EmptyLine
                      If Param(Line, 2) = "" Then NotesErase Nick
                  Case "erase"
                      NotesErase Nick
                      TU vsock, "3*** Erased all notes."
                  Case Else
                      TU vsock, MakeMsg(ERR_CommandUsage, ".notes (read / erase)")
                End Select
              Else
                Select Case LCase(Param(Line, 2))
                  Case "read", ""
                      TU vsock, "5*** There are no notes waiting for you."
                  Case "erase"
                      TU vsock, "5*** There are no notes to erase."
                  Case Else
                      TU vsock, MakeMsg(ERR_CommandUsage, ".notes (read / erase)")
                End Select
              End If
          Case ".fwd"
            If Param(Line, 2) = "x" Then SetUserData SocketItem(vsock).UserNum, "fwdaddr", "": TU vsock, "3*** Forward adress deleted.": Exit Sub
            If Param(Line, 2) = "" Or InStr(Param(Line, 2), "@") = 0 Then
              TU vsock, "5*** Usage: .fwd <nick@bot> or .fwd x to delete."
            Else
              Rest = GetPartBot(Param(Line, 2))
              If GetBotPos(Rest) = 0 Then TU vsock, MakeMsg(ERR_BotNotFound, Param(Line, 2)): Exit Sub
              SetUserData SocketItem(vsock).UserNum, "fwdaddr", Param(Line, 2)
              TU vsock, "3*** Forward adress set to: 10" & Param(Line, 2)
            End If
          Case ".flagnote"
              Delivered = 0
              If Not (IsValidChannel(LCase(Left(Param(Line, 3), 1)))) Then
                If Param(Line, 3) = "" Or Left(Param(Line, 2), 1) <> "+" Then TU vsock, MakeMsg(ERR_CommandUsage, ".flagnote <+flags> ([" & ServerChannelPrefixes & "]channel) <message>"): Exit Sub
                TU vsock, "3*** Delivering flagnote... please wait."
                CheckHost = CStr(Now)
                For u = 1 To BotUserCount
                  'Don't send notes to bots
                  If MatchFlags(BotUsers(u).Flags, "-b") Then
                    If MatchFlags(BotUsers(u).Flags, Param(Line, 2)) Then
                      If Nick <> BotUsers(u).Name Then
                        FoundOne = False
                        For u2 = 1 To SocketCount
                          If IsValidSocket(u2) Then
                            If (SocketItem(u2).RegNick = BotUsers(u).Name) And (GetSockFlag(u2, SF_Status) = SF_Status_Party) And (SocketItem(u2).AwayMessage = "") Then TU u2, "5*" & Nick & " [" & LCase(Param(Line, 2)) & "]* " & GetRest(Line, 3): FoundOne = True
                          End If
                        Next u2
                        If Not FoundOne Then SendNote Nick, BotUsers(u).Name, Param(Line, 2), GetRest(Line, 3), CheckHost
                        Delivered = Delivered + 1
                      End If
                    End If
                  End If
                Next u
              Else
                If Param(Line, 4) = "" Or Left(Param(Line, 2), 1) <> "+" Then TU vsock, MakeMsg(ERR_CommandUsage, ".flagnote <+flags> ([" & ServerChannelPrefixes & "]channel) <message>"): Exit Sub
                TU vsock, "3*** Delivering channel flagnote... please wait."
                For u = 1 To BotUserCount
                  'Don't winsock2_send notes to bots
                  If MatchFlags(BotUsers(u).Flags, "-b") Then
                    If MatchFlags(GetUserChanFlags2(u, Param(Line, 3)), Param(Line, 2)) Then
                      If Nick <> BotUsers(u).Name Then
                        FoundOne = False
                        For u2 = 1 To SocketCount
                          If IsValidSocket(u2) Then
                            If (SocketItem(u2).RegNick = BotUsers(u).Name) And (GetSockFlag(u2, SF_Status) = SF_Status_Party) And (SocketItem(u2).AwayMessage = "") Then TU u2, "5*" & Nick & " [" & LCase(Param(Line, 2)) & " " & Param(Line, 3) & "]* " & GetRest(Line, 4): FoundOne = True
                          End If
                        Next u2
                        If Not FoundOne Then SendNote Nick, BotUsers(u).Name, Param(Line, 2) & " " & LCase(Param(Line, 3)), GetRest(Line, 4)
                        Delivered = Delivered + 1
                      End If
                    End If
                  End If
                Next u
              End If
              If Delivered = 0 Then
                TU vsock, "5*** Sorry, I didn't find any users matching the flags you specified."
              Else
                TU vsock, "3*** Your note was sent to " & CStr(Delivered) & " '" & Param(Line, 2) & "' users."
              End If
              CheckNotes ""
          Case ".+chan"
              If Param(Line, 2) = "" Or Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".+chan <[" & ServerChannelPrefixes & "]channel>"): Exit Sub
              If InStr(Param(Line, 2), ",") > 0 Then TU vsock, "5*** You can't use commas in channel names!": Exit Sub
              If Not LineInFile(Param(Line, 2), HomeDir & "Autojoin.txt") Then
                AddLineToFile Param(Line, 2), HomeDir & "Autojoin.txt"
                If FindChan(Param(Line, 2)) = 0 Then SendLine "join " & Param(Line, 2) & IIf(Param(Line, 3) <> "", " " & Param(Line, 3), ""), 1
                ReadAutoJoinChannels
                If Param(Line, 3) = "" Then
                  TU vsock, "3*** Added permanent channel " & Param(Line, 2) & "."
                Else
                  WritePPString Param(Line, 2), "EnforceModes", CombineModes(GetPPString(Param(Line, 2), "EnforceModes", "", HomeDir & "Channels.ini"), "+k " & Param(Line, 3)), HomeDir & "Channels.ini"
                  TU vsock, "3*** Added permanent channel " & Param(Line, 2) & " with key '" & Param(Line, 3) & "'."
                End If
              Else
                TU vsock, "5*** This channel is already permanent."
              End If
              SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
          Case ".-chan"
              If Param(Line, 2) = "" Or Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".-chan <[" & ServerChannelPrefixes & "]channel>"): Exit Sub
              If RemovePermChan(Param(Line, 2)) = 0 Then TU vsock, "5*** This is not one of my permanent channels.": Exit Sub
              If FindChan(Param(Line, 2)) <> 0 Then SendLine "part " & Param(Line, 2) & " :C-Ya!", 1
              TU vsock, "3*** Removed permanent channel " & Param(Line, 2) & "."
          Case ".cycle", ".rejoin"
              If Param(Line, 2) = "" Or Not (IsValidChannel(LCase(Left(Param(Line, 2), 1)))) Then TU vsock, MakeMsg(ERR_CommandUsage, ".cycle <[" & ServerChannelPrefixes & "]channel>"): Exit Sub
              If FindChan(Param(Line, 2)) = 0 Then TU vsock, "5*** This is none of my channels.": Exit Sub
              SendLine "part " & Param(Line, 2) & " :Rejoining... ", 1
              TimedEvent "join " & Param(Line, 2), 5
          Case ".chanlist", ".cl"
              ChanList vsock, Line
          Case ".chpass"
              If MatchFlags(Flags, "-n") Then
                If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".chpass <bot> (password)"): Exit Sub
                If GetRealNick(Param(Line, 2)) = "" Then TU vsock, "5*** Sorry, I couldn't find this bot.": Exit Sub
                If MatchFlags(GetUserFlags(GetRealNick(Param(Line, 2))), "-b") Then
                  TU vsock, "5*** Sorry, you may only change passwords of bots."
                  FailCommand vsock, "+m", ".chpass " & Param(Line, 2) + IIf(Param(Line, 3) <> "", " (something)", " (nothing)")
                  Exit Sub
                End If
              Else
                If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".chpass <user/bot> (password)"): Exit Sub
                If GetRealNick(Param(Line, 2)) = "" Then TU vsock, MakeMsg(ERR_UserNotFound, Param(Line, 2)): Exit Sub
              End If
              TargetNick = GetRealNick(Param(Line, 2))
              If MatchFlags(Flags, "-s") And MatchFlags(GetUserFlags(TargetNick), "+s") Then
                TU vsock, "5*** " & TargetNick & " is a super owner. You can't change the password of this user."
                FailCommand vsock, "+m", ".chpass " & Param(Line, 2) + IIf(Param(Line, 3) <> "", " (something)", " (nothing)")
                Exit Sub
              End If
              If Param(Line, 3) <> "" Then
                If Len(Param(Line, 3)) > 5 Then
                  If WeakPass(Param(Line, 3), TargetNick) Then TU vsock, "5*** This password is too weak. Please use a more intelligent one. :)": Exit Sub
                  'Don't encrypt bot passwords
                  If MatchFlags(GetUserFlags(TargetNick), "+b") Then
                    UserNum = GetUserNum(Param(Line, 2))
                    BotUsers(UserNum).Password = Param(Line, 3)
                  Else
                    UserNum = GetUserNum(Param(Line, 2))
                    BotUsers(UserNum).Password = EncryptIt(Param(Line, 3))
                  End If
                  TU vsock, "3*** Changed password to '" & Param(Line, 3) & "'."
                  SucceedCommand vsock, "+m", ".chpass " & Param(Line, 2) & " (something)"
                Else
                  TU vsock, "5*** This password is too short! Please use more than 5 characters."
                End If
              Else
                UserNum = GetUserNum(Param(Line, 2))
                BotUsers(UserNum).Password = ""
                TU vsock, "3*** Removed password."
                SucceedCommand vsock, "+m", ".chpass " & Param(Line, 2) & " (nothing)"
              End If
          Case ".link"
              Rest = LCase(Param(Line, 2))
              If Rest = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".link <bot>"): Exit Sub
              UserNum = GetUserNum(Rest)
              If UserNum = 0 Then TU vsock, "5*** Sorry, I couldn't find that bot.": Exit Sub
              If MatchFlags(BotUsers(UserNum).Flags, "-b") Then TU vsock, "5*** Sorry, " & BotUsers(UserNum).Name & " is not a bot (+b flag is missing).": Exit Sub
              For u = 1 To BotCount
                If LCase(Bots(u).Nick) = Rest Then TU vsock, "5*** Sorry, that bot is already connected.": Exit Sub
              Next u
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If LCase(SocketItem(u).RegNick) = Rest Then
                    Select Case GetSockFlag(u, SF_Status)
                      Case SF_Status_BotLinking, SF_Status_BotGetPass, SF_Status_InitBotLink
                        If SocketItem(u).Hostmask & ":" & SocketItem(u).RemotePort = GetUserData(UserNum, UD_LinkAddr, "") Then
                          TU vsock, "5*** I'm already trying to link to this bot."
                          Exit Sub
                        Else
                          TU vsock, "3*** Stopped trying to link to " & SocketItem(u).RegNick
                          RTU u, "bye " & Rest
                          RemoveSocket u, 0, "", True
                          Exit For
                        End If
                    End Select
                  End If
                End If
              Next u
              If InStr(GetUserData(UserNum, UD_LinkAddr, ""), ":") = 0 Then
                Rest = ""
                For u = 1 To ChanCount
                  For u2 = 1 To Channels(u).UserCount
                    If Channels(u).User(u2).RegNick = BotUsers(UserNum).Name Then
                      Rest = Channels(u).User(u2).Nick
                      Exit For
                    End If
                  Next u2
                  If Rest <> "" Then Exit For
                Next u
                If Rest = "" Then
                  TU vsock, "5*** Please set a connect address with '.chaddr " & BotUsers(UserNum).Name & " <address:port>' first."
                Else
                  TU vsock, "3*** Trying to get " & BotUsers(UserNum).Name & "'s connect address..."
                  TU vsock, "3*** If there's no reaction, set an address yourself: .chaddr " & BotUsers(UserNum).Name & " <address:port>"
                  SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLBotNetLinking, BotUsers(UserNum).Name, "Sending my link address (" & IrcGetAscIp(MyIP) & ":" & Trim(Str(BotnetPort)) & ")")
                  If BotUsers(UserNum).Password <> "" Then
                    SendLine "privmsg " & Rest & " :ANGEL LINK REQ! " & BotNetNick & " " & EncryptString(RandString & " " & IrcGetAscIp(MyIP) & ":" & Trim(Str(BotnetPort)) & " " & RandString, EncryptIt(BotUsers(UserNum).Password)) & "", 2
                  Else
                    SendLine "privmsg " & Rest & " :ANGEL LINK REQ " & BotNetNick & " " & IrcGetAscIp(MyIP) & ":" & Trim(Str(BotnetPort)) & "", 2
                  End If
                End If
              ElseIf (InStr(Param(Line, 3), "*") > 0) Or (InStr(Param(Line, 3), "!") > 0) Or (InStr(Param(Line, 3), "?") > 0) Or (InStr(Param(Line, 3), "@") > 0) Then
                TU vsock, "5*** The connect address may not contain the following characters: *, !, ?, @"
                TU vsock, "5*** Please set a valid address with '.chaddr " & BotUsers(UserNum).Name & " <address:port>'."
              Else
                TU vsock, "3*** Trying to link to " & BotUsers(UserNum).Name & " at 10" & GetUserData(UserNum, UD_LinkAddr, "") & "3."
                InitiateBotChat UserNum, False
              End If
          Case ".botinfo"
              TU vsock, "14[" & BotNetNick & "] AnGeL " & BotVersion & IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", "") & " (" & BotChannels & ") [uptime: " & TimeSince(CStr(StartUpTime)) & "]"
              For u = 1 To SocketCount
                If IsValidSocket(u) Then
                  If GetSockFlag(u, SF_Status) = SF_Status_Bot Then RTU u, "i? " & CStr(vsock) & ":" & Nick & "@" & BotNetNick
                End If
              Next u
          Case ".unlink"
              If Param(Line, 2) <> "*" Then
                Rest = GetRest(Line, 3)
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then
                    RealNick = SocketItem(u).RegNick
                    If LCase(RealNick) = LCase(Param(Line, 2)) Then
                      Select Case GetSockFlag(u, SF_Status)
                        Case SF_Status_Bot, SF_Status_BotLinking
                            TU vsock, "3*** Breaking link with " & RealNick
                            RTU u, "bye " & Rest
                            RemoveSocket u, 0, MakeMsg(MSG_BNDisconnect, IIf(Rest <> "", "³" & Rest, "")), True
                            FoundOne = True
                        Case SF_Status_InitBotLink
                            TU vsock, "3*** Stopped trying to link to " & RealNick
                            RTU u, "bye " & Rest
                            RemoveSocket u, 0, "", True
                            FoundOne = True
                      End Select
                    End If
                  End If
                Next u
                If Not FoundOne Then
                  u = GetBotPos(Param(Line, 2))
                  If u > 0 Then
                    TU vsock, "3*** Sending unlink request to " & Bots(u).SubBotOf & "..."
                    SendToBot Bots(u).SubBotOf, "ul " & CStr(vsock) & ":" & Nick & "@" & BotNetNick & " " & Bots(u).SubBotOf & " " & Bots(u).Nick + IIf(Rest <> "", " " & Rest, "")
                  Else
                    TU vsock, "5*** Sorry, there's no bot called '" & Param(Line, 2) & "'."
                  End If
                End If
              Else
                For u = 1 To SocketCount
                  If IsValidSocket(u) Then
                    RealNick = SocketItem(u).RegNick
                    If LCase(RealNick) <> BotNetNick Then
                      Select Case GetSockFlag(u, SF_Status)
                        Case SF_Status_Bot, SF_Status_BotLinking
                            TU vsock, "3*** Breaking link with " & RealNick
                            RTU u, "bye"
                            RemoveSocket u, 0, "Unlinked from", True
                            FoundOne = True
                        Case SF_Status_InitBotLink
                            TU vsock, "3*** Stopped trying to link to " & RealNick
                            RTU u, "bye"
                            RemoveSocket u, 0, "", True
                            FoundOne = True
                      End Select
                    End If
                  End If
                Next u
                BotCount = 0
                ReDim Preserve Bots(5)
                AddBot BotNetNick, "", "", LongToBase64(LongBotVersion), 0
                TU vsock, "3*** Unlinked from all bots."
              End If
          Case ".killnotes"
              If Param(Line, 2) <> "all" Then TU vsock, "5*** This erases ALL notes stored on the bot. Are you sure?": TU vsock, "5*** If you're really sure, type '.killnotes all'.": Exit Sub
              On Local Error Resume Next
              NotesKill
              On Error GoTo 0
              If Err.Number > 0 Then Err.Clear
              TU vsock, "3*** All notes stored on the bot were deleted."
          Case ".botnick"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".botnick <new nick>"): Exit Sub
              If Len(Param(Line, 2)) > ServerNickLen Then TU vsock, "5*** A nick can't be longer than " & CStr(ServerNickLen) & " characters!": Exit Sub
              If IsValidNick(Param(Line, 2)) = False Then TU vsock, "5*** """ & Param(Line, 2) & """ Erroneous Nickname": Exit Sub
              SendLine "NICK " & Param(Line, 2), 2
              TU vsock, "3*** Trying to change my nick to 10" & Param(Line, 2) & "3..."
          Case ".save"
              TU vsock, "5*** Hey, this is no eggdrop! I'm saving automatically =)"
          Case ".realname", ".newident", ".botnetnick", ".primarynick", ".prinick", ".secondarynick", ".secnick", ".ident", ".killprot"
              TU vsock, "5*** This command is outdated. It can be reached via '.botsetup' now."
          Case ".chaddr"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".chaddr <bot nick> (new address:port)"): Exit Sub
              UserNum = GetUserNum(Param(Line, 2))
              If UserNum = 0 Then TU vsock, "5*** Sorry, I couldn't find this bot.": Exit Sub
              If MatchFlags(BotUsers(UserNum).Flags, "-b") Then TU vsock, "5*** Sorry, " & BotUsers(UserNum).Name & " is not a bot (+b flag is missing).": Exit Sub
              If Param(Line, 3) = "" Then
                TU vsock, "3*** Current address of " & BotUsers(UserNum).Name & ":10 " & GetUserData(UserNum, UD_LinkAddr, "")
                Exit Sub
              End If
              If InStr(Param(Line, 3), ":") = 0 Then TU vsock, "5*** Please specify the new address like this: 'Host.Domain:Port'": Exit Sub
              If (InStr(Param(Line, 3), "*") > 0) Or (InStr(Param(Line, 3), "!") > 0) Or (InStr(Param(Line, 3), "?") > 0) Or (InStr(Param(Line, 3), "@") > 0) Then TU vsock, "5*** The connect address may not contain the following characters: *, !, ?, @": Exit Sub
              SetUserData UserNum, UD_LinkAddr, Param(Line, 3)
              TU vsock, "3*** Changed address of " & BotUsers(UserNum).Name & " to '10" & Param(Line, 3) & "3'."
              SharingSpreadMessage SocketItem(vsock).RegNick, "cmd .chaddr " & Param(Line, 2) & " " & Param(Line, 3)
          Case ".uptime"
              TU vsock, "3*** Bot started: " & TimeSpan(CStr(StartUpTime))
              If Connected Then TU vsock, "3*** Got connect: " & TimeSpan(CStr(ConnectTime))
          Case ".motd"
              If (Param(Line, 2) = "") Or (LCase(Param(Line, 2)) = LCase(BotNetNick)) Then
                If Not ShowMOTD(vsock) Then TU vsock, "5*** Sorry, no MOTD file found."
              Else
                If GetNextBot(Param(Line, 2)) = 0 Then TU vsock, "5*** Sorry, there's no bot called '" & Param(Line, 2) & "'."
                SendToBot Param(Line, 2), "m !" & CStr(vsock) & ":" & Nick & "@" & BotNetNick & " " & Param(Line, 2)
              End If
          Case ".floodprot"
              If Param(Line, 2) = "" Then
                TU vsock, "3*** My current flood protection (max bytes / 5 secs) is: 10" & CStr(MaxBytesToServer)
                TU vsock, "3*** Type '.floodprot <new max bytes/sec>' to change it."
                Exit Sub
              End If
              If Val(Param(Line, 2)) < 80 Then TU vsock, "5*** Sorry, this value is too low.": Exit Sub
              MaxBytesToServer = Val(Param(Line, 2))
              WritePPString "Server", "FloodProtection", CStr(MaxBytesToServer), AnGeL_INI
              TU vsock, "3*** The flood protection was set to: 10" & CStr(MaxBytesToServer) & " bytes/sec"
          Case ".invites"
              ListInvites vsock, Line
          Case ".+invite"
              AddInvite vsock, Line
          Case ".-invite"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-invite <hostmask or number>"): Exit Sub
              For u = 1 To InviteCount
                If LCase(Invites(u).Hostmask) = LCase(Param(Line, 2)) Or u = Val(Param(Line, 2)) Then
                  If (Invites(u).Channel = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to remove global invites.": Exit Sub
                  If MatchFlags(GetUserChanFlags(Nick, Invites(u).Channel), "-m") Then TU vsock, "5*** Sorry, you're not allowed to remove invites in this channel.": Exit Sub
                  DeletePPString Invites(u).Hostmask, "", HomeDir & "Invites.ini"
                  TU vsock, "3*** Removed " & IIf(Invites(u).Channel = "*", "global", Invites(u).Channel & " channel") & " invite '10" & Invites(u).Hostmask & "3'" & IIf(LCase(Param(Line, 1)) = ".-sinvite", " (silent, no immediate -b in channels)", "") & "."
                  If LCase(Param(Line, 1)) <> ".-sinvite" Then
                    For u2 = 1 To ChanCount
                      SearchIn = ""
                      For u3 = 1 To Channels(u2).InviteCount
                        SearchIn = SearchIn + vbCrLf + Channels(u2).InviteList(u3).Mask + vbCrLf
                      Next u3
                      If Invites(u).Channel = "*" Or LCase(Invites(u).Channel) = LCase(Channels(u2).Name) Then
                        If InStr(SearchIn, vbCrLf + Invites(u).Hostmask + vbCrLf) <> 0 Then SendLine "mode " & Channels(u2).Name & " -b " & Invites(u).Hostmask, 2
                      End If
                    Next u2
                  End If
                  ReadInvites
                  FoundOne = True
                  Exit For
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, I couldn't find this invitation."
          Case ".excepts"
              ListExcepts vsock, Line
          Case ".+except"
              AddExcept vsock, Line
          Case ".-except"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-except <hostmask or number>"): Exit Sub
              For u = 1 To ExceptCount
                If LCase(Excepts(u).Hostmask) = LCase(Param(Line, 2)) Or u = Val(Param(Line, 2)) Then
                  If (Excepts(u).Channel = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to remove global excepts.": Exit Sub
                  If MatchFlags(GetUserChanFlags(Nick, Excepts(u).Channel), "-m") Then TU vsock, "5*** Sorry, you're not allowed to remove excepts in this channel.": Exit Sub
                  DeletePPString Excepts(u).Hostmask, "", HomeDir & "Excepts.ini"
                  TU vsock, "3*** Removed " & IIf(Excepts(u).Channel = "*", "global", Excepts(u).Channel & " channel") & " except '10" & Excepts(u).Hostmask & "3'" & IIf(LCase(Param(Line, 1)) = ".-sexcept", " (silent, no immediate -b in channels)", "") & "."
                  If LCase(Param(Line, 1)) <> ".-sexcept" Then
                    For u2 = 1 To ChanCount
                      SearchIn = ""
                      For u3 = 1 To Channels(u2).ExceptCount
                        SearchIn = SearchIn + vbCrLf + Channels(u2).ExceptList(u3).Mask + vbCrLf
                      Next u3
                      If Excepts(u).Channel = "*" Or LCase(Excepts(u).Channel) = LCase(Channels(u2).Name) Then
                        If InStr(SearchIn, vbCrLf + Excepts(u).Hostmask + vbCrLf) <> 0 Then SendLine "mode " & Channels(u2).Name & " -b " & Excepts(u).Hostmask, 2
                      End If
                    Next u2
                  End If
                  ReadExcepts
                  FoundOne = True
                  Exit For
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, I couldn't find this exception."
          Case ".bans"
              ListBans vsock, Line
          Case ".+ban", ".+sban"
              AddBan vsock, Line
          Case ".-ban", ".-sban"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-ban <hostmask or number>"): Exit Sub
              For u = 1 To BanCount
                If LCase(Bans(u).Hostmask) = LCase(Param(Line, 2)) Or u = Val(Param(Line, 2)) Then
                  If (Bans(u).Channel = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to remove global bans.": Exit Sub
                  If MatchFlags(GetUserChanFlags(Nick, Bans(u).Channel), "-m") Then TU vsock, "5*** Sorry, you're not allowed to remove bans in this channel.": Exit Sub
                  DeletePPString Bans(u).Hostmask, "", HomeDir & "Bans.ini"
                  TU vsock, "3*** Removed " & IIf(Bans(u).Channel = "*", "global", Bans(u).Channel & " channel") & " ban '10" & Bans(u).Hostmask & "3'" & IIf(LCase(Param(Line, 1)) = ".-sban", " (silent, no immediate -b in channels)", "") & "."
                  If LCase(Param(Line, 1)) <> ".-sban" Then
                    For u2 = 1 To ChanCount
                      SearchIn = ""
                      For u3 = 1 To Channels(u2).BanCount
                        SearchIn = SearchIn + vbCrLf + Channels(u2).BanList(u3).Mask + vbCrLf
                      Next u3
                      If Bans(u).Channel = "*" Or LCase(Bans(u).Channel) = LCase(Channels(u2).Name) Then
                        If InStr(SearchIn, vbCrLf + Bans(u).Hostmask + vbCrLf) <> 0 Then SendLine "mode " & Channels(u2).Name & " -b " & Bans(u).Hostmask, 2
                      End If
                    Next u2
                  End If
                  ReadBans
                  FoundOne = True
                  Exit For
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, I couldn't find this ban."
          Case ".stick"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".stick <hostmask or number>"): Exit Sub
              For u = 1 To BanCount
                If LCase(Bans(u).Hostmask) = LCase(Param(Line, 2)) Or u = Val(Param(Line, 2)) Then
                  If (Bans(u).Channel = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to modify global bans.": Exit Sub
                  If MatchFlags(GetUserChanFlags(Nick, Bans(u).Channel), "-m") Then TU vsock, "5*** Sorry, you're not allowed to modify bans in this channel.": Exit Sub
                  WritePPString Bans(u).Hostmask, "Sticky", "yes", HomeDir & "Bans.ini"
                  Bans(u).Sticky = True
                  TU vsock, "3*** Stuck " & IIf(Bans(u).Channel = "*", "global", Bans(u).Channel & " channel") & " ban '10" & Bans(u).Hostmask & "3'."
                  CheckBans IIf(Bans(u).Channel = "*", "", Bans(u).Channel)
                  FoundOne = True
                  Exit For
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, I couldn't find this ban."
          Case ".unstick"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".unstick <hostmask or number>"): Exit Sub
              For u = 1 To BanCount
                If LCase(Bans(u).Hostmask) = LCase(Param(Line, 2)) Or u = Val(Param(Line, 2)) Then
                  If (Bans(u).Channel = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to modify global bans.": Exit Sub
                  If MatchFlags(GetUserChanFlags(Nick, Bans(u).Channel), "-m") Then TU vsock, "5*** Sorry, you're not allowed to modify bans in this channel.": Exit Sub
                  WritePPString Bans(u).Hostmask, "Sticky", "no", HomeDir & "Bans.ini"
                  Bans(u).Sticky = False
                  TU vsock, "3*** Unstuck " & IIf(Bans(u).Channel = "*", "global", Bans(u).Channel & " channel") & " ban '10" & Bans(u).Hostmask & "3'."
                  FoundOne = True
                  Exit For
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, I couldn't find this ban."
          Case ".ignores"
              If IgnoreCount = 0 Then TU vsock, "5*** I'm not ignoring anybody.": Exit Sub
              If GetSockFlag(vsock, SF_Colors) = SF_YES Then
                TU vsock, "2*** Currently ignoring:"
                TU vsock, "0,1 Nr: | Hostmask:                              | Created by: "
              Else
                TU vsock, "*** Currently ignoring:"
                TU vsock, " Nr:   Hostmask:                                Created by: "
                TU vsock, " ----  ---------------------------------------  ------------"
              End If
              For u = 1 To IgnoreCount
                If u > UBound(Ignores()) Then Exit For
                Messg = Ignores(u).Hostmask
                If Len(Messg) > 38 Then Messg = Left(Messg, 35) & "..."
                If GetSockFlag(vsock, SF_Colors) = SF_YES Then TU vsock, String(5 - Len(CStr(u)), " ") + CStr(u) & "14 | " & Messg + String(38 - Len(Messg), " ") & "14 | " & Ignores(u).CreatedBy
                If GetSockFlag(vsock, SF_Colors) = SF_NO Then TU vsock, String(5 - Len(CStr(u)), " ") + CStr(u) & "   " & Messg + String(38 - Len(Messg), " ") & "   " & Ignores(u).CreatedBy
              Next u
              TU vsock, EmptyLine
          Case ".+ignore"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+ignore <hostmask> (comment)"): Exit Sub
              If InStr(Param(Line, 2), "!") = 0 Or InStr(Param(Line, 2), "@") = 0 Then TU vsock, "5*** A valid hostmask must look like this: nick!identd@host.domain": Exit Sub
              For u = 1 To IgnoreCount
                If u > UBound(Ignores()) Then Exit For
                If LCase(Ignores(u).Hostmask) = LCase(Param(Line, 2)) Then TU vsock, "5*** This hostmask is already being ignored.": Exit Sub
              Next u
              WritePPString Param(Line, 2), "CreatedAt", Now, HomeDir & "Ignores.ini"
              WritePPString Param(Line, 2), "CreatedBy", Nick, HomeDir & "Ignores.ini"
              IgnoreCount = IgnoreCount + 1: If IgnoreCount > UBound(Ignores()) Then ReDim Preserve Ignores(UBound(Ignores()) + 5)
              Ignores(IgnoreCount).Hostmask = Param(Line, 2)
              Ignores(IgnoreCount).CreatedAt = Now
              Ignores(IgnoreCount).CreatedBy = Nick
              TU vsock, "3*** Added '10" & Param(Line, 2) & "3' to the ignore list."
          Case ".-ignore"
              If Param(Line, 2) = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".-ignore <hostmask or number>"): Exit Sub
              For u = 1 To IgnoreCount
                If u > UBound(Ignores()) Then Exit For
                If LCase(Ignores(u).Hostmask) = LCase(Param(Line, 2)) Or u = Val(Param(Line, 2)) Then
                  DeletePPString Ignores(u).Hostmask, "", HomeDir & "Ignores.ini"
                  TU vsock, "3*** Removed '10" & Ignores(u).Hostmask & "3' from the ignore list."
                  ReadIgnores
                  FoundOne = True
                  Exit For
                End If
              Next u
              If Not FoundOne Then TU vsock, "5*** Sorry, I couldn't find this hostmask on the ignore list."
          Case ".splits"
              ListSplits vsock
          Case ".clearsplits"
              SplitServerCount = 0: ReDim Preserve SplitServers(5)
              TU vsock, "Cleared splits."
          Case ".bottree", ".bt", ".botree", ".bottre", ".vbottree", ".vbt", ".vbotree", ".vbottre"
              If MatchFlags(Flags, "+t") Or MatchFlags(AllFlags(Nick), "+n") Then
                If BotCount > 1 Then
                  TU vsock, "2*** Bot Tree:"
                  DrawEbene 0, 1, "", vsock, (LCase(Left(Line, 2)) = ".v")
                  u4 = 1
                  For u = 1 To BotCount
                    u4 = u4 + CountHopsToBot(Bots(u).Nick)
                  Next u
                  u4 = u4 / BotCount
                  Rest = Format(u4, "0.0")
                  Mid(Rest, Len(Rest) - 1) = "."
                  TU vsock, "2*** Average hops: " & Rest & ", total bots: " & CStr(BotCount)
                  TU vsock, EmptyLine
                Else
                  TU vsock, "5*** There are no bots linked."
                End If
              End If
          Case ".help"
              ShowHelp vsock, Line
          Case ".servers"
              If GetPPString("Server", "Server", "", AnGeL_INI) = "" Then TU vsock, "5*** No servers - add one with '.+server <host>:<port>'.": Exit Sub
              TU vsock, "2*** Listing servers:"
              u = 1
              Do
                If u = 1 Then Rest = "Server" Else Rest = "Server" & CStr(u)
                Rest = GetPPString("Server", Rest, "", AnGeL_INI)
                If Rest = "" Then Exit Do
                TU vsock, " 2" & CStr(u) & ") " & Rest
                u = u + 1
              Loop
              TU vsock, EmptyLine
          Case ".+server"
              AddServer vsock, Line
          Case ".-server"
              RemServer vsock, Line
          Case ".run"
              If MatchFlags(Flags, "+s") Then
                If AllowRunToS = True Then
                  If Dir(GetRest(Line, 2)) <> "" Then
                    TU vsock, "5*** Application started... PID:" & Shell(GetRest(Line, 2), vbNormalFocus)
                  Else
                    TU vsock, "5*** .RUN failed... File not found"
                  End If
                End If
              End If
          Case ".userport", ".up"
              If Param(Line, 2) = "" Then
                TU vsock, "3*** My current user port: " & CStr(TelnetPort)
              ElseIf MatchFlags(Flags, "+s") Then
                TU vsock, "3*** Go to botsetup to change!"
              End If
          Case ".botport", ".bp"
              If Param(Line, 2) = "" Then
                TU vsock, "3*** My current bot port: " & CStr(BotnetPort)
              ElseIf MatchFlags(Flags, "+s") Then
                TU vsock, "3*** Go to botsetup to change!"
              End If
          Case ".bytes"
              TU vsock, "3*** Sent: " & SizeToString(GlobalBytesSent) & " | Received: " & SizeToString(GlobalBytesReceived)
          Case ".traffic"
              GUITrafficGraph vsock
          Case ".scripts"
              ListScripts vsock
          Case ".+script"
              GUIAddScript vsock, Line
          Case ".-script"
              GUIRemScript vsock, Line
          Case ".xpfix"
              SetSockFlag vsock, SF_Echo, SF_NO
          Case ".reload", ".rescript"
              If Param(Line, 2) = "" Then TU vsock, "3*** Usage: .reload <script>": Exit Sub
              GUIRemScript vsock, Line
              GUIAddScript vsock, Line
          Case ".iscripts"
              On Local Error Resume Next
              If Dir(FileAreaHome & "Incoming\MSSCRIPT.OCX") <> "" Then
                Rest = Space(260): u = kernel32_GetSystemDirectoryA(Rest, 260)
                Rest = Left(Rest, u): If Right(Rest, 1) <> "\" Then Rest = Rest & "\"
                Name FileAreaHome & "Incoming\MSSCRIPT.OCX" As Rest & "MSScript.ocx"
                If Dir(Rest & "MSScript.ocx") = "" Then
                  TU vsock, "5*** Installation failed; couldn't move MSSCRIPT.OCX."
                Else
                  Err.Clear
                  Shell "regsvr32 /s " & Rest & "MSScript.ocx"
                  If Err.Number = 0 Then
                    TU vsock, "3*** Success!"
                  Else
                    If Dir(FileAreaHome & "Incoming\regsvr32.exe") <> "" Then
                      Err.Clear
                      u = FreeFile
                      Open FileAreaHome & "Incoming\regsvr32.exe" For Input As #u
                      If LOF(u) = 37136 Then
                        Close #u
                        Shell FileAreaHome & "Incoming\regsvr32 /s " & Rest & "MSScript.ocx"
                      Else
                        Close #u
                        TU vsock, "5*** Sorry, invalid REGSVR32.EXE!"
                        Exit Sub
                      End If
                    End If
                    If Err.Number = 0 Then
                      TU vsock, "3*** Success!"
                    Else
                      TU vsock, "5*** Installation failed - " & Err.Description
                      TU vsock, "5*** Send the file REGSVR32.EXE to the bot and try again."
                    End If
                  End If
                End If
              Else
                TU vsock, "5*** Please upload MSSCRIPT.OCX first."
              End If
        End Select
      Else
        'Spread User Messages
        If Param(Line, 1) <> "ACTION" Then
          For u = 1 To ParamCount(Line)
            If EncryptIt(Param(Line, u)) = BotUsers(UserNum).Password Then
              TU vsock, "4*** Shhht! NEVER enter your password on the party line!!!"
              TU vsock, "4*** Your message was not shown due to security reasons."
              Exit Sub
            End If
          Next u
          If Line <> "" Then
            If SocketItem(vsock).AwayMessage <> "" Then SetAway vsock, ""
            ToBotNet 0, "c " & Nick & "@" & BotNetNick & " " & LongToBase64(SocketItem(vsock).PLChannel) & " " & Line
            SpreadMessage vsock, SocketItem(vsock).PLChannel, MakeMsg(MSG_PLTalk, Nick, Line)
            SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
          End If
        Else
          If SocketItem(vsock).AwayMessage <> "" Then SetAway vsock, ""
          ToBotNet 0, "a " & Nick & "@" & BotNetNick & " " & LongToBase64(SocketItem(vsock).PLChannel) & " " & Left(Right(Line, Len(Line) - 8), Len(Line) - 9)
          SpreadMessage vsock, SocketItem(vsock).PLChannel, MakeMsg(MSG_PLAct, Nick, Left(Right(Line, Len(Line) - 8), Len(Line) - 9))
          SocketItem(vsock).NumOfServerEvents = SocketItem(vsock).NumOfServerEvents + 1
        End If
      End If
    Case SF_Status_SharingSetup
      'SharingSetup vsock, Line
    Case SF_Status_PersonalSetup
      PersonalSetup vsock, Line
    Case SF_Status_ChanSetup
      ChanSetup vsock, Line
    Case SF_Status_BotSetup
      BotSetup vsock, Line
    Case SF_Status_KISetup
      KISetup vsock, Line
    Case SF_Status_NETSetup
      Netsetup vsock, Line
    Case SF_Status_POLSetup
      PolSetup vsock, Line
    Case SF_Status_AUTHSetup
      AUTHsetup vsock, Line
    Case SF_Status_FileArea
      FileArea vsock, Line
  End Select
  
  'Flood check
  If SocketItem(vsock).NumOfServerEvents > 20 Then
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If LCase(SocketItem(u).RegNick) = LCase(Nick) Then
          If GetSockFlag(u, SF_Status) = SF_Status_Party Then
            TU u, "4*** You were killed for flooding."
            ToBotNet 0, "pt " & BotNetNick & " " & SocketItem(u).RegNick & " " & SocketItem(u).OrderSign & " Killed for flooding"
            SpreadMessage u, SocketItem(u).PLChannel, "3*** " & SocketItem(u).RegNick & " was killed for flooding"
            WriteSeenEntry SocketItem(u).RegNick, "", Now, "*mine*", "*partyline*", Mask(SocketItem(u).Hostmask, 10)
            RemoveSocket u, 0, "", True
          End If
        End If
      End If
    Next u
  End If
End Sub

