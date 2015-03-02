Attribute VB_Name = "Server_BanList"
',-======================- ==-- -  -
'|   AnGeL - Server - BanList
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public Type Ban
  Hostmask As String
  Channel As String
  CreatedAt As Date
  CreatedBy As String
  ExpiresAt As Currency
  Comment As String
  Sticky As Boolean
End Type


Public BanCount As Long
Public Bans() As Ban
Public InviteCount As Long
Public Invites() As Invite
Public ExceptCount As Long
Public Excepts() As Except


Sub BanList_Load()
  ReDim Preserve Bans(5)
  ReDim Preserve Excepts(5)
  ReDim Preserve Invites(5)
End Sub


Sub BanList_Unload()
'
End Sub


Sub AddBan(vsock As Long, Line As String)
  Dim u As Long, u2 As Long, ToBan As String, Rest As String, Nick As String
  Dim Flags As String, Messg As String
  Nick = SocketItem(vsock).RegNick
  Flags = SocketItem(vsock).Flags
  ToBan = Param(Line, 2)
  If ToBan = "" Then TU vsock, MakeMsg(ERR_CommandUsage, ".+(s)ban <hostmask> ([" & ServerChannelPrefixes & "]channel) (comment)"): Exit Sub
  If IsValidHostmask(ToBan) = False Then TU vsock, "5*** A valid hostmask must look like this: nick!identd@host.domain": Exit Sub
  If Len(Mask(ToBan, 12)) > 10 Then TU vsock, "5*** Sorry, the ident is too long (longer than 10 characters).": Exit Sub
  If Len(Mask(ToBan, 13)) > ServerNickLen Then TU vsock, "5*** Sorry, the nick is too long (longer than " & CStr(ServerNickLen) & " characters).": Exit Sub
  If IsValidChannel(Left(Param(Line, 3), 1)) Then Rest = Param(Line, 3) Else Rest = "*"
  If (Rest = "*") And MatchFlags(Flags, "-m") Then TU vsock, "5*** Sorry, you're not allowed to place global bans. Use: .+ban <hostmask> <#YOUR CHANNEL> (comment)": Exit Sub
  If MatchFlags(GetUserChanFlags(Nick, Rest), "-m") Then TU vsock, "5*** Sorry, you're not allowed to place bans in this channel.": Exit Sub
  If MatchWM(ToBan, MyHostmask) Or MatchWM(ToBan, MyIPmask) Then
    TU vsock, "5*** Sure. I'll ban myself if you pay $500000 to Hippo first.": Exit Sub
  End If
  For u = 1 To BotUserCount
    For u2 = 1 To BotUsers(u).HostMaskCount
      If MatchWM(ToBan, BotUsers(u).HostMasks(u2)) Or MatchWM(BotUsers(u).HostMasks(u2), ToBan) Then
        If MatchFlags(BotUsers(u).Flags, "+b") Then
          TU vsock, "5*** Sorry, this banmask is matching the bot '" & BotUsers(u).Name & "'.": Exit Sub
        End If
        If MatchFlags(BotUsers(u).Flags, "+n") Then
          TU vsock, "5*** Sorry, this banmask is matching the owner '" & BotUsers(u).Name & "'.": Exit Sub
        End If
      End If
    Next u2
  Next u
  
  For u = 1 To BanCount
    If (LCase(Bans(u).Hostmask) = LCase(ToBan)) And (LCase(Bans(u).Channel) = LCase(Rest)) Then TU vsock, "5*** This hostmask is already banned.": Exit Sub
  Next u
  If IsValidChannel(Left(Param(Line, 3), 1)) Then
    If Param(Line, 4) <> "" Then Messg = Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " " & Param(Line, 3) & " ")) Else Messg = "requested"
  Else
    If Param(Line, 3) <> "" Then Messg = Right(Line, Len(Line) - Len(Param(Line, 1) & " " & Param(Line, 2) & " ")) Else Messg = "requested"
  End If
  WritePPString Param(Line, 2), "Channel", Rest, HomeDir & "Bans.ini"
  WritePPString Param(Line, 2), "CreatedAt", Now, HomeDir & "Bans.ini"
  WritePPString Param(Line, 2), "CreatedBy", Nick, HomeDir & "Bans.ini"
  WritePPString Param(Line, 2), "Comment", Messg, HomeDir & "Bans.ini"
  WritePPString Param(Line, 2), "Sticky", "no", HomeDir & "Bans.ini"
  BanCount = BanCount + 1: If BanCount > UBound(Bans()) Then ReDim Preserve Bans(UBound(Bans()) + 5)
  Bans(BanCount).Hostmask = Param(Line, 2)
  Bans(BanCount).Channel = Rest
  Bans(BanCount).CreatedAt = Now
  Bans(BanCount).CreatedBy = Nick
  Bans(BanCount).Comment = Messg
  Bans(BanCount).Sticky = False
  If LCase(Param(Line, 1)) <> ".+sban" Then
    For u = 1 To ChanCount
      If Channels(u).GotOPs Or Channels(u).GotHOPs Then
        If (Bans(BanCount).Channel = "*") Or (LCase(Bans(BanCount).Channel) = LCase(Channels(u).Name)) Then
          For u2 = 1 To Channels(u).UserCount
            If MatchWM(Bans(BanCount).Hostmask, Channels(u).User(u2).Hostmask) And (Channels(u).User(u2).Nick <> MyNick) Then
              If InStr(Channels(u).User(u2).Status, "@") > 0 Then AddMassMode2 "-o", Channels(u).User(u2).Nick
            End If
          Next u2
          If Not IsBanned(u, Bans(BanCount).Hostmask) Then
            If Not IsOrdered("ban " & Channels(u).Name & " " & Bans(BanCount).Hostmask) Then
              Order "ban " & Channels(u).Name & " " & Bans(BanCount).Hostmask, 30
              AddMassMode2 "+b", Bans(BanCount).Hostmask
            End If
          End If
          DoMassMode2 Channels(u).Name
          For u2 = 1 To Channels(u).UserCount
            If MatchWM(Bans(BanCount).Hostmask, Channels(u).User(u2).Hostmask) And (Channels(u).User(u2).Nick <> MyNick) Then
              AddKickUser u, Channels(u).User(u2).Nick, Bans(BanCount).Hostmask, "Banned: " & Messg
            End If
          Next u2
        End If
      End If
    Next u
  End If
  TU vsock, "3*** Added " & IIf(Rest = "*", "global", Rest & " channel") & " ban '10" & Param(Line, 2) & "3' as number " & CStr(BanCount) & "" & IIf(LCase(Param(Line, 1)) = ".+sban", " (silent, no immediate +b in channels)", "") & "."
End Sub


Sub ReadBans()
  Dim FileNum As Integer, Line As String
  If Dir(HomeDir & "Bans.ini") = "" Then Exit Sub
  On Local Error Resume Next
  BanCount = 0
  FileNum = FreeFile
  Open HomeDir & "Bans.ini" For Input As #FileNum
  If Err.Number = 0 Then
    While Not EOF(FileNum)
      Line Input #FileNum, Line
      Line = Trim(Line)
      If Left(Line, 1) = "[" And Right(Line, 1) = "]" Then
        Line = MakeNormalNick(Mid(Line, 2, Len(Line) - 2))
        BanCount = BanCount + 1
        If BanCount > UBound(Bans()) Then ReDim Preserve Bans(UBound(Bans()) + 5)
        Bans(BanCount).Hostmask = Line
        Bans(BanCount).Channel = GetPPString(Line, "Channel", "", HomeDir & "Bans.ini")
        If Bans(BanCount).Channel <> "" Then
          Bans(BanCount).CreatedAt = GetPPString(Line, "CreatedAt", "", HomeDir & "Bans.ini")
          Bans(BanCount).CreatedBy = GetPPString(Line, "CreatedBy", "", HomeDir & "Bans.ini")
          Bans(BanCount).Comment = GetPPString(Line, "Comment", "", HomeDir & "Bans.ini")
          Bans(BanCount).Sticky = (GetPPString(Line, "Sticky", "", HomeDir & "Bans.ini") = "yes")
        Else
          BanCount = BanCount - 1
        End If
      End If
    Wend
    Close #FileNum
  Else
    Err.Clear
  End If
End Sub


Sub ReadInvites()
  Dim FileNum As Integer, Line As String
  If Dir(HomeDir & "Invites.ini") = "" Then Exit Sub
  On Local Error Resume Next
  InviteCount = 0
  FileNum = FreeFile
  Open HomeDir & "Invites.ini" For Input As #FileNum
  If Err.Number = 0 Then
    While Not EOF(FileNum)
      Line Input #FileNum, Line
      Line = Trim(Line)
      If Left(Line, 1) = "[" And Right(Line, 1) = "]" Then
        Line = MakeNormalNick(Mid(Line, 2, Len(Line) - 2))
        InviteCount = InviteCount + 1
        If InviteCount > UBound(Invites()) Then ReDim Preserve Invites(UBound(Invites()) + 5)
        Invites(InviteCount).Hostmask = Line
        Invites(InviteCount).Channel = GetPPString(Line, "Channel", "", HomeDir & "Invites.ini")
        If Invites(InviteCount).Channel <> "" Then
          Invites(InviteCount).CreatedAt = GetPPString(Line, "CreatedAt", "", HomeDir & "Invites.ini")
          Invites(InviteCount).CreatedBy = GetPPString(Line, "CreatedBy", "", HomeDir & "Invites.ini")
        Else
          InviteCount = InviteCount - 1
        End If
      End If
    Wend
    Close #FileNum
  Else
    Err.Clear
  End If
End Sub

Sub ReadExcepts()
  Dim FileNum As Integer, Line As String
  If Dir(HomeDir & "Excepts.ini") = "" Then Exit Sub
  On Local Error Resume Next
  ExceptCount = 0
  FileNum = FreeFile
  Open HomeDir & "Excepts.ini" For Input As #FileNum
  If Err.Number = 0 Then
    While Not EOF(FileNum)
      Line Input #FileNum, Line
      Line = Trim(Line)
      If Left(Line, 1) = "[" And Right(Line, 1) = "]" Then
        Line = MakeNormalNick(Mid(Line, 2, Len(Line) - 2))
        ExceptCount = ExceptCount + 1
        If ExceptCount > UBound(Excepts()) Then ReDim Preserve Excepts(UBound(Excepts()) + 5)
        Excepts(ExceptCount).Hostmask = Line
        Excepts(ExceptCount).Channel = GetPPString(Line, "Channel", "", HomeDir & "Excepts.ini")
        If Excepts(ExceptCount).Channel <> "" Then
          Excepts(ExceptCount).CreatedAt = GetPPString(Line, "CreatedAt", "", HomeDir & "Excepts.ini")
          Excepts(ExceptCount).CreatedBy = GetPPString(Line, "CreatedBy", "", HomeDir & "Excepts.ini")
        Else
          ExceptCount = ExceptCount - 1
        End If
      End If
    Wend
    Close #FileNum
  Else
    Err.Clear
  End If
End Sub


Public Function BannedMask(Hostmask As String, Channel As String) As String ' : AddStack "Server_BannedMask(" & Hostmask & ", " & Channel & ")"
Dim u As Long
  For u = 1 To BanCount
    If Bans(u).Channel = "*" Or LCase(Bans(u).Channel) = LCase(Channel) Then
      If MatchWM(Bans(u).Hostmask, Hostmask) Then BannedMask = Bans(u).Hostmask:  Exit Function
    End If
  Next u
  BannedMask = ""
End Function


Public Sub RemDisturbingBans(ChNum As Long) ' : AddStack "ServerRoutines_RemDisturbingBans(" & ChNum & ")"
  Dim u2 As Long, u3 As Long, KUserFlags As String, ExistingBan As Boolean
  If (Channels(ChNum).DidUnban = False) And (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) And (Channels(ChNum).CompletedBANS = True) Then
    For u2 = 1 To Channels(ChNum).BanCount
      'Remove bans matching myself
      If (MatchWM(Channels(ChNum).BanList(u2).Mask, MyHostmask) = True) Or (MatchWM(Channels(ChNum).BanList(u2).Mask, MyIPmask) = True) Then
        AddMassMode "-b", Channels(ChNum).BanList(u2).Mask
      Else
        'Remove bans matching bots, owners and superowners
        KUserFlags = GetUserChanFlags(SearchUserFromHostmask2(Channels(ChNum).BanList(u2).Mask), Channels(ChNum).Name)
        If MatchFlags(KUserFlags, "+m") Or MatchFlags(KUserFlags, "+b") Then
          ExistingBan = False
          For u3 = 1 To BanCount
            If (Bans(u3).Channel = "*") Or (LCase(Bans(u3).Channel) = LCase(Channels(ChNum).Name)) Then
              If MatchWM(Bans(u3).Hostmask, Channels(ChNum).BanList(u2).Mask) Then
                ExistingBan = True: Exit For
              End If
            End If
          Next u3
          If ExistingBan = False Then AddMassMode "-b", Channels(ChNum).BanList(u2).Mask
        End If
      End If
    Next u2
    DoMassMode Channels(ChNum).Name
    Channels(ChNum).DidUnban = True
  End If
End Sub

Public Sub RemDisturbingExcepts(ChNum As Long) ' : AddStack "ServerRoutines_RemDisturbingExcepts(" & ChNum & ")"
'  Dim u2 As Long, u3 As Long, KUserFlags As String, ExistingExcept As Boolean
'  If (Channels(ChNum).DidUnExcept = False) And (Channels(ChNum).GotOPs = True) And (Channels(ChNum).CompletedExcepts = True) Then
'    For u2 = 1 To Channels(ChNum).ExceptCount
'      For u3 = 1 To ExceptCount
'        If (Excepts(u3).Channel = "*") Or (LCase(Excepts(u3).Channel) = LCase(Channels(ChNum).Name)) Then
'          If MatchWM(Excepts(u3).Hostmask, Channels(ChNum).ExceptList(u2).Mask) Then
'            ExistingExcept = True: Exit For
'          End If
'        End If
'      Next u3
'      If ExistingExcept = False Then AddMassMode "-e", Channels(ChNum).ExceptList(u2).Mask
'    Next u2
'    DoMassMode Channels(ChNum).Name
'    Channels(ChNum).DidUnExcept = True
'  End If
End Sub
Public Sub RemDisturbingInvites(ChNum As Long) ' : AddStack "ServerRoutines_RemDisturbingInvites(" & ChNum & ")"
'  Dim u2 As Long, u3 As Long, KUserFlags As String, ExistingInvite As Boolean
'  If (Channels(ChNum).DidUnInvite = False) And (Channels(ChNum).GotOPs = True) And (Channels(ChNum).CompletedInvites = True) Then
'    For u2 = 1 To Channels(ChNum).InviteCount
'      For u3 = 1 To InviteCount
'        If (Invites(u3).Channel = "*") Or (LCase(Invites(u3).Channel) = LCase(Channels(ChNum).Name)) Then
'          If MatchWM(Invites(u3).Hostmask, Channels(ChNum).InviteList(u2).Mask) Then
'            ExistingInvite = True: Exit For
'          End If
'        End If
'      Next u3
'      If ExistingInvite = False Then AddMassMode "-I", Channels(ChNum).InviteList(u2).Mask
'    Next u2
'    DoMassMode Channels(ChNum).Name
'    Channels(ChNum).DidUnInvite = True
'  End If
End Sub

'Check if a user matches any permanent bans in the bot
Public Function CheckPermBans(ChNum As Long, UsNum As Long) As Boolean ' : AddStack "ServerRoutines_CheckPermBans(" & ChNum & ", " & UsNum & ")"
Dim u As Long, BanNum As Long, MatchedIP As Boolean
  If (ChNum = 0) Or (UsNum = 0) Then CheckPermBans = False: Exit Function
  'Avoid self-bans
  If LCase(Channels(ChNum).User(UsNum).Nick) = LCase(MyNick) Then CheckPermBans = False: Exit Function
  
  'Search ban list for matching bans
  BanNum = 0
  MatchedIP = False
  For u = 1 To BanCount
    If (Bans(u).Channel = "*") Or (LCase(Bans(u).Channel) = LCase(Channels(ChNum).Name)) Then
      If MatchWM(Bans(u).Hostmask, Channels(ChNum).User(UsNum).Hostmask) Then
        BanNum = u
        Exit For
      ElseIf MatchWM(Bans(u).Hostmask, Channels(ChNum).User(UsNum).IPmask) Then
        BanNum = u: MatchedIP = True
        Exit For
      End If
    End If
  Next u
  
  'Found nothing
  If BanNum = 0 Then CheckPermBans = False: Exit Function
  
  'Found a matching ban, ban+kick user
  If AddToBan(ChNum, Bans(BanNum).Hostmask) Then FixTimedEvent "UnBan " & Channels(ChNum).Name & " " & Bans(BanNum).Hostmask, UnbanTime
  If MatchedIP Then
    AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).IPmask, "Banned: " & Bans(BanNum).Comment
  Else
    AddKickUser ChNum, Channels(ChNum).User(UsNum).Nick, Channels(ChNum).User(UsNum).Hostmask, "Banned: " & Bans(BanNum).Comment
  End If
  CheckPermBans = True
End Function
Public Sub CheckExcepts(Channel As String) ' : AddStack "ServerRoutines_CheckExcepts(" & Channel & ")"
  Dim u As Long, u2 As Long, SearchIn As String, TempStr As String
  If Channel = "" Then
    For u = 1 To ChanCount
      If Left(Channels(u).Name, 1) <> "&" Then
        SearchIn = ""
        For u2 = 1 To Channels(u).ExceptCount
          SearchIn = SearchIn + vbCrLf + Channels(u).ExceptList(u2).Mask + vbCrLf
        Next u2
        For u2 = 1 To ExceptCount
          If Excepts(u2).Channel = "*" Or LCase(Excepts(u2).Channel) = LCase(Channels(u).Name) Then
            If InStr(SearchIn, vbCrLf + Excepts(u2).Hostmask + vbCrLf) = 0 Then
              TempStr = Excepts(u2).Hostmask
              AddToExcept u, TempStr
            End If
          End If
        Next u2
      End If
    Next u
  Else
    u = FindChan(Channel)
    If u = 0 Then Exit Sub
    SearchIn = ""
    For u2 = 1 To Channels(u).ExceptCount
      SearchIn = SearchIn + vbCrLf + Channels(u).ExceptList(u2).Mask + vbCrLf
    Next u2
    For u2 = 1 To ExceptCount
      If Excepts(u2).Channel = "*" Or LCase(Excepts(u2).Channel) = LCase(Channels(u).Name) Then
        If InStr(SearchIn, vbCrLf + Excepts(u2).Hostmask + vbCrLf) = 0 Then
          TempStr = Excepts(u2).Hostmask
          AddToExcept u, TempStr
        End If
      End If
    Next u2
  End If
End Sub
Public Sub CheckInvites(Channel As String) ' : AddStack "ServerRoutines_CheckInvites(" & Channel & ")"
  Dim u As Long, u2 As Long, SearchIn As String, TempStr As String
  If Channel = "" Then
    For u = 1 To ChanCount
      If Left(Channels(u).Name, 1) <> "&" Then
        SearchIn = ""
        For u2 = 1 To Channels(u).InviteCount
          SearchIn = SearchIn + vbCrLf + Channels(u).InviteList(u2).Mask + vbCrLf
        Next u2
        For u2 = 1 To InviteCount
          If Invites(u2).Channel = "*" Or LCase(Invites(u2).Channel) = LCase(Channels(u).Name) Then
            If InStr(SearchIn, vbCrLf + Invites(u2).Hostmask + vbCrLf) = 0 Then
              TempStr = Invites(u2).Hostmask
              AddToInvite u, TempStr
            End If
          End If
        Next u2
      End If
    Next u
  Else
    u = FindChan(Channel)
    If u = 0 Then Exit Sub
    SearchIn = ""
    For u2 = 1 To Channels(u).InviteCount
      SearchIn = SearchIn + vbCrLf + Channels(u).InviteList(u2).Mask + vbCrLf
    Next u2
    For u2 = 1 To InviteCount
      If Invites(u2).Channel = "*" Or LCase(Invites(u2).Channel) = LCase(Channels(u).Name) Then
        If InStr(SearchIn, vbCrLf + Invites(u2).Hostmask + vbCrLf) = 0 Then
          TempStr = Invites(u2).Hostmask
          AddToInvite u, TempStr
        End If
      End If
    Next u2
  End If
End Sub

Public Sub CheckBans(Channel As String) ' : AddStack "ServerRoutines_CheckBans(" & Channel & ")"
Dim u As Long, u2 As Long, SearchIn As String, TempStr As String
  If Channel = "" Then
    For u = 1 To ChanCount
      If Left(Channels(u).Name, 1) <> "&" Then
        SearchIn = ""
        For u2 = 1 To Channels(u).BanCount
          SearchIn = SearchIn + vbCrLf + Channels(u).BanList(u2).Mask + vbCrLf
        Next u2
        For u2 = 1 To BanCount
          If Bans(u2).Channel = "*" Or LCase(Bans(u2).Channel) = LCase(Channels(u).Name) Then
            If Bans(u2).Sticky Then
              If InStr(SearchIn, vbCrLf + Bans(u2).Hostmask + vbCrLf) = 0 Then
                TempStr = Bans(u2).Hostmask
                AddToBan u, TempStr
              End If
            End If
          End If
        Next u2
      End If
    Next u
  Else
    u = FindChan(Channel)
    If u = 0 Then Exit Sub
    SearchIn = ""
    For u2 = 1 To Channels(u).BanCount
      SearchIn = SearchIn + vbCrLf + Channels(u).BanList(u2).Mask + vbCrLf
    Next u2
    For u2 = 1 To BanCount
      If Bans(u2).Channel = "*" Or LCase(Bans(u2).Channel) = LCase(Channels(u).Name) Then
        If Bans(u2).Sticky Then
          If InStr(SearchIn, vbCrLf + Bans(u2).Hostmask + vbCrLf) = 0 Then
            TempStr = Bans(u2).Hostmask
            AddToBan u, TempStr
          End If
        End If
      End If
    Next u2
  End If
End Sub

