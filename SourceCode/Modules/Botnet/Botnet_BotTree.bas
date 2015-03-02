Attribute VB_Name = "Botnet_BotTree"
',-======================- ==-- -  -
'|   AnGeL - Botnet - BotTree
'|   � 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public Type Bot
  Nick As String
  SharingFlag As String
  Version As String
  Ebene As Long
  SubBotOf As String
  vSocket As Long
  SentPing As Boolean
  GotPingReply As Boolean
  SendRequests As Boolean
End Type


Public BotCount As Long
Public Bots() As Bot


Sub BotTree_Load()
  ReDim Preserve Bots(5)
End Sub

Sub BotTree_Unload()
'
End Sub


Public Sub AddBot(Nick As String, SubBotOf As String, SharingFlag As String, Version As String, vSocket As Long) ' : AddStack "BotNet_AddBot(" & Nick & ", " & SubBotOf & ", " & SharingFlag & ", " & Version & ", " & SockNum & ")"
Dim u As Long, BEb As Long, BPos As Long
  If GetBotPos(Nick) > 0 Then SpreadFlagMessage 0, "+m", "14[" & Time & "] *** ERROR - bot is already in net: " & Nick: Exit Sub
  BotCount = BotCount + 1
  If BotCount > UBound(Bots()) Then ReDim Preserve Bots(UBound(Bots()) + 5)
  If SubBotOf = "" Then
    Bots(BotCount).Nick = Nick
    If SharingFlag <> "" Then Bots(BotCount).SharingFlag = SharingFlag Else Bots(BotCount).SharingFlag = "-"
    Bots(BotCount).Version = Version
    Bots(BotCount).Ebene = 0
    Bots(BotCount).SubBotOf = SubBotOf
    Bots(BotCount).vSocket = vSocket
    Bots(BotCount).SentPing = False
    Bots(BotCount).SendRequests = True
  Else
    BPos = GetBotPos(SubBotOf) + 1
    BEb = GetBotEbene(SubBotOf) + 1
    If BPos = 1 Then SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Fake link from non-existing " & SubBotOf & ": " & Nick: BotCount = BotCount - 1: Exit Sub
    For u = BPos To BotCount - 1
      If Bots(u).Ebene < BEb Then BPos = u: Exit For
    Next u
    For u = BotCount To BPos + 1 Step -1
      Bots(u) = Bots(u - 1)
    Next u
    Bots(BPos).Nick = Nick
    If SharingFlag <> "" Then Bots(BPos).SharingFlag = SharingFlag Else Bots(BPos).SharingFlag = "-"
    Bots(BPos).Version = Version
    Bots(BPos).Ebene = BEb
    Bots(BPos).SubBotOf = SubBotOf
    Bots(BPos).vSocket = vSocket
    Bots(BPos).SentPing = False
    Bots(BPos).SendRequests = True
  End If
End Sub

'Remove a bot from the botnet
Public Sub RemBot(ByVal Nick As String, ByVal vSocket As Long, Reason As String)
Dim u As Long, BEb As Long, BPos As Long, StopIt As Boolean, FoundOne As Boolean
Dim LostUsers As Long, LostBots As Long, SpreadUnlinkMSG As Boolean
  BPos = 1
  'Remove Bots
  For u = 2 To BotCount
    If Not StopIt Then BPos = BPos + 1
    If (StopIt = True) And (Bots(u).Ebene <= BEb) Then StopIt = False
    If LCase(Nick) = LCase(Bots(u).Nick) And ((vSocket > 0 And vSocket = Bots(u).vSocket) Or (vSocket = 0)) Then
      BEb = Bots(u).Ebene: StopIt = True
      If LCase(Bots(u).SubBotOf) = LCase(BotNetNick) Then SpreadUnlinkMSG = True Else SpreadUnlinkMSG = False
    End If
    If Not StopIt Then Bots(BPos) = Bots(u)
  Next u
  If BPos = 1 Then Exit Sub
  If StopIt Then BPos = BPos - 1
  LostBots = BotCount - BPos
  BotCount = BPos
  'ReDim Bot Array
  u = ((BotCount \ 5) + 1) * 5
  If u < UBound(Bots()) Then ReDim Preserve Bots(u)
  'Remove Users
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If GetSockFlag(u, SF_Status) = SF_Status_BotNetParty Then
        If GetBotPos(SocketItem(u).OnBot) = 0 Then LostUsers = LostUsers + 1: RemoveSocket u, 0, "", True
      End If
    End If
  Next u
  If SpreadUnlinkMSG Then
    If ParamXCount(Reason, "�") > 1 Then
      SpreadMessageEx 0, -1, SF_Local_Bot, "14[" & Time & "] 3*** " & ParamX(Reason, "�", 1) & " " & Nick & ": " & ParamX(Reason, "�", 2) & " (Lost " & CStr(LostBots) + IIf(LostBots = 1, " bot", " bots") & " & " & CStr(LostUsers) + IIf(LostUsers = 1, " user)", " users)")
      ToBotNet 0, "un " & Nick & " " & ParamX(Reason, "�", 1) & " " & Nick & ": " & ParamX(Reason, "�", 2) & " (Lost " & CStr(LostBots) + IIf(LostBots = 1, " bot", " bots") & " & " & CStr(LostUsers) + IIf(LostUsers = 1, " user)", " users)")
    Else
      SpreadMessageEx 0, -1, SF_Local_Bot, "14[" & Time & "] 3*** " & Reason & " " & Nick & " (Lost " & CStr(LostBots) + IIf(LostBots = 1, " bot", " bots") & " & " & CStr(LostUsers) + IIf(LostUsers = 1, " user)", " users)")
      ToBotNet 0, "un " & Nick & " " & Reason & " " & Nick & " (Lost " & CStr(LostBots) + IIf(LostBots = 1, " bot", " bots") & " & " & CStr(LostUsers) + IIf(LostUsers = 1, " user)", " users)")
    End If
  End If
  CheckBotLinks
End Sub

Public Sub CheckBotLinks() ' : AddStack "BotNet_CheckBotLinks()"
Dim u As Long, Nick As String, FoundOne As Boolean
  Do
    FoundOne = False
    For u = 2 To BotCount
      If CountHopsToBot(Bots(u).Nick) = 0 Then
        RemBot Bots(u).Nick, 0, "Not really connected:"
        FoundOne = True
        Exit For
      End If
    Next u
    If Not FoundOne Then Exit Do
  Loop
End Sub


Function DupeCheck(vsock As Long) As Boolean ' : AddStack "BotNet_DupeCheck(" & vsock & ")"
Dim u As Long
  'Reject link if same nick is already in botnet but not connected directly to me
  For u = 1 To BotCount
    If LCase(Bots(u).Nick) = LCase(SocketItem(vsock).RegNick) Then
      If LCase(Bots(u).SubBotOf) <> LCase(BotNetNick) Then
        SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Link from " & SocketItem(vsock).RegNick & " rejected: Bot is already in botnet"
        RTU vsock, "error You're already in this botnet!"
        RemoveSocket vsock, 0, "", True
        DupeCheck = True
        Exit Function
      Else
        Bots(u).GotPingReply = False
      End If
    End If
  Next u
  'If same nick is already in botnet and directly connected to me, ping old bot and wait 3 seconds
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If (u <> vsock) And (LCase(SocketItem(u).RegNick) = LCase(SocketItem(vsock).RegNick)) Then
        If GetSockFlag(u, SF_Status) = SF_Status_Bot Then
          TimedEvent "KickOldBot " & Trim(Str(vsock)) & " " & SocketItem(vsock).RegNick, 3
          SpreadFlagMessage vsock, "+t", MakeMsg(MSG_PLBotNetLinkFrom, SocketItem(vsock).RegNick, "Duplicate, pinging old bot...")
          RTU u, "pi"
          DupeCheck = True
          Exit Function
        End If
      End If
    End If
  Next u
  DupeCheck = False
End Function

Function StrictDupeCheck(vsock As Long) As Boolean ' : AddStack "BotNet_StrictDupeCheck(" & vsock & ")"
Dim u As Long
  'Reject link if same nick is already in botnet
  For u = 1 To BotCount
    If LCase(Bots(u).Nick) = LCase(SocketItem(vsock).RegNick) Then
      SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Link from " & SocketItem(vsock).RegNick & " rejected: Bot is already in botnet"
      RTU vsock, "error You're already in this botnet!"
      RemoveSocket vsock, 0, "", True
      StrictDupeCheck = True
      Exit Function
    End If
  Next u
  StrictDupeCheck = False
End Function

Sub KickOldBot(vsock As Long, RegNick As String) ' : AddStack "BotNet_KickOldBot(" & vsock & ", " & RegNick & ")"
Dim u As Long, BotPosition As Long
  If IsValidSocket(vsock) = False Then Exit Sub
  If SocketItem(vsock).RegNick <> RegNick Then Exit Sub
  BotPosition = GetBotPos(RegNick)
  'Kick old bot with the same nick if bot is connected directly to me
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If (u <> vsock) And (LCase(SocketItem(u).RegNick) = LCase(RegNick)) Then
        If GetSockFlag(u, SF_Status) = SF_Status_Bot Then
          If Bots(BotPosition).GotPingReply = False Then
            RemoveSocket u, 0, "Forced ping timeout:", True
          Else
            SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** Link from " & SocketItem(vsock).RegNick & " rejected: Old bot is still alive!"
            RTU vsock, "error A bot with your nick is already in this botnet!"
            RemoveSocket vsock, 0, "", True
            Exit Sub
          End If
        End If
      End If
    End If
  Next u
  
  'winsock2_send OK to other bot
  SetSockFlag vsock, SF_Status, SF_Status_BotLinking
  SetSockFlag vsock, SF_LoggedIn, SF_YES
  SetSockFlag vsock, SF_Silent, SF_YES
  RTU vsock, "*hello!"
  RTU vsock, "version " & LongBotVersion & " " & CStr(ServerNickLen) & " AnGeL " & BotVersion & " <" & ServerNetwork & ">"
End Sub

'winsock2_send something to all bots connected
Sub ToBotNet(vsock As Long, Line As String) ' : AddStack "BotNet_ToBotNet(" & vsock & ", " & Line & ")"
Dim u As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If vsock <> u And ((GetSockFlag(u, SF_Status) = SF_Status_Bot) Or (GetSockFlag(u, SF_Status) = SF_Status_BotPreCache)) And Not (LCase(Left(Line, 1)) = "c" And MatchFlags(BotUsers(SocketItem(u).UserNum).BotFlags, "+i") = True) Then
        RTU u, Line
      End If
    End If
  Next u
End Sub

'Retrieves the bot connected to me directly that is on the path to <Bot>
Public Function GetNextBot(Bot As String) As Long ' : AddStack "BotNet_GetNextBot(" & Bot & ")"
Dim u As Long, CurCheckBot As String, FoundOne As Boolean
  CurCheckBot = Bot
  Do
    FoundOne = False
    For u = 1 To BotCount
      If LCase(Bots(u).Nick) = LCase(CurCheckBot) Then
        If LCase(Bots(u).SubBotOf) = LCase(BotNetNick) Then GetNextBot = u: Exit Function
        CurCheckBot = Bots(u).SubBotOf
        FoundOne = True
        Exit For
      End If
    Next u
    If Not FoundOne Then Exit Do
  Loop
  GetNextBot = 0
End Function

Function FakeCheck(ByVal Nick As String, ByVal OnBot As String, vsock As Long) As Boolean ' : AddStack "BotNet_FakeCheck(" & Nick & ", " & OnBot & ", " & vsock & ")"
Dim u As Long, BP As Long, FoundOne As Boolean
  If Left(Nick, 1) = "!" Then Nick = Mid(Nick, 2)
  If Left(OnBot, 1) = "!" Then OnBot = Mid(OnBot, 2)
  If InStr(Nick, ":") > 0 Then Nick = Mid(Nick, InStr(Nick, ":") + 1)
  BP = GetBotPos(OnBot)
  If LCase(OnBot) = LCase(BotNetNick) Then
    RTU vsock, "error Fake message rejected (source != me)"
    SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Fake message from " & SocketItem(vsock).RegNick & " rejected (source != me)"
    FakeCheck = True
    Exit Function
  End If
  If BP = 0 Then
    RTU vsock, "error Fake message rejected (no such bot '" & OnBot & "'!)"
    SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Fake message from " & SocketItem(vsock).RegNick & " rejected (no such bot: '" & OnBot & "'!)"
    FakeCheck = True
    Exit Function
  End If
  If Bots(GetNextBot(OnBot)).vSocket <> vsock Then
    RTU vsock, "error Fake message rejected ('" & OnBot & "' is not connected via you!)"
    SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Fake message from " & SocketItem(vsock).RegNick & " rejected ('" & OnBot & "' is not connected via this bot!)"
    FakeCheck = True
    Exit Function
  End If
  If Nick <> "" Then
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If LCase(SocketItem(u).RegNick) = LCase(Nick) Then
          If LCase(SocketItem(u).OnBot) = LCase(OnBot) Then FoundOne = True: Exit For
        End If
      End If
    Next u
    If Not FoundOne Then
      RTU vsock, "error Fake message rejected (no such user: '" & Nick & "@" & OnBot & "')"
      SpreadFlagMessage vsock, "+t", "14[" & Time & "] *** ERROR: Fake message from " & SocketItem(vsock).RegNick & " rejected (no such user: '" & Nick & "@" & OnBot & "')"
      FakeCheck = True
      Exit Function
    End If
  End If
  FakeCheck = False
End Function

'Counts the bots between me and <Bot>
Public Function CountHopsToBot(Bot As String) As Long ' : AddStack "BotNet_CountHopsToBot(" & Bot & ")"
Dim u As Long, CurCheckBot As String, FoundOne As Boolean, cnt As Long
  CurCheckBot = Bot
  Do
    FoundOne = False
    For u = 1 To BotCount
      If LCase(Bots(u).Nick) = LCase(CurCheckBot) Then
        cnt = cnt + 1
        If LCase(Bots(u).SubBotOf) = LCase(BotNetNick) Then CountHopsToBot = cnt: Exit Function
        CurCheckBot = Bots(u).SubBotOf
        FoundOne = True
        Exit For
      End If
    Next u
    If Not FoundOne Then Exit Do
  Loop
  CountHopsToBot = 0
End Function

Public Function IsBotConnected(Name As String) As Boolean ' : AddStack "BotNet_IsBotConnected(" & Name & ")"
Dim u As Long
  For u = 1 To BotCount
    If LCase(Bots(u).Nick) = LCase(Name) Then IsBotConnected = True: Exit Function
  Next u
  IsBotConnected = False
End Function


Public Sub SendToBot(ToBot As String, BMsg As String) ' : AddStack "BotNet_SendToBot(" & ToBot & ", " & BMsg & ")"
Dim SNum As Long, u As Long
  SNum = GetNextBot(ToBot)
  If SNum <> 0 Then
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If u = Bots(SNum).vSocket Then RTU u, BMsg: Exit Sub
      End If
    Next u
  End If
End Sub


Sub LinkCheck(vsock As Long) ' : AddStack "BotNet_LinkCheck(" & vsock & ")"
Dim u As Long
  'Link check
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If u <> vsock And LCase(SocketItem(u).RegNick) = LCase(SocketItem(vsock).RegNick) Then
        If GetSockFlag(u, SF_Status) = SF_Status_InitBotLink Then
          SpreadFlagMessage 0, "+t", "14[" & Time & "] *** Stopped trying to link to " & SocketItem(u).RegNick & ": Bot is connecting to me!"
          RemoveSocket u, 0, "", True
          Exit For
        End If
        If GetSockFlag(u, SF_Status) = SF_Status_BotLinking Or GetSockFlag(u, SF_Status) = SF_Status_BotGetName Or GetSockFlag(u, SF_Status) = SF_Status_BotGetPass Then
          If Val(SocketItem(u).OrderSign) > Val(SocketItem(vsock).OrderSign) Then
            SpreadFlagMessage 0, "+t", "14[" & Time & "] *** Broke outgoing link to " & SocketItem(u).RegNick & ": Duplicate link!"
            RTU u, "error Duplicate Link!"
            RemoveSocket u, 0, "Broke duplicate link to", True
            Exit For
          Else
            SpreadFlagMessage 0, "+t", "14[" & Time & "] *** Broke incoming link from " & SocketItem(u).RegNick & ": Duplicate link!"
            RTU vsock, "error Duplicate Link!"
            RemoveSocket vsock, 0, "Broke duplicate link to", True
            Exit For
          End If
        End If
      End If
    End If
  Next u
End Sub


Public Sub ConnectHubs() ' : AddStack "BotNet_ConnectHubs()"
Dim u As Long, u2 As Long, DontLinkThere As Boolean, LinkedVia As String
Dim FoundHubBot As Boolean
  For u = 1 To BotUserCount
    If MatchFlags(BotUsers(u).BotFlags, "+h") Then
      If GetUserData(u, UD_LinkAddr, "") <> "" Then
        FoundHubBot = True
        DontLinkThere = IsBotConnected(BotUsers(u).Name)
        'Link to hub bots even if they're already in the botnet if uplink is nonhub
        If DontLinkThere Then
          u2 = GetNextBot(BotUsers(u).Name)
          If u2 > 1 Then 'Only reset DontLinkThere if I'm not the uplink
            If MatchFlags(BotUsers(GetUserNum(Bots(u2).Nick)).BotFlags, "-h") Then
              DontLinkThere = False
            End If
          End If
        End If
        'Don't initiate link if bot is already trying to link
        If Not DontLinkThere Then
          For u2 = 1 To SocketCount
            If IsValidSocket(u2) Then
              If LCase(SocketItem(u2).RegNick) = LCase(BotUsers(u).Name) Then
                Select Case GetSockFlag(u2, SF_Status)
                  Case SF_Status_BotLinking, SF_Status_BotGetPass, SF_Status_InitBotLink
                    DontLinkThere = True: Exit For
                End Select
              End If
            End If
          Next u2
        End If
        If Not DontLinkThere Then InitiateBotChat u, True
      End If
    End If
  Next u
  If FoundHubBot = False Then ConnectAlternates
End Sub

Public Sub ConnectAlternates() ' : AddStack "BotNet_ConnectAlternates()"
Dim u As Long, u2 As Long, DontLinkThere As Boolean
  For u = 1 To BotUserCount
    If MatchFlags(BotUsers(u).BotFlags, "+a") Then
      If GetUserData(u, UD_LinkAddr, "") <> "" Then
        DontLinkThere = IsBotConnected(BotUsers(u).Name)
        If Not DontLinkThere Then
          For u2 = 1 To SocketCount
            If IsValidSocket(u2) Then If LCase(SocketItem(u2).RegNick) = LCase(BotUsers(u).Name) Then DontLinkThere = True: Exit For
          Next u2
          If Not DontLinkThere Then InitiateBotChat u, True
        End If
      End If
    End If
  Next u
End Sub

Public Sub UnlinkDummyBots() ' : AddStack "BotNet_UnlinkDummyBots()"
Dim u As Long, t As Long, vsock As Long, RemovedOne As Boolean
Dim Nick As String
  Do
    RemovedOne = False
    For t = 1 To BotCount
      vsock = 0
      If t > UBound(Bots()) Then BotCount = t - 1: Exit For
      If LCase(Bots(t).SubBotOf) = LCase(BotNetNick) Then
        For u = 1 To SocketCount
          If IsValidSocket(u) Then If u = Bots(t).vSocket Then vsock = u: Exit For
        Next u
        If vsock = 0 Then
          SpreadFlagMessage u, "+t", "14[" & Time & "] *** ERROR: I'm not connected to " & Bots(t).Nick & ", but thought so!"
          RemBot Bots(t).Nick, 0, "I'm confused about"
          RemovedOne = True
          Exit For
        End If
      End If
    Next t
    If Not RemovedOne Then Exit Do
  Loop
End Sub

Public Sub PingBots() ' : AddStack "BotNet_PingBots()"
Dim u As Long, t As Long, vsock As Long
  For t = 1 To BotCount
    vsock = 0
    If t > UBound(Bots()) Then BotCount = t - 1: Exit For
    If LCase(Bots(t).SubBotOf) = LCase(BotNetNick) Then
      If Bots(t).SentPing Then
        If Bots(t).GotPingReply Then
          Bots(t).SentPing = False
        Else
          For u = 1 To SocketCount
            If IsValidSocket(u) Then If u = Bots(t).vSocket Then vsock = u: Exit For
          Next u
          If vsock > 0 And SocketItem(vsock).LastEvent + CDate("00:02:00") < Now Then
            u = Bots(t).vSocket
            SS u, "error " & MakeMsg(MSG_BNPingTimeout)
            RemoveSocket vsock, 0, MakeMsg(MSG_BNPingTimeout) & ":", False
          End If
        End If
      Else
        Bots(t).SentPing = True: Bots(t).GotPingReply = False
        SS Bots(t).vSocket, "pi"
      End If
    End If
  Next t
End Sub


Public Function GetLevelSign(Flags As String) As String
  Dim TheLevel As String
  TheLevel = "-"
  If MatchFlags(Flags, "+o") Then TheLevel = "@"
  If MatchFlags(Flags, "+t") Then TheLevel = "%"
  If MatchFlags(Flags, "+m") Then TheLevel = "+"
  If MatchFlags(Flags, "+n") Then TheLevel = "*"
  GetLevelSign = TheLevel
End Function

Public Function GetBotEbene(BotNick As String) As Long
  Dim u As Long
  For u = 1 To BotCount
    If LCase(Bots(u).Nick) = LCase(BotNick) Then
      GetBotEbene = Bots(u).Ebene
      Exit Function
    End If
  Next u
  GetBotEbene = 0
End Function

Public Function GetBotPos(BotNick As String) As Long
  Dim u As Long
  For u = 1 To BotCount
    If LCase(Bots(u).Nick) = LCase(BotNick) Then
      GetBotPos = u
      Exit Function
    End If
  Next u
  GetBotPos = 0
End Function

Public Sub DrawEbene(Ebene As Long, StartPos As Long, TreeLines As String, TUsr As Long, ShowVersions As Boolean) ' : AddStack "Routines_DrawEbene(" & Ebene & ", " & StartPos & ", " & TreeLines & ", " & TUsr & ", " & ShowVersions & ")"
Dim u As Long, u2 As Long, SubBots As Long, a As Long, Part As String, NTL As String
Dim UsNum As Long, BotFlags As String
For u = StartPos To BotCount
  If Bots(u).Ebene = Ebene Then
    SubBots = 0
    For u2 = u + 1 To BotCount
      If Bots(u2).Ebene = Ebene Then SubBots = SubBots + 1
      If Bots(u2).Ebene < Ebene Then Exit For
    Next u2
    Part = u & " " & Bots(u).vSocket & " "
    If Ebene > 0 Then
      For a = 1 To Ebene - 1
        If Mid(TreeLines, a, 1) = "x" Then Part = Part & "  |  " Else Part = Part & "     "
      Next a
      If SubBots = 0 Then
        Part = Part & "  `-" & Bots(u).SharingFlag
      Else
        Part = Part & "  |-" & Bots(u).SharingFlag
      End If
    End If
    'Get bot flags
    BotFlags = ""
    UsNum = GetUserNum(Bots(u).Nick)
    If UsNum > 0 Then
      If BotUsers(UsNum).BotFlags <> "" Then
        BotFlags = " 14(+" & BotUsers(UsNum).BotFlags & ")"
      End If
    End If
    'Show one line of the tree
    If ShowVersions Then
      If Base64ToLong(Bots(u).Version) > 0 Then
        Part = "14 " & Part & "" & Bots(u).Nick & " (" & MakeVersionString(Base64ToLong(Bots(u).Version)) & ")"
      Else
        Part = "14 " & Part & "" & Bots(u).Nick
      End If
    Else
      Part = "14 " & Part & "" & Bots(u).Nick + BotFlags
    End If
    TU TUsr, Part
    NTL = TreeLines
    If Ebene > 0 Then
      If SubBots > 0 Then
        If Ebene <= Len(NTL) Then Mid(NTL, Ebene, 1) = "x" Else NTL = NTL & "x"
      Else
        If Ebene <= Len(NTL) Then Mid(NTL, Ebene, 1) = " " Else NTL = NTL & " "
      End If
    End If
    DrawEbene Ebene + 1, u + 1, NTL, TUsr, ShowVersions
  ElseIf Bots(u).Ebene < Ebene Then
    Exit Sub
  End If
Next u
End Sub

