Attribute VB_Name = "Server_Channels"
',-======================- ==-- -  -
'|   AnGeL - Server - Channels
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Type Chatter
  Nick As String
  RegNick As String
  Status As String
  Hostmask As String
  IPmask As String
  LastLine As String
  CTCPs As Long
  UserNum As Long
  NickChanges As Long
  LineCount As Long
  CharCount As Long
  RepeatCount As Long
  LastEvent As Currency
End Type

Private Type DesiredBan
  Mask As String
  Expires As Currency
End Type

Private Type DesiredExcept
  Mask As String
  Expires As Currency
End Type

Private Type DesiredInvite
  Mask As String
  Expires As Currency
End Type

Private Type ChannelBan
  Mask As String
  CreatedAt As Date
End Type

Private Type ChannelExcept
  Mask As String
  CreatedAt As Date
End Type

Private Type ChannelInvite
  Mask As String
  CreatedAt As Date
End Type

Private Type KickListEntry
  Nick As String
  Hostmask As String
  Message As String
End Type


Public Type Except
  Hostmask As String
  Channel As String
  CreatedAt As Date
  CreatedBy As String
End Type

Public Type Invite
  Hostmask As String
  Channel As String
  CreatedAt As Date
  CreatedBy As String
End Type


Private Type Channel
  Name As String
  UserCount As Long
  User() As Chatter
  Mode As String
  Topic As String
  
  DesiredBanCount As Long
  DesiredBanList() As DesiredBan
  DesiredExceptCount As Long
  DesiredExceptList() As DesiredBan
  DesiredInviteCount As Long
  DesiredInviteList() As DesiredBan
  
  KickCount As Long
  KickList() As KickListEntry
  ToBanCount As Long
  ToBanList() As String
  ToExceptCount As Long
  ToExceptList() As String
  ToInviteCount As Long
  ToInviteList() As String
  
  GotOPs As Boolean
  GotHOPs As Boolean
  CompletedWHO As Boolean
  CompletedMode As Boolean
  CompletedBANS As Boolean
  CompletedExcepts As Boolean
  CompletedInvites As Boolean
  NewbieGreeting As String
  CloneKick As Boolean
  AutoVoice As Byte
  ProtectFriends As Boolean
  DeopUnknownUsers As Byte
  ReactToSeen As Byte
  ReactToWhois As Boolean
  ReactToWhatis As Boolean
  ColorKick As Boolean
  AllowVoiceControl As Boolean
  Secret As Boolean
  EnforceBans As Boolean
  BanMask As Long
  BanCount As Long
  ExceptCount As Long
  InviteCount As Long
  BanList(50) As ChannelBan
  InviteList(50) As ChannelInvite
  ExceptList(50) As ChannelExcept
  DidUnban As Boolean
  DidUnExcept As Boolean
  DidUnInvite As Boolean
  FloodEvents As Long
  InFlood As Boolean
  MaxLines As Long
  MaxChars As Long
  MaxRepeats As Long
End Type


Private Type PermChannel
  Name As String
  Status As String
End Type


Public PermChanCount As Byte
Public PermChannels() As PermChannel
Public ChanCount As Byte
Public Channels() As Channel


Sub Channels_Load()
  ReDim Preserve Channels(5)
  ReDim Preserve PermChannels(5)
End Sub


Sub Channels_Unload()
'
End Sub


'Add a channel to the Channels() array and read ChanSetup settings
Sub AddChannel(Chan As String) ' : AddStack "ServerRoutines_AddChannel(" & Chan$ & ")"
Dim u As Long, NewChanNum As Integer, TempStr As String
  'Check if channel is already in array
  For u = 1 To ChanCount
    If LCase(Channels(u).Name) = LCase(Chan) Then
      NewChanNum = u: Exit For
    End If
  Next u
  If NewChanNum = 0 Then
    ChanCount = ChanCount + 1: If ChanCount > UBound(Channels) Then ReDim Preserve Channels(UBound(Channels) + 5)
    NewChanNum = ChanCount
    GUI_frmWinsock.lstChannels.AddItem Chan
  End If
  
  'Reset channel information
  Channels(NewChanNum).Name = Chan
  Channels(NewChanNum).UserCount = 0
  Channels(NewChanNum).KickCount = 0
  Channels(NewChanNum).ToBanCount = 0
  Channels(NewChanNum).ToInviteCount = 0
  Channels(NewChanNum).ToExceptCount = 0
  Channels(NewChanNum).DesiredBanCount = 0
  Channels(NewChanNum).DesiredExceptCount = 0
  Channels(NewChanNum).DesiredInviteCount = 0
  Channels(NewChanNum).GotOPs = False
  Channels(NewChanNum).GotHOPs = False
  Channels(NewChanNum).CompletedWHO = False
  Channels(NewChanNum).CompletedMode = False
  Channels(NewChanNum).CompletedBANS = False
  Channels(NewChanNum).BanCount = 0
  Channels(NewChanNum).ExceptCount = 0
  Channels(NewChanNum).InviteCount = 0
  Channels(NewChanNum).DidUnban = False
  Channels(NewChanNum).InFlood = False
  Channels(NewChanNum).FloodEvents = 0
  Channels(NewChanNum).Topic = ""
  ReDim Preserve Channels(NewChanNum).User(5)
  ReDim Preserve Channels(NewChanNum).KickList(5)
  ReDim Preserve Channels(NewChanNum).ToBanList(5)
  ReDim Preserve Channels(NewChanNum).ToInviteList(5)
  ReDim Preserve Channels(NewChanNum).ToExceptList(5)
  ReDim Preserve Channels(NewChanNum).DesiredBanList(5)
  ReDim Preserve Channels(NewChanNum).DesiredExceptList(5)
  ReDim Preserve Channels(NewChanNum).DesiredInviteList(5)
  
  SetPermChanStat Chan$, ChanStat_OK
  SpreadFlagMessage 0, "+m", "10*** I joined the channel " & Chan$ & "."
  If Not Initializing Then
    SendLine "who " & Chan$, 1
    SendLine "mode " & Chan$, 1
    SendLine "mode " & Chan$ & " +b", 1
    If InStr(ServerChannelModes, "e") <> 0 Then SendLine "mode " & Chan$ & " +e", 1
    If InStr(ServerChannelModes, "I") <> 0 Then SendLine "mode " & Chan$ & " +I", 1
  End If
  
  'Read settings made in ChanSetup
  Channels(NewChanNum).NewbieGreeting = GetChannelSetting(Chan$, "NewbieGreeting", "")
  Channels(NewChanNum).ProtectFriends = (LCase(GetChannelSetting(Chan$, "ProtectFriends", "off")) = "on")
  Channels(NewChanNum).CloneKick = (LCase(GetChannelSetting(Chan$, "CloneKick", "off")) = "on")
  Select Case LCase(GetChannelSetting(Chan$, "AutoVoice", "off"))
    Case "off": Channels(NewChanNum).AutoVoice = 0
    Case "on": Channels(NewChanNum).AutoVoice = 1
    Case "ext": Channels(NewChanNum).AutoVoice = 2
  End Select
  Select Case LCase(GetChannelSetting(Chan$, "DeopUnknownUsers", "off"))
    Case "off": Channels(NewChanNum).DeopUnknownUsers = 0
    Case "on": Channels(NewChanNum).DeopUnknownUsers = 1
    Case "ext": Channels(NewChanNum).DeopUnknownUsers = 2
  End Select
  TempStr = GetChannelSetting(Chan$, "ReactToSeen", "1")
  If Not IsNumeric(TempStr) Then
    Select Case LCase(TempStr)
      Case "off"
        WritePPString Chan$, "ReactToSeen", "0", HomeDir & "Channels.ini"
        Channels(NewChanNum).ReactToSeen = 0
      Case Else
        WritePPString Chan$, "ReactToSeen", "1", HomeDir & "Channels.ini"
        Channels(NewChanNum).ReactToSeen = 1
    End Select
  Else
    Channels(NewChanNum).ReactToSeen = CByte(GetChannelSetting(Chan$, "ReactToSeen", "1"))
  End If
  Channels(NewChanNum).ReactToWhois = (LCase(GetChannelSetting(Chan$, "ReactToWhois", "on")) = "on")
  Channels(NewChanNum).AllowVoiceControl = (LCase(GetChannelSetting(Chan$, "AllowVoiceControl", "on")) = "on")
  Channels(NewChanNum).ColorKick = (LCase(GetChannelSetting(Chan$, "ColorKick", "off")) = "on")
  Channels(NewChanNum).Secret = (LCase(GetChannelSetting(Chan$, "Secret", "off")) = "on")
  Channels(NewChanNum).ReactToWhatis = (LCase(GetChannelSetting(Chan$, "ReactToWhatis", "on")) = "on")
  Channels(NewChanNum).EnforceBans = (LCase(GetChannelSetting(Chan$, "EnforceBans", "on")) = "on")
  TempStr = GetChannelSetting(Chan$, "FloodSettings", "10 900 3")
  If LCase(TempStr) = "off" Then
    Channels(NewChanNum).MaxLines = 0
    Channels(NewChanNum).MaxChars = 0
    Channels(NewChanNum).MaxRepeats = 0
  Else
    If IsNumeric(Param(TempStr, 1)) Then Channels(NewChanNum).MaxLines = CLng(Param(TempStr, 1))
    If IsNumeric(Param(TempStr, 2)) Then Channels(NewChanNum).MaxChars = CLng(Param(TempStr, 2))
    If IsNumeric(Param(TempStr, 3)) Then Channels(NewChanNum).MaxRepeats = CLng(Param(TempStr, 3))
  End If
  Channels(NewChanNum).BanMask = GetChannelSetting(Chan$, "BanMask", "3")
  
  RemOrder "gop " & Chan$
  RemOrder "xgop " & Chan$
End Sub

'Add a user to a channel in the Channels() array
Function AddChanUser(ChNum As Long, Nick As String, RegUser As String, RegUserNum As Long, Hostmask As String) ' : AddStack "ServerRoutines_AddChanUser(" & ChNum & ", " & Nick$ & ", " & RegUser & ", " & RegUserNum & ", " & HostMask & ")"
  Dim Host As String, Resolved As String
  Channels(ChNum).UserCount = Channels(ChNum).UserCount + 1
  If Channels(ChNum).UserCount > UBound(Channels(ChNum).User()) Then ReDim Preserve Channels(ChNum).User(UBound(Channels(ChNum).User()) + 5)
  Channels(ChNum).User(Channels(ChNum).UserCount).Nick = Nick
  Channels(ChNum).User(Channels(ChNum).UserCount).Hostmask = Hostmask
  Host = Mask(Hostmask, 11)
  If Not IsValidIP(Host) Then
    Resolved = GetCacheIP(Host, False)
    If Resolved = "" Then
      Channels(ChNum).User(Channels(ChNum).UserCount).IPmask = ""
    Else
      Channels(ChNum).User(Channels(ChNum).UserCount).IPmask = Mask(Hostmask, 14) & Resolved
    End If
  Else
    Channels(ChNum).User(Channels(ChNum).UserCount).IPmask = ""
  End If
  Channels(ChNum).User(Channels(ChNum).UserCount).Status = ""
  Channels(ChNum).User(Channels(ChNum).UserCount).RegNick = RegUser
  Channels(ChNum).User(Channels(ChNum).UserCount).UserNum = RegUserNum
  Channels(ChNum).User(Channels(ChNum).UserCount).LastLine = ""
  Channels(ChNum).User(Channels(ChNum).UserCount).CTCPs = 0
  Channels(ChNum).User(Channels(ChNum).UserCount).CharCount = 0
  Channels(ChNum).User(Channels(ChNum).UserCount).LineCount = 0
  Channels(ChNum).User(Channels(ChNum).UserCount).RepeatCount = 0
  Channels(ChNum).User(Channels(ChNum).UserCount).LastEvent = WinTickCount
  AddChanUser = Channels(ChNum).UserCount
End Function

Public Sub CheckChannels() ' : AddStack "ServerRoutines_CheckChannels()"
  Dim Cha As Long, u As Long, Chan As String, JoinLine As String, KeyLine As String, Rest As String
  JoinLine = "": KeyLine = ""
  For Cha = 1 To PermChanCount
    Chan = PermChannels(Cha).Name
    If FindChan(Chan) = 0 And PermChannels(Cha).Status <> ChanStat_Left Then
      If FindChan(Chan) = 0 And PermChannels(Cha).Status <> ChanStat_Unsup Then
        If FindChan(Chan) = 0 And ChanCount < ServerMaxChannels Then
          If Chan <> "" Then
            Rest = GetChannelSetting(Chan, "EnforceModes", "x")
            If ParamCount(Rest) > 1 Then Rest = Param(Rest, ParamCount(Rest)) Else Rest = "x"
            If ServerInfo.SupportsMultiChanJoin Then
              If JoinLine = "" Then JoinLine = Chan: KeyLine = KeyLine + Rest Else JoinLine = JoinLine & "," & Chan: KeyLine = KeyLine & "," & Rest
            Else
              SendLine "join " & Chan & " " & Rest, 1
            End If
          End If
        End If
      End If
    End If
  Next Cha
  If JoinLine <> "" Then SendLine "join " & JoinLine & " " & KeyLine, 1
End Sub

Sub RemChanUser(Name As String, Chan As Long) ' : AddStack "ServerRoutines_RemChanUser(" & Name & ", " & Chan & ")"
Dim u As Long, StartP As Long, TempA As String, TempB As String, FoundOne As Boolean
  StartP = FindUser(Name, Chan)
  If StartP = 0 Then Exit Sub
  For u = StartP To Channels(Chan).UserCount - 1
    Channels(Chan).User(u) = Channels(Chan).User(u + 1)
  Next u
  Channels(Chan).UserCount = Channels(Chan).UserCount - 1
  
  'Remove flood settings
  TempA = LCase("RemLine " & Channels(Chan).Name & " " & Name)
  TempB = LCase("RemChars " & Channels(Chan).Name & " " & Name & " ")
  Do
    FoundOne = False
    For u = 1 To EventCount
      If (LCase(Events(u).DoThis) = TempA) Or (InStr(LCase(Events(u).DoThis), TempB) > 0) Then
        Events(u).DoThis = ""
        FoundOne = True
        Exit For
      End If
    Next u
  Loop While FoundOne
End Sub

Sub RemChan(Chan As String) ' : AddStack "ServerRoutines_RemChan(" & Chan & ")"
Dim u As Long, StartP As Long, TempA As String, TempB As String, FoundOne As Boolean
  StartP = FindChan(Chan)
  If StartP = 0 Then Exit Sub
  
  'Remove flood settings
  TempA = LCase("RemLine " & Channels(StartP).Name & " ")
  TempB = LCase("RemChars " & Channels(StartP).Name & " ")
  Do
    FoundOne = False
    For u = 1 To EventCount
      If (InStr(LCase(Events(u).DoThis), TempA) > 0) Or (InStr(LCase(Events(u).DoThis), TempB) > 0) Then
        Events(u).DoThis = ""
        FoundOne = True
        Exit For
      End If
    Next u
  Loop While FoundOne
  
  'Remove channel
  For u = StartP To ChanCount - 1
    Channels(u) = Channels(u + 1)
  Next u
  ChanCount = ChanCount - 1
  If ChanCount < (UBound(Channels) - 5) Then ReDim Preserve Channels(UBound(Channels) - 5)
  For u = 0 To GUI_frmWinsock.lstChannels.ListCount - 1
    If GUI_frmWinsock.lstChannels.List(u) = Chan Then GUI_frmWinsock.lstChannels.RemoveItem u
  Next u
End Sub

Public Function GetChannelSetting(Channel As String, Setting As String, Default As String) As String
  GetChannelSetting = GetPPString(Channel, Setting, GetPPString("Default", Setting, Default, HomeDir & "Channels.ini"), HomeDir & "Channels.ini")
End Function


Function RemovePermChan(Channel As String) As Long
  Dim FileName As String
  FileName = HomeDir & "Autojoin.txt"
  If Not LineInFile(Channel, FileName) Then RemovePermChan = 0: Exit Function
  RemLineFromFile Channel, FileName
  RemovePermChan = 1
  ReadAutoJoinChannels
End Function

Function FindUser(Name As String, Chan As Long) As Long
  Dim u As Long
  FindUser = 0
  If Chan > ChanCount Or Chan <= 0 Or Name = "" Then Exit Function
  For u = 1 To Channels(Chan).UserCount
    If LCase(Channels(Chan).User(u).Nick) = LCase(Name) Then
      FindUser = u
      Exit Function
    End If
  Next u
End Function
         
'Add a ban to the Channels().DesiredBanList() array
Sub AddDesiredBan(ChNum As Long, Hostmask As String) ' : AddStack "Routines_AddDesiredBan(" & ChNum & ", " & Hostmask & ")"
Dim u As Long
  If Not IsBanned(ChNum, Hostmask) Then
    For u = 1 To Channels(ChNum).DesiredBanCount
      If LCase(Channels(ChNum).DesiredBanList(u).Mask) = LCase(Hostmask) Then Exit Sub
    Next u
    Channels(ChNum).DesiredBanCount = Channels(ChNum).DesiredBanCount + 1
    If Channels(ChNum).DesiredBanCount > UBound(Channels(ChNum).DesiredBanList()) Then ReDim Preserve Channels(ChNum).DesiredBanList(UBound(Channels(ChNum).DesiredBanList()) + 5)
    Channels(ChNum).DesiredBanList(Channels(ChNum).DesiredBanCount).Mask = Hostmask
    Channels(ChNum).DesiredBanList(Channels(ChNum).DesiredBanCount).Expires = WinTickCount + 60000
  End If
End Sub

'Remove a ban from the Channels().DesiredBanList() array
Sub RemDesiredBan(ChNum As Long, ByVal Hostmask As String) ' : AddStack "Routines_RemDesiredBan(" & ChNum & ", " & Hostmask & ")"
Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).DesiredBanCount
    If LCase(Channels(ChNum).DesiredBanList(u).Mask) = LCase(Hostmask) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).DesiredBanCount - 1
    Channels(ChNum).DesiredBanList(u) = Channels(ChNum).DesiredBanList(u + 1)
  Next u
  Channels(ChNum).DesiredBanCount = Channels(ChNum).DesiredBanCount - 1
  u = ((Channels(ChNum).DesiredBanCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).DesiredBanList()) Then ReDim Preserve Channels(ChNum).DesiredBanList(u)
End Sub

'Add a ban to the Channels().ToBanList() array
Function AddToBan(ChNum As Long, Hostmask As String) As Boolean ' : AddStack "Routines_AddToBan(" & ChNum & ", " & Hostmask & ")"
Dim u As Long, IPos As Long, u2 As Long, DidReplace As Boolean, TempStr As String
  If Not IsBanned(ChNum, Hostmask) Then
    If Not IsOrdered("ban " & Channels(ChNum).Name & " " & Hostmask) Then
      If Channels(ChNum).ToBanCount + Channels(ChNum).BanCount > BanLimit Then AddToBan = False: Exit Function
      For u = 1 To Channels(ChNum).ToBanCount
        If MatchWM(LCase(Channels(ChNum).ToBanList(u)), LCase(Hostmask)) Then AddToBan = False: Exit Function
      Next u
      Order "ban " & Channels(ChNum).Name & " " & Hostmask, 40
      Channels(ChNum).ToBanCount = Channels(ChNum).ToBanCount + 1
      If Channels(ChNum).ToBanCount > UBound(Channels(ChNum).ToBanList()) Then ReDim Preserve Channels(ChNum).ToBanList(UBound(Channels(ChNum).ToBanList()) + 5)
      IPos = Int(Rnd * Channels(ChNum).ToBanCount) + 1
      For u = Channels(ChNum).ToBanCount To IPos + 1 Step -1
        Channels(ChNum).ToBanList(u) = Channels(ChNum).ToBanList(u - 1)
      Next u
      Channels(ChNum).ToBanList(IPos) = Hostmask
      AddToBan = True
    Else
      AddToBan = False
    End If
  Else
    AddToBan = False
  End If
End Function

'Add a user to the Channels().KickList() array
Sub AddKickUser(ChNum As Long, Nick As String, Hostmask As String, Message As String) ' : AddStack "Routines_AddKickUser(" & ChNum & ", " & Nick$ & ", " & Hostmask & ", " & Message & ")"
Dim u As Long, IPos As Long
  For u = 1 To Channels(ChNum).KickCount
    If LCase(Channels(ChNum).KickList(u).Nick) = LCase(Nick) Then Exit Sub
  Next u
  Channels(ChNum).KickCount = Channels(ChNum).KickCount + 1
  If Channels(ChNum).KickCount > UBound(Channels(ChNum).KickList()) Then ReDim Preserve Channels(ChNum).KickList(UBound(Channels(ChNum).KickList()) + 5)
  IPos = Int(Rnd * Channels(ChNum).KickCount) + 1
  For u = Channels(ChNum).KickCount To IPos + 1 Step -1
    Channels(ChNum).KickList(u) = Channels(ChNum).KickList(u - 1)
  Next u
  Channels(ChNum).KickList(IPos).Nick = Nick
  Channels(ChNum).KickList(IPos).Hostmask = Hostmask
  Channels(ChNum).KickList(IPos).Message = Message
End Sub


'Add a Invite to the Channels().ToInviteList() array
Function AddToInvite(ChNum As Long, Hostmask As String) As Boolean ' : AddStack "Routines_AddToInvite(" & ChNum & ", " & Hostmask & ")"
Dim u As Long, IPos As Long, u2 As Long, DidReplace As Boolean, TempStr As String
  If Not IsInvited(ChNum, Hostmask) Then
    If Not IsOrdered("Invite " & Channels(ChNum).Name & " " & Hostmask) Then
      If Channels(ChNum).ToInviteCount + Channels(ChNum).InviteCount > BanLimit Then AddToInvite = False: Exit Function
      For u = 1 To Channels(ChNum).ToInviteCount
        If MatchWM(LCase(Channels(ChNum).ToInviteList(u)), LCase(Hostmask)) Then AddToInvite = False: Exit Function
      Next u
      Order "Invite " & Channels(ChNum).Name & " " & Hostmask, 40
      Channels(ChNum).ToInviteCount = Channels(ChNum).ToInviteCount + 1
      If Channels(ChNum).ToInviteCount > UBound(Channels(ChNum).ToInviteList()) Then ReDim Preserve Channels(ChNum).ToInviteList(UBound(Channels(ChNum).ToInviteList()) + 5)
      IPos = Int(Rnd * Channels(ChNum).ToInviteCount) + 1
      For u = Channels(ChNum).ToInviteCount To IPos + 1 Step -1
        Channels(ChNum).ToInviteList(u) = Channels(ChNum).ToInviteList(u - 1)
      Next u
      Channels(ChNum).ToInviteList(IPos) = Hostmask
      AddToInvite = True
    Else
      AddToInvite = False
    End If
  Else
    AddToInvite = False
  End If
End Function
'Add a Except to the Channels().ToExceptList() array
Function AddToExcept(ChNum As Long, Hostmask As String) As Boolean ' : AddStack "Routines_AddToExcept(" & ChNum & ", " & Hostmask & ")"
Dim u As Long, IPos As Long, u2 As Long, DidReplace As Boolean, TempStr As String
  If Not IsExceptd(ChNum, Hostmask) Then
    If Not IsOrdered("Except " & Channels(ChNum).Name & " " & Hostmask) Then
      If Channels(ChNum).ToExceptCount + Channels(ChNum).ExceptCount > BanLimit Then AddToExcept = False: Exit Function
      For u = 1 To Channels(ChNum).ToExceptCount
        If MatchWM(LCase(Channels(ChNum).ToExceptList(u)), LCase(Hostmask)) Then AddToExcept = False: Exit Function
      Next u
      Order "Except " & Channels(ChNum).Name & " " & Hostmask, 40
      Channels(ChNum).ToExceptCount = Channels(ChNum).ToExceptCount + 1
      If Channels(ChNum).ToExceptCount > UBound(Channels(ChNum).ToExceptList()) Then ReDim Preserve Channels(ChNum).ToExceptList(UBound(Channels(ChNum).ToExceptList()) + 5)
      IPos = Int(Rnd * Channels(ChNum).ToExceptCount) + 1
      For u = Channels(ChNum).ToExceptCount To IPos + 1 Step -1
        Channels(ChNum).ToExceptList(u) = Channels(ChNum).ToExceptList(u - 1)
      Next u
      Channels(ChNum).ToExceptList(IPos) = Hostmask
      AddToExcept = True
    Else
      AddToExcept = False
    End If
  Else
    AddToExcept = False
  End If
End Function

'Add a invite to the Channels().DesiredInviteList() array
Sub AddDesiredInvite(ChNum As Long, Hostmask As String) ' : AddStack "Routines_AddDesiredInvite(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long
  If Not IsInvited(ChNum, Hostmask) Then
    For u = 1 To Channels(ChNum).DesiredInviteCount
      If LCase(Channels(ChNum).DesiredInviteList(u).Mask) = LCase(Hostmask) Then Exit Sub
    Next u
    Channels(ChNum).DesiredInviteCount = Channels(ChNum).DesiredInviteCount + 1
    If Channels(ChNum).DesiredInviteCount > UBound(Channels(ChNum).DesiredInviteList()) Then ReDim Preserve Channels(ChNum).DesiredInviteList(UBound(Channels(ChNum).DesiredInviteList()) + 5)
    Channels(ChNum).DesiredInviteList(Channels(ChNum).DesiredInviteCount).Mask = Hostmask
    Channels(ChNum).DesiredInviteList(Channels(ChNum).DesiredInviteCount).Expires = WinTickCount + 60000
  End If
End Sub

'Add a except to the Channels().DesiredExceptList() array
Sub AddDesiredExcept(ChNum As Long, Hostmask As String) ' : AddStack "Routines_AddDesiredExcept(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long
  If Not IsExceptd(ChNum, Hostmask) Then
    For u = 1 To Channels(ChNum).DesiredExceptCount
      If LCase(Channels(ChNum).DesiredExceptList(u).Mask) = LCase(Hostmask) Then Exit Sub
    Next u
    Channels(ChNum).DesiredExceptCount = Channels(ChNum).DesiredExceptCount + 1
    If Channels(ChNum).DesiredExceptCount > UBound(Channels(ChNum).DesiredExceptList()) Then ReDim Preserve Channels(ChNum).DesiredExceptList(UBound(Channels(ChNum).DesiredExceptList()) + 5)
    Channels(ChNum).DesiredExceptList(Channels(ChNum).DesiredExceptCount).Mask = Hostmask
    Channels(ChNum).DesiredExceptList(Channels(ChNum).DesiredExceptCount).Expires = WinTickCount + 60000
  End If
End Sub


'Remove a except from the Channels().DesiredExceptList() array
Sub RemDesiredExcept(ChNum As Long, ByVal Hostmask As String) ' : AddStack "Routines_RemDesiredExcept(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).DesiredExceptCount
    If LCase(Channels(ChNum).DesiredExceptList(u).Mask) = LCase(Hostmask) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).DesiredExceptCount - 1
    Channels(ChNum).DesiredExceptList(u) = Channels(ChNum).DesiredExceptList(u + 1)
  Next u
  Channels(ChNum).DesiredExceptCount = Channels(ChNum).DesiredExceptCount - 1
  u = ((Channels(ChNum).DesiredExceptCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).DesiredExceptList()) Then ReDim Preserve Channels(ChNum).DesiredExceptList(u)
End Sub

'Remove a invite from the Channels().DesiredInviteList() array
Sub RemDesiredInvite(ChNum As Long, ByVal Hostmask As String) ' : AddStack "Routines_RemDesiredInvite(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).DesiredInviteCount
    If LCase(Channels(ChNum).DesiredInviteList(u).Mask) = LCase(Hostmask) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).DesiredInviteCount - 1
    Channels(ChNum).DesiredInviteList(u) = Channels(ChNum).DesiredInviteList(u + 1)
  Next u
  Channels(ChNum).DesiredInviteCount = Channels(ChNum).DesiredInviteCount - 1
  u = ((Channels(ChNum).DesiredInviteCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).DesiredInviteList()) Then ReDim Preserve Channels(ChNum).DesiredInviteList(u)
End Sub

'Remove a user from the Channels().KickList() array
Sub RemKickUser(ChNum As Long, Nick As String) ' : AddStack "Routines_RemKickUser(" & ChNum & ", " & Nick$ & ")"
Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).KickCount
    If LCase(Channels(ChNum).KickList(u).Nick) = LCase(Nick) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).KickCount - 1
    Channels(ChNum).KickList(u) = Channels(ChNum).KickList(u + 1)
  Next u
  Channels(ChNum).KickCount = Channels(ChNum).KickCount - 1
  u = ((Channels(ChNum).KickCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).KickList()) Then ReDim Preserve Channels(ChNum).KickList(u)
End Sub

'Remove a ban from the Channels().ToBanList() array
Sub RemToBan(ChNum As Long, Hostmask As String) ' : AddStack "Routines_RemToBan(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).ToBanCount
    If LCase(Channels(ChNum).ToBanList(u)) = LCase(Hostmask) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).ToBanCount - 1
    Channels(ChNum).ToBanList(u) = Channels(ChNum).ToBanList(u + 1)
  Next u
  Channels(ChNum).ToBanCount = Channels(ChNum).ToBanCount - 1
  u = ((Channels(ChNum).ToBanCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).ToBanList()) Then ReDim Preserve Channels(ChNum).ToBanList(u)
End Sub

'Remove a Except from the Channels().ToExceptList() array
Sub RemToExcept(ChNum As Long, Hostmask As String) ' : AddStack "Routines_RemToExcept(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).ToExceptCount
    If LCase(Channels(ChNum).ToExceptList(u)) = LCase(Hostmask) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).ToExceptCount - 1
    Channels(ChNum).ToExceptList(u) = Channels(ChNum).ToExceptList(u + 1)
  Next u
  Channels(ChNum).ToExceptCount = Channels(ChNum).ToExceptCount - 1
  u = ((Channels(ChNum).ToExceptCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).ToExceptList()) Then ReDim Preserve Channels(ChNum).ToExceptList(u)
End Sub

'Remove a Invite from the Channels().ToInviteList() array
Sub RemToInvite(ChNum As Long, Hostmask As String) ' : AddStack "Routines_RemToInvite(" & ChNum & ", " & Hostmask & ")"
Dim u As Long, UsNum As Long
  For u = 1 To Channels(ChNum).ToInviteCount
    If LCase(Channels(ChNum).ToInviteList(u)) = LCase(Hostmask) Then UsNum = u: Exit For
  Next u
  If UsNum = 0 Then Exit Sub
  For u = UsNum To Channels(ChNum).ToInviteCount - 1
    Channels(ChNum).ToInviteList(u) = Channels(ChNum).ToInviteList(u + 1)
  Next u
  Channels(ChNum).ToInviteCount = Channels(ChNum).ToInviteCount - 1
  u = ((Channels(ChNum).ToInviteCount \ 5) + 1) * 5
  If u < UBound(Channels(ChNum).ToInviteList()) Then ReDim Preserve Channels(ChNum).ToInviteList(u)
End Sub

Function IsPermChan(Chan As String) As Boolean ' : AddStack "Routines_IsPermChan(" & Chan$ & ")"
Dim u As Long
  For u = 1 To PermChanCount
    If LCase(Chan) = LCase(PermChannels(u).Name) Then IsPermChan = True: Exit Function
  Next u
  IsPermChan = False
End Function

Sub SetPermChanStat(Chan As String, ToWhat As String) ' : AddStack "Routines_SetPermChanStat(" & Chan$ & ", " & ToWhat & ")"
Dim u As Long
  For u = 1 To PermChanCount
    If LCase(Chan) = LCase(PermChannels(u).Name) Then PermChannels(u).Status = ToWhat: Exit Sub
  Next u
End Sub

Public Function ChannelMatch(Wild As String, Full As String) As Boolean
  Dim Wild2 As String, Full2 As String
  Select Case Mid(Full, 1, 1)
    Case "#"
      Wild2 = Replace(LCase(MakeININick(Wild)), "#", "µ")
      Full2 = Replace(LCase(MakeININick(Full)), "#", "µ")
    Case "!"
      Wild2 = Replace(LCase(MakeININick(Wild)), "!", "µ")
      Full2 = Replace(LCase(MakeININick(Full)), "!", "µ")
    Case "+"
      Wild2 = Replace(LCase(MakeININick(Wild)), "+", "µ")
      Full2 = Replace(LCase(MakeININick(Full)), "+", "µ")
    Case "&"
      Wild2 = Replace(LCase(MakeININick(Wild)), "&", "µ")
      Full2 = Replace(LCase(MakeININick(Full)), "&", "µ")
  End Select
  ChannelMatch = Full2 Like Wild2
End Function

Public Function InAutoJoinChannels(Channel As String) As Boolean ' : AddStack "Server_InAutoJoinChannels(" & Channel & ")"
  Dim u As Long, JoinLine As String
  If Channel = "" Then InAutoJoinChannels = False: Exit Function
  For u = 1 To PermChanCount
    If LCase(Channel) = LCase(PermChannels(u).Name) Then InAutoJoinChannels = True: Exit Function
  Next u
  InAutoJoinChannels = False
End Function

Public Sub ReadAutoJoinChannels() ' : AddStack "Server_ReadAutoJoinChannels()"
Dim u As Long, Chan As String, FileNum As Integer
  On Local Error Resume Next
  PermChanCount = 0
  FileNum = FreeFile: Open HomeDir & "Autojoin.txt" For Input As #FileNum
  If Err.Number > 0 Then Close #FileNum: Exit Sub
    Do While Not EOF(FileNum)
      Line Input #FileNum, Chan
      If Trim(Chan) <> "" Then
        PermChanCount = PermChanCount + 1
        If PermChanCount > UBound(PermChannels) Then ReDim Preserve PermChannels(UBound(PermChannels) + 5)
        PermChannels(PermChanCount).Name = Trim(Chan)
        PermChannels(PermChanCount).Status = IIf(FindChan(Chan) <> 0, ChanStat_OK, ChanStat_NotOn)
      End If
    Loop
  Close #FileNum
End Sub

Public Function GetChannelKey(ChNum As Long) As String ' : AddStack "Routines_GetChannelKey(" & ChNum & ")"
  If InStr(Param(Channels(ChNum).Mode, 1), "k") > 0 Then
    GetChannelKey = Param(Channels(ChNum).Mode, ParamCount(Channels(ChNum).Mode))
  Else
    GetChannelKey = ""
  End If
End Function

Public Function GetChannelLimit(ChNum As Long) As Long ' : AddStack "Routines_GetChannelLimit(" & ChNum & ")"
  If InStr(Param(Channels(ChNum).Mode, 1), "l") > 0 Then
    GetChannelLimit = CLng(Param(Channels(ChNum).Mode, 2))
  Else
    GetChannelLimit = 0
  End If
End Function

Function IsValidChannel(ChannelName As String) As Boolean ' : AddStack "SCMsExtensions_IsValidChannel(" & ChannelName & ")"
  Dim i As Integer
  If Len(ChannelName) < 1 Then IsValidChannel = False: Exit Function
  For i = 1 To Len(ServerChannelPrefixes)
    If Mid(ServerChannelPrefixes, i, 1) = Mid(ChannelName, 1, 1) Then
      IsValidChannel = True: Exit Function
    End If
  Next i
End Function

