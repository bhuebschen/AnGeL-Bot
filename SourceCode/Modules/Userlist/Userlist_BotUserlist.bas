Attribute VB_Name = "Userlist_BotUserlist"
',-======================- ==-- -  -
'|   AnGeL - Userlist - BotUserlist
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Type ChannelFlag
  Channel As String
  Flags As String
End Type


Private Type BotUser
  Name As String
  Password As String
  Flags As String
  BotFlags As String
  ChannelFlagCount As Long
  ChannelFlags(20) As ChannelFlag
  HostMaskCount As Long
  HostMasks(20) As String
  UserData As String
  PLChannel As Long
  ValidSession As Byte
  OldFullAddress As String
End Type


Public BotUserCount As Long
Public BotUserNum As Long
Public OldCheckSum As Long
Public BaseFlags As String
Public BotUsers() As BotUser


Public Const UD_LinkAddr As String = "addr"


Sub BotUserlist_load()
  ReDim Preserve BotUsers(5)
End Sub

Sub BotUserlist_Unload()
'
End Sub


'Adds a user to the bot
Public Function AddUser(UserName As String, GlobalFlags As String) As Byte ' : AddStack "UserList_AddUser(" & UserName & ", " & GlobalFlags & ")"
Dim UsNum As Long, ScNum As Long
  If ServerNickLen = 0 Then ServerNickLen = 30
  If Len(UserName) > ServerNickLen Then AddUser = AU_TooLong: ExtReply = Trim(Str(ServerNickLen)): Exit Function
  If IsValidNick(UserName) = False Then AddUser = AU_InvalidNick: Exit Function
  UsNum = GetUserNum(UserName)
  If UsNum > 0 Then AddUser = AU_UserExists: Exit Function
  AddBotUser UserName
  BotUsers(BotUserCount).Flags = GlobalFlags
  AddUser = AU_Success
  
  'Check script Hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.AddedUser Then
      RunScriptX ScNum, "AddedUser", UserName
    End If
  Next ScNum
End Function

'Removes a user from the bot
Public Function RemUser(UserName As String) ' : AddStack "UserList_RemUser(" & UserName & ")"
Dim RealNick As String, ScNum As Long
  RealNick = GetRealNick(UserName)
  If RealNick = "" Then RemUser = RU_UserNotFound: Exit Function
  RemoveUser UserName
  RemUser = RU_Success
  
  'Check script Hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.RemovedUser Then
      RunScriptX ScNum, "RemovedUser", UserName
    End If
  Next ScNum
End Function

Sub ChangeNick(ByVal OldNick As String, ByVal NewNick As String) ' : AddStack "UserList_ChangeNick(" & OldNick & ", " & NewNick & ")"
Dim OneLine As String, InRightSection As Boolean, LeaveThisOneOut As Boolean
Dim ChangedSomething As Boolean, SavOldNick As String, SavNewNick As String
Dim FileName As String, u As Long, ChNum As Long, UsNum As Long, Buf As String
Dim FNum1 As Integer, FNum2 As Integer, FoundOne As Boolean, ScNum As Long
  UsNum = GetUserNum(OldNick)
  If UsNum = 0 Then Exit Sub
  
  BotUsers(UsNum).Name = NewNick
  
  NotesChangeNick OldNick, NewNick
  
  'Change nick of channel users
  For ChNum = 1 To ChanCount
    For UsNum = 1 To Channels(ChNum).UserCount
      If LCase(Channels(ChNum).User(UsNum).RegNick) = LCase(OldNick) Then Channels(ChNum).User(UsNum).RegNick = NewNick
    Next UsNum
  Next ChNum
  
  'Change nick of party line users
  FoundOne = False
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If LCase(SocketItem(u).RegNick) = LCase(OldNick) Then
        If GetSockFlag(u, SF_LocalVisibleUser) = SF_YES Then
          ToBotNet 0, "nc " & BotNetNick & " " & SocketItem(u).OrderSign & " " & NewNick
          If Not FoundOne Then SpreadMessage 0, SocketItem(u).PLChannel, MakeMsg(MSG_PLNick, OldNick, NewNick)
          SocketItem(u).RegNick = NewNick
          FoundOne = True
        End If
      End If
    End If
  Next u
  
  'Check script Hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.ChangedNick Then
      RunScriptX ScNum, "ChangedNick", OldNick, NewNick
    End If
  Next ScNum
End Sub


'Adds a host to a user
Public Function AddHost(ExecutingBotUser As Long, RegUser As String, HostToAdd As String) As Byte ' : AddStack "UserList_AddHost(" & ExecutingBotUser & ", " & RegUser & ", " & HostToAdd & ")"
Dim u As Long, u2 As Long, UsNum As Long, Messg As String
Dim RemovedOne As Boolean, ScNum As Long
  UsNum = GetUserNum(RegUser)
  '+ 17.07.2004 - User dürfen keine Host bei leuten mit höheren Flags adden -> Gefahr des erschleichens von Rechtem
  If ExecutingBotUser <> UsNum Then
    If Not MatchFlags(BotUser(ExecutingBotUser, 2), "+s") Then
      If MatchFlags(BotUser(UsNum, 2), "+s") Then
        AddHost = AH_DENIED
        Exit Function
      Else
        If Not MatchFlags(BotUser(ExecutingBotUser, 2), "+n") Then
          If MatchFlags(BotUser(UsNum, 2), "+n") Then
            AddHost = AH_DENIED
            Exit Function
          Else
            If Not MatchFlags(BotUser(ExecutingBotUser, 2), "+m") Then
              If MatchFlags(BotUser(UsNum, 2), "+m") Then
                AddHost = AH_DENIED
                Exit Function
              End If
            End If
          End If
        End If
      End If
    End If
  End If
  '+ /17.07.2004
  If UsNum = 0 Then AddHost = AH_UserNotFound: Exit Function
  If IsValidHostmask(HostToAdd) = False Then AddHost = AH_InvalidHost: Exit Function
  For u = 1 To BotUsers(UsNum).HostMaskCount
    If LCase(BotUsers(UsNum).HostMasks(u)) = LCase(HostToAdd) Then AddHost = AH_AlreadyThere: ExtReply = BotUsers(UsNum).Name: Exit Function
  Next u
  Messg = SearchUserFromHostmask3(HostToAdd)
  If (Messg <> "") Then AddHost = AH_MatchingUser: ExtReply = Messg: Exit Function
  If ExecutingBotUser > 0 Then
    Messg = SearchHigherMatchingUser(ExecutingBotUser, HostToAdd)
    If (Messg <> "") Then AddHost = AH_MatchingUser: ExtReply = Messg: Exit Function
  End If
  'Remove matching hostmasks
  Do
    RemovedOne = False
    For u = 1 To BotUsers(UsNum).HostMaskCount
      If MatchWM(HostToAdd, BotUsers(UsNum).HostMasks(u)) Then
        RemHost BotUsers(UsNum).Name, BotUsers(UsNum).HostMasks(u)
        RemovedOne = True
        Exit For
      End If
    Next u
  Loop While RemovedOne
  If BotUsers(UsNum).HostMaskCount = 20 Then AddHost = AH_TooManyHosts: Exit Function
  
  'No errors -> add host
  BotUsers(UsNum).HostMaskCount = BotUsers(UsNum).HostMaskCount + 1
  BotUsers(UsNum).HostMasks(BotUsers(UsNum).HostMaskCount) = HostToAdd
  AddHost = AH_Success: ExtReply = BotUsers(UsNum).Name

  'Check script Hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.AddedHost Then
      RunScriptX ScNum, "AddedHost", RegUser, HostToAdd
    End If
  Next ScNum
End Function

'Removes a host from a user
Public Function RemHost(UserName As String, Hostmask As String) As Byte ' : AddStack "UserList_RemHost(" & UserName & ", " & Hostmask & ")"
Dim LeaveThisOneOut As Boolean, ChangedSomething As Boolean
Dim CheckLine As String, CheckHost As String, CheckNick As String
Dim FileName As String, UsNum As Long, ScNum As Long
Dim HostNum As Long, u As Long, FNum1 As Integer, FNum2 As Integer

  UsNum = GetUserNum(UserName)
  If UsNum = 0 Then RemHost = RH_UserNotFound: Exit Function    'User not found
  HostNum = GetHostNum(UsNum, Hostmask)
  If HostNum = 0 Then RemHost = RH_HostNotFound: Exit Function  'Hostmask not found
  
  For u = HostNum To BotUsers(UsNum).HostMaskCount - 1
    BotUsers(UsNum).HostMasks(u) = BotUsers(UsNum).HostMasks(u + 1)
  Next u
  BotUsers(UsNum).HostMaskCount = BotUsers(UsNum).HostMaskCount - 1
  
  'Check script Hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.RemovedHost Then
      RunScriptX ScNum, "RemovedHost", UserName, Hostmask
    End If
  Next ScNum
End Function


Public Sub SetUserData(UsNum As Long, Entry As String, Value As String)
Dim OldIndex As String, NewIndex As String, OldData As String, NewData As String
Dim CurPiece As String, CurEntry As String, CurLen As Long, CurPos As Long
Dim u As Long, FoundIt As Boolean
  If (UsNum < 1) Or (UsNum > BotUserCount) Or (Entry = "") Then Exit Sub
  OldIndex = Param(BotUsers(UsNum).UserData, 1)
  OldData = GetRest(BotUsers(UsNum).UserData, 2)
  CurPos = 1
  For u = 1 To ParamXCount(OldIndex, ",")
    CurPiece = ParamX(OldIndex, ",", u)
    CurEntry = LCase(ParamX(CurPiece, ":", 1))
    Dim sDummy As String
    sDummy = ParamX(CurPiece, ":", 2)
    If sDummy = "" Then sDummy = 0
    CurLen = CLng(sDummy)
    If CurEntry = LCase(Entry) Then
      If Value <> "" Then
        NewIndex = NewIndex + Entry & ":" & Trim(Str(Len(Value))) & ","
        NewData = NewData + Value
        FoundIt = True
      End If
    Else
      NewIndex = NewIndex + CurPiece & ","
      NewData = NewData + Mid(OldData, CurPos, CurLen)
    End If
    CurPos = CurPos + CurLen
  Next u
  If (FoundIt = False) And (Value <> "") Then
    NewIndex = NewIndex + Entry & ":" & Trim(Str(Len(Value))) & ","
    NewData = NewData + Value
  End If
  
  If NewIndex <> "" Then
    NewIndex = Left(NewIndex, Len(NewIndex) - 1)
    BotUsers(UsNum).UserData = NewIndex & " " & NewData
  Else
    BotUsers(UsNum).UserData = ""
  End If
End Sub

Public Function GetUserData(UsNum As Long, Entry As String, Default As String) As String
Dim OldIndex As String, OldData As String, u As Long
Dim CurPiece As String, CurEntry As String, CurLen As Long, CurPos As Long
  If (UsNum < 1) Or (UsNum > BotUserCount) Or (Entry = "") Then GetUserData = Default: Exit Function
  OldIndex = Param(BotUsers(UsNum).UserData, 1)
  OldData = GetRest(BotUsers(UsNum).UserData, 2)
  CurPos = 1
  For u = 1 To ParamXCount(OldIndex, ",")
    CurPiece = ParamX(OldIndex, ",", u)
    CurEntry = LCase(ParamX(CurPiece, ":", 1))
    CurLen = CLng(ParamX(CurPiece, ":", 2))
    If CurEntry = LCase(Entry) Then
      GetUserData = Mid(OldData, CurPos, CurLen)
      Exit Function
    End If
    CurPos = CurPos + CurLen
  Next u
  GetUserData = Default
End Function

Public Sub ReadUserList()
  If Dir(HomeDir & "UserList.txt") <> "" Then
    ReadNewUserList
  ElseIf Dir(HomeDir & "extusers.ini") <> "" Then
    ReadOldUserList
  End If
End Sub


Public Sub ReadOldUserList()
  Dim CheckLine As String, CheckHost As String, CheckNick As String, FileNum As Integer
  Dim unum As Long, ChFlags As String, SeeChannel As String, SeenTime As String
  Dim SeenChan As String, SeenMessage As String, Last As Date, blubb As Long
  Dim ChangeLine As String, Rest As String
  
  On Local Error Resume Next
  
  FileNum = FreeFile
  Open HomeDir & "Hostmask.txt" For Input As #FileNum
  If Err.Number = 0 Then
    On Error GoTo Erro
    Do While Not EOF(FileNum)
      Line Input #FileNum, CheckLine
blubb = 5
      If Trim(CheckLine) <> "" Then
blubb = 6
        CheckHost = Param(CheckLine, 1)
blubb = 7
        CheckNick = Param(CheckLine, 2)
blubb = 8
        If CheckNick <> "" Then
          unum = GetUserNum(CheckNick)
blubb = 9
          If unum = 0 Then AddBotUser CheckNick: unum = BotUserCount
blubb = 10
          BotUsers(unum).HostMaskCount = BotUsers(unum).HostMaskCount + 1
blubb = 11
          BotUsers(unum).HostMasks(BotUsers(unum).HostMaskCount) = CheckHost
        End If
blubb = 12
      End If
    Loop
    Close #FileNum
  Else
    Close #FileNum
    Err.Clear
  End If
blubb = 13
  unum = 0
  On Local Error Resume Next
  Open HomeDir & "ExtUsers.ini" For Input As #FileNum
  If Err.Number = 0 Then
    On Error GoTo Erro
blubb = 14
    Do While Not EOF(FileNum)
blubb = 15
      Line Input #FileNum, SeeChannel
blubb = 16
      If Left(SeeChannel, 1) = "[" Then
blubb = 17
        CheckNick = Trim(MakeNormalNick(Mid(SeeChannel, 2, Len(SeeChannel) - 2)))
blubb = 18
        If CheckNick <> "" Then
          unum = GetUserNum(CheckNick)
blubb = 19
          If unum = 0 Then AddBotUser CheckNick: unum = BotUserCount
blubb = 20
          BotUsers(unum).Flags = GetPPString(CheckNick, "Flags", "", HomeDir & "ExtUsers.ini")
          If MatchFlags(BotUsers(unum).Flags, "-b") Then
            SetUserData unum, "colors", IIf(LCase(GetPPString(CheckNick, "Colors", "no", HomeDir & "ExtUsers.ini")) = "yes", SF_YES, SF_NO)
          Else
            Rest = GetPPString(CheckNick, "ConnectAddress", "", HomeDir & "ExtUsers.ini")
            SetUserData unum, UD_LinkAddr, Rest
          End If
          BotUsers(unum).BotFlags = GetPPString(CheckNick, "BotFlags", "", HomeDir & "ExtUsers.ini")
          BotUsers(unum).Password = GetPPString(CheckNick, "Password", "", HomeDir & "ExtUsers.ini")
          BotUsers(unum).PLChannel = CLng(GetPPString(CheckNick, "PLChannel", "0", HomeDir & "ExtUsers.ini"))
          SetUserData unum, "sflags", GetPPString(CheckNick, "SocketFlags", "", HomeDir & "ExtUsers.ini")
          SetUserData unum, "comment", GetPPString(CheckNick, "Comment", "", HomeDir & "ExtUsers.ini")
          SetUserData unum, "info", GetPPString(CheckNick, "Info", "", HomeDir & "ExtUsers.ini")
        Else
          unum = 0
        End If
      End If
blubb = 25
      If (LCase(Left(SeeChannel, 7)) = "chflags") And (unum > 0) Then
blubb = 26
        ChFlags = Trim(Right(SeeChannel, Len(SeeChannel) - InStr(SeeChannel, "=")))
blubb = 27
        SeeChannel = Trim(Left(SeeChannel, InStr(SeeChannel, "=") - 1))
blubb = 28
        SeeChannel = Right(SeeChannel, Len(SeeChannel) - 7)
blubb = 29
        BotUsers(unum).ChannelFlagCount = BotUsers(unum).ChannelFlagCount + 1
blubb = 30
        BotUsers(unum).ChannelFlags(BotUsers(unum).ChannelFlagCount).Channel = SeeChannel
blubb = 31
        BotUsers(unum).ChannelFlags(BotUsers(unum).ChannelFlagCount).Flags = ChFlags
      End If
    Loop
    Close #FileNum
  Else
    Close #FileNum
    Err.Clear
  End If
  
  For unum = 1 To BotUserCount
blubb = 301
    SeenTime = GetPPString(BotUsers(unum).Name, "LastSeen", "", HomeDir & "ExtUsers.ini")
    If SeenTime <> "" Then
blubb = 302
      SeenChan = GetPPString(BotUsers(unum).Name, "LastSeenChan", "", HomeDir & "ExtUsers.ini")
blubb = 303
      SeenMessage = GetPPString(BotUsers(unum).Name, "LastSeenMessage", "", HomeDir & "ExtUsers.ini")
      On Local Error Resume Next
      Last = CDate(Val(SeenTime))
      If Err.Number = 0 Then
blubb = 304
        WriteSeenEntry CheckNick, "", Last, SeenChan, SeenMessage, IIf(BotUsers(unum).HostMaskCount > 0, Replace(Mask(BotUsers(unum).HostMasks(1), 10), "*", ""), "unknown@hostmask.com")
        Err.Clear
      End If
      On Error GoTo Erro
blubb = 305
      DeletePPString BotUsers(unum).Name, "LastSeen", HomeDir & "ExtUsers.ini"
      DeletePPString BotUsers(unum).Name, "LastSeenChan", HomeDir & "ExtUsers.ini"
      DeletePPString BotUsers(unum).Name, "LastSeenMessage", HomeDir & "ExtUsers.ini"
    End If
    'Check bots
    If MatchFlags(BotUsers(unum).Flags, "-b") Then
      'Delete bot flags of non-bots
      If BotUsers(unum).BotFlags <> "" Then
blubb = 306
        DeletePPString BotUsers(unum).Name, "BotFlags", HomeDir & "ExtUsers.ini"
        BotUsers(unum).BotFlags = ""
      End If
      'Set SocketFlags if not set
      'If Len(GetUserData(unum, "sflags", "")) <> (SavedEnd - SavedStart + 1) Then
      '  GiveSockFlags unum
      'End If
    Else
      'Move +h flag from "Flags" to "BotFlags"
      If InStr(BotUsers(unum).Flags, "h") > 0 Then
        BotUsers(unum).Flags = GetChattrResult(BotUsers(unum).Flags, "-h")
        BotUsers(unum).BotFlags = GetBotattrResult(BotUsers(unum).BotFlags, "+h")
        WritePPString BotUsers(unum).Name, "Flags", BotUsers(unum).Flags, HomeDir & "ExtUsers.ini"
        WritePPString BotUsers(unum).Name, "BotFlags", BotUsers(unum).BotFlags, HomeDir & "ExtUsers.ini"
      End If
    End If
  Next unum
  
  'Check userfile
  For unum = 1 To BotUserCount
    'Encrypt users' passwords on the fly
    If (BotUsers(unum).Password <> "") And MatchFlags(BotUsers(unum).Flags, "-b") Then
      If Left(BotUsers(unum).Password, 1) <> "¤" Then
blubb = 23
        BotUsers(unum).Password = EncryptIt(BotUsers(unum).Password)
blubb = 24
        WritePPString BotUsers(unum).Name, "Password", BotUsers(unum).Password, HomeDir & "ExtUsers.ini"
      End If
    End If
    ChangeLine = ""
    'Ensure that super owners have +ijmnptw and owners have +ijmptw
    If MatchFlags(BotUsers(unum).Flags, "+s") Then ChangeLine = ChangeLine & "+ijmnptw"
    If MatchFlags(BotUsers(unum).Flags, "+n") Then ChangeLine = ChangeLine & "+ijmptw"
    If ChangeLine <> "" Then
      If GetChattrResult(BotUsers(unum).Flags, ChangeLine) <> BotUsers(unum).Flags Then
blubb = 21
        BotUsers(unum).Flags = GetChattrResult(BotUsers(unum).Flags, ChangeLine)
blubb = 22
        WritePPString BotUsers(unum).Name, "Flags", BotUsers(unum).Flags, HomeDir & "ExtUsers.ini"
      End If
    End If
  Next unum
Exit Sub

Erro:
  MsgBox "hhm.. das wäre der fehler nr. " & CStr(blubb) & "!" & " - VB meint dazu: " & Err.Number & " - " & Err.Description
  End
End Sub

'Reads the new userlist format
Public Sub ReadNewUserList()
Dim FileNum As Integer, unum As Long, Line As String, Rest As String
Dim CurNick As String, TheFlags As String, u As Long, blubb As Long
Dim ChangeLine As String
On Local Error Resume Next
  unum = 0
blubb = 0
  FileNum = FreeFile
  Open HomeDir & "UserList.txt" For Input As #FileNum
  If Err.Number = 0 Then
    On Error GoTo Erroar
    Do While Not EOF(FileNum)
      Line Input #FileNum, Line
      Select Case Param(Line, 1)
        Case "---"
          CurNick = Param(Line, 2)
          unum = GetUserNum(CurNick)
          If unum = 0 Then AddBotUser CurNick: unum = BotUserCount
          TheFlags = GetRest(Line, 3)
          For u = 1 To ParamXCount(TheFlags, ",")
            Rest = ParamX(TheFlags, ",", u)
            If ParamCount(Rest) = 1 Then
              BotUsers(unum).Flags = Rest
            Else
              If Param(Rest, 2) = "bot" Then
                BotUsers(unum).BotFlags = Param(Rest, 1)
              Else
                BotUsers(unum).ChannelFlagCount = BotUsers(unum).ChannelFlagCount + 1
                BotUsers(unum).ChannelFlags(BotUsers(unum).ChannelFlagCount).Flags = Param(Rest, 1)
                BotUsers(unum).ChannelFlags(BotUsers(unum).ChannelFlagCount).Channel = Param(Rest, 2)
              End If
            End If
          Next u
        Case "p"
          BotUsers(unum).Password = GetRest(Line, 2)
        Case "h"
          For u = 2 To ParamCount(Line)
            BotUsers(unum).HostMaskCount = BotUsers(unum).HostMaskCount + 1
            BotUsers(unum).HostMasks(BotUsers(unum).HostMaskCount) = Param(Line, u)
          Next u
        Case "d"
          BotUsers(unum).UserData = GetRest(Line, 2)
      End Select
    Loop
    Close #FileNum
  Else
    Close #FileNum
    Err.Clear
  End If
  On Error GoTo Erroar
  
  For unum = 1 To BotUserCount
blubb = 1
    'Check bots
    If MatchFlags(BotUsers(unum).Flags, "-b") Then
      'Delete bot flags of non-bots
      If BotUsers(unum).BotFlags <> "" Then BotUsers(unum).BotFlags = ""
    Else
      'Move +h flag from "Flags" to "BotFlags"
      If InStr(BotUsers(unum).Flags, "h") > 0 Then
        BotUsers(unum).Flags = Replace(BotUsers(unum).Flags, "h", "")
        BotUsers(unum).BotFlags = GetBotattrResult(BotUsers(unum).BotFlags, "+h")
      End If
    End If
blubb = 2
    'Encrypt users' passwords on the fly
    If (BotUsers(unum).Password <> "") And MatchFlags(BotUsers(unum).Flags, "-b") Then
      If Left(BotUsers(unum).Password, 1) <> "¤" Then BotUsers(unum).Password = EncryptIt(BotUsers(unum).Password)
    End If
blubb = 3
    'Ensure that super owners have +fijmnptw and owners have +fijmptw
    ChangeLine = ""
    If MatchFlags(BotUsers(unum).Flags, "+s") Then ChangeLine = ChangeLine & "+fijmnptw"
    If MatchFlags(BotUsers(unum).Flags, "+n") Then ChangeLine = ChangeLine & "+fijmptw"
    If ChangeLine <> "" Then BotUsers(unum).Flags = GetChattrResult(BotUsers(unum).Flags, ChangeLine)
  Next unum
Exit Sub

Erroar:
  MsgBox "Ein Fehler ist in Zeile " & CStr(blubb) & " aufgetreten!" & " - VB meint dazu: " & Err.Number & " - '" & Err.Description & "'. Zeile: " & Line
  End
End Sub


Sub WriteUserList()
Dim FileName As String, FileNumber As Integer, u As Long, ErrLine As Long
Dim TheFlags As String, ChSum As Long
  
  'Calculate Userlist checksum
  ChSum = BotUserCount * 10
  For u = 1 To BotUserCount
    ChSum = AddCheck(ChSum, BotUsers(u).Name)
    ChSum = AddCheck(ChSum, CombineAllFlags(u))
    If BotUsers(u).Password <> "" Then ChSum = AddCheck(ChSum, BotUsers(u).Password)
    If BotUsers(u).HostMaskCount > 0 Then ChSum = AddCheck(ChSum, CombineAllHosts(u))
    If BotUsers(u).UserData <> "" Then ChSum = AddCheck(ChSum, BotUsers(u).UserData)
  Next u
  
  'Re-write userlist only if checksum has changed or bot is exitting
  If (ChSum = OldCheckSum) Or (Exitting = True) Then Exit Sub
  OldCheckSum = ChSum
  
  FileName = HomeDir & "UserList.txt"
  WaitForAccess FileName
  AddAccessedFile FileName
  On Error GoTo WULErr
  FileNumber = FreeFile: Open FileName For Output As #FileNumber
  Print #FileNumber, "' This is an AnGeL Bot userfile. Please don't change anything"
  Print #FileNumber, "' in here unless you know what you're doing."
  Print #FileNumber, ""
  For u = 1 To BotUserCount
    Print #FileNumber, "--- " & BotUsers(u).Name & " " & CombineAllFlags(u)
    If BotUsers(u).Password <> "" Then Print #FileNumber, "p " & BotUsers(u).Password
    If BotUsers(u).HostMaskCount > 0 Then Print #FileNumber, "h " & CombineAllHosts(u)
    If BotUsers(u).UserData <> "" Then Print #FileNumber, "d " & BotUsers(u).UserData
  Next u
  Close #FileNumber
  RemAccessedFile FileName
Exit Sub
WULErr:
  Close #FileNumber
  RemAccessedFile FileName
  Err.Clear
  'Stop
End Sub

Public Function GetUserNum(UserName As String) As Long ' : AddStack "Routines_GetUserNum(" & UserName & ")"
Dim u As Long
  If UserName = "" Then GetUserNum = 0: Exit Function
  For u = 1 To BotUserCount
    If LCase(UserName) = LCase(BotUsers(u).Name) Then GetUserNum = u: Exit Function
  Next u
  GetUserNum = 0
End Function

Public Function GetHostNum(UsNum As Long, Hostmask As String) As Long ' : AddStack "Routines_GetHostNum(" & UsNum & ", " & Hostmask & ")"
Dim u As Long
  For u = 1 To BotUsers(UsNum).HostMaskCount
    If LCase(Hostmask) = LCase(BotUsers(UsNum).HostMasks(u)) Then GetHostNum = u: Exit Function
  Next u
  GetHostNum = 0
End Function

Public Sub AddBotUser(UserName As String) ' : AddStack "Routines_AddBotUser(" & UserName & ")"
  BotUserCount = BotUserCount + 1
  If BotUserCount > UBound(BotUsers()) Then ReDim Preserve BotUsers(UBound(BotUsers()) + 5)
  BotUsers(BotUserCount).Name = UserName
  BotUsers(BotUserCount).ChannelFlagCount = 0
  BotUsers(BotUserCount).HostMaskCount = 0
  BotUsers(BotUserCount).BotFlags = ""
  BotUsers(BotUserCount).Flags = ""
  BotUsers(BotUserCount).Password = ""
  BotUsers(BotUserCount).UserData = ""
  GiveSockFlags BotUserCount
End Sub

Function GetRealNick(Nick As String) As String
  Dim u As Long
  u = GetUserNum(Nick)
  If u > 0 Then GetRealNick = BotUsers(u).Name Else GetRealNick = ""
End Function

Public Function GetUserFlags(Nick As String) As String
  Dim u As Long
  GetUserFlags = ""
  If Nick <> "" Then
    u = GetUserNum(Nick)
    If u > 0 Then GetUserFlags = BotUsers(u).Flags
  End If
End Function

Public Function AllFlags(Nick As String) As String
  Dim u As Long, u2 As Long, Flags As String
  AllFlags = ""
  If Nick <> "" Then
    u = GetUserNum(Nick)
    If u = 0 Then Exit Function
    Flags = BotUsers(u).Flags
    For u2 = 1 To BotUsers(u).ChannelFlagCount
      Flags = CombineFlags(Flags, "+" & BotUsers(u).ChannelFlags(u2).Flags)
    Next u2
    AllFlags = Flags
  End If
End Function

Public Function GetUserChanFlags(Nick As String, Channel As String) As String
  Dim u As Long, u2 As Long, NewFlags As String, s As Byte
  GetUserChanFlags = ""
  If Nick <> "" Then
    u = GetUserNum(Nick)
    If u > 0 Then
      NewFlags = BotUsers(u).Flags
      For u2 = 1 To BotUsers(u).ChannelFlagCount
        If LCase(BotUsers(u).ChannelFlags(u2).Channel) = LCase(Channel) Then
          NewFlags = CombineFlags(BotUsers(u).Flags, "+" & BotUsers(u).ChannelFlags(u2).Flags + ChattrChanges(BotUsers(u).ChannelFlags(u2).Flags))
          Exit For
        End If
      Next u2
      GetUserChanFlags = NewFlags
    End If
  End If
End Function

Public Function GetUserChanFlags2(u As Long, Channel As String) As String
  Dim u2 As Long, NewFlags As String, s As Byte
  If u > 0 Then
    NewFlags = BotUsers(u).Flags
    For u2 = 1 To BotUsers(u).ChannelFlagCount
      If LCase(BotUsers(u).ChannelFlags(u2).Channel) = LCase(Channel) Then NewFlags = CombineFlags(BotUsers(u).Flags, "+" & BotUsers(u).ChannelFlags(u2).Flags + ChattrChanges(BotUsers(u).ChannelFlags(u2).Flags)): Exit For
    Next u2
    GetUserChanFlags2 = NewFlags
  Else
    GetUserChanFlags2 = ""
  End If
End Function

Public Function ChanFlagsOnly(u As Long, Channel As String) As String
  Dim Flags As String, u2 As Long
  ChanFlagsOnly = ""
  If u > 0 Then
    For u2 = 1 To BotUsers(u).ChannelFlagCount
      If LCase(BotUsers(u).ChannelFlags(u2).Channel) = LCase(Channel) Then
        ChanFlagsOnly = BotUsers(u).ChannelFlags(u2).Flags
        Exit Function
      End If
    Next u2
  End If
End Function

Function UserExist(Nick As String) As Boolean
  Dim u As Long
  For u = 1 To BotUserCount
    If LCase(BotUsers(u).Name) = LCase(Nick) Then
      UserExist = True
      Exit Function
    End If
  Next u
  UserExist = False
End Function

Sub RemoveUser(UserName As String)
  Dim LeaveThisOneOut As Boolean, ChangedSomething As Boolean, Rest As String
  Dim CheckLine As String, CheckHost As String, CheckNick As String, UsNum As Long
  Dim FileName As String, u As Long, u2 As Long, FNum1 As Integer, FNum2 As Integer

  UsNum = GetUserNum(UserName)
  If UsNum = 0 Then Exit Sub 'User not found
  
  For u = UsNum To BotUserCount - 1
    BotUsers(u) = BotUsers(u + 1)
  Next u
  BotUserCount = BotUserCount - 1
  
  'Update channel user numbers
  For u = 1 To ChanCount
    For u2 = 1 To Channels(u).UserCount
      If Channels(u).User(u2).UserNum = UsNum Then
        Channels(u).User(u2).UserNum = 0
      ElseIf Channels(u).User(u2).UserNum > UsNum Then
        Channels(u).User(u2).UserNum = Channels(u).User(u2).UserNum - 1
      End If
    Next u2
  Next u

  'Update Socket user numbers
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If SocketItem(u).UserNum = UsNum Then
        SocketItem(u).UserNum = 0
      ElseIf SocketItem(u).UserNum > UsNum Then
        SocketItem(u).UserNum = SocketItem(u).UserNum - 1
      End If
    End If
  Next u
  
  'Clear seen entries of this user
  ClearSeenEntries
End Sub

Public Function Chattr(ChattrNick As String, ChangeLine As String, Optional WhoDid As String = "") As Byte ' : AddStack "UserList_Chattr(" & ChattrNick & ", " & ChangeLine & ")"
Dim UsNum As Long, Chan As String, OldFlags As String, NewFlags As String
Dim ChgPos As String, HostNum As Long, u As Long, Rest As String
Dim ChangeFlags As String
  UsNum = GetUserNum(ChattrNick)
  If UsNum = 0 Then Exit Function
  If IsValidChannel(Left(Param(ChangeLine, 2), 1)) Then Chan = Param(ChangeLine, 2): ChgPos = "ChFlags" & Chan Else ChgPos = "Flags"
  If Chan = "" Then
    OldFlags = BotUsers(UsNum).Flags
    ChangeFlags = Param(ChangeLine, 1) + ChattrChanges(Param(ChangeLine, 1))
  Else
    OldFlags = ChanFlagsOnly(UsNum, Chan)
    ChangeFlags = Param(ChangeLine, 1)
    Rest = GetPosFlags(ChangeFlags)
    'Check if channel flag is possible (i.e., +b or +p is only a global flag)
    If MatchFlags(Rest, "+i") Or MatchFlags(Rest, "+j") Or MatchFlags(Rest, "+p") Or MatchFlags(Rest, "+s") Or MatchFlags(Rest, "+t") Or MatchFlags(Rest, "+w") Then
      Chattr = CH_NoChanFlag: Exit Function
    End If
  End If

  NewFlags = GetChattrResult(OldFlags, ChangeFlags)
  ExtReply = NewFlags
  If NewFlags = OldFlags Then Chattr = CH_NoChanges: Exit Function
  
  If Chan = "" Then
    BotUsers(UsNum).Flags = NewFlags
  Else
    HostNum = 0
    If NewFlags <> "" Then
      For u = 1 To BotUsers(UsNum).ChannelFlagCount
        If LCase(BotUsers(UsNum).ChannelFlags(u).Channel) = LCase(Chan) Then BotUsers(UsNum).ChannelFlags(u).Flags = NewFlags: HostNum = 1: Exit For
      Next u
      If HostNum = 0 Then
        BotUsers(UsNum).ChannelFlagCount = BotUsers(UsNum).ChannelFlagCount + 1
        BotUsers(UsNum).ChannelFlags(BotUsers(UsNum).ChannelFlagCount).Channel = Chan
        BotUsers(UsNum).ChannelFlags(BotUsers(UsNum).ChannelFlagCount).Flags = NewFlags
      End If
    Else
      For u = 1 To BotUsers(UsNum).ChannelFlagCount
        If LCase(BotUsers(UsNum).ChannelFlags(u).Channel) = LCase(Chan) Then HostNum = u: Exit For
      Next u
      If HostNum > 0 Then
        For u = HostNum To BotUsers(UsNum).ChannelFlagCount - 1
          BotUsers(UsNum).ChannelFlags(u) = BotUsers(UsNum).ChannelFlags(u + 1)
        Next u
        BotUsers(UsNum).ChannelFlagCount = BotUsers(UsNum).ChannelFlagCount - 1
      End If
    End If
  End If
  
  If Chan = "" Then
    'Adjust Socket flags
    If MatchFlags(OldFlags, "-m") And MatchFlags(NewFlags, "+m") Then
      SetMasterSockFlags UsNum, True
    ElseIf MatchFlags(OldFlags, "+m") And MatchFlags(NewFlags, "-m") Then
      SetMasterSockFlags UsNum, False
    End If
    
    'User got an upgrade :)
    Rest = ""
    If MatchFlags(OldFlags, "-s") And MatchFlags(NewFlags, "+s") Then
      Rest = "a SUPER OWNER"
    ElseIf MatchFlags(OldFlags, "-n") And MatchFlags(NewFlags, "+n") Then
      Rest = "an OWNER"
    ElseIf MatchFlags(OldFlags, "-m") And MatchFlags(NewFlags, "+m") Then
      Rest = "a MASTER"
    ElseIf MatchFlags(OldFlags, "-t") And MatchFlags(NewFlags, "+t") Then
      Rest = "a BOTNET MASTER"
    End If
    
    'Set flags and notify affected user on the party line
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If (SocketItem(u).RegNick = ChattrNick) And (SocketItem(u).OnBot = BotNetNick) Then
          If Rest <> "" Then TU u, "10*** Congratulations! You are now " & Rest & " of this bot!!!"
          
          'User got a downgrade ;)
          If MatchFlags(SocketItem(u).Flags, "+s") And MatchFlags(NewFlags, "-s") Then TU u, "4*** PANG!!! You are no longer a super owner of this bot!"
          If MatchFlags(SocketItem(u).Flags, "+n") And MatchFlags(NewFlags, "-n") Then TU u, "4*** PANG!!! You are no longer an owner of this bot!"
          If MatchFlags(SocketItem(u).Flags, "+m") And MatchFlags(NewFlags, "-m") Then TU u, "4*** PANG!!! You are no longer a master of this bot!"
          If MatchFlags(SocketItem(u).Flags, "+t") And MatchFlags(NewFlags, "-t") Then TU u, "4*** PANG!!! You are no longer a botnet master of this bot!"
          
          If GetLevelSign(SocketItem(u).Flags) <> GetLevelSign(NewFlags) Then
            If MatchFlags(SocketItem(u).Flags, "-b") And (GetSockFlag(u, SF_LocalVisibleUser) = SF_YES) Then ToBotNet 0, "j " & BotNetNick & " " & SocketItem(u).RegNick & " " & LongToBase64(SocketItem(u).PLChannel) & " " & GetLevelSign(NewFlags) + SocketItem(u).OrderSign & " " & Mask(SocketItem(u).Hostmask, 10)
          End If
          
          'Set sockflags and normal flags
          ReadSockFlags u
          SocketItem(u).Flags = NewFlags
        End If
      End If
    Next u
  End If
  Chattr = CH_Success
End Function

Public Function MatchUserFlags(UserNum As Long, MatchString As String) As Boolean
  MatchUserFlags = False
  If UserNum <= BotUserCount Then
    MatchUserFlags = MatchFlags(BotUsers(UserNum).Flags, MatchString)
  End If
End Function
Public Function BotUser(Nr, GetWhat As Byte, Optional OptParam) ' : AddStack "PartyHandler_BotUser(" & Nr & ", " & GetWhat & ", " & CStr(OptParam) & ")"
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
    Case Else
      BotUser = ""
  End Select
End Function

Sub GiveSockFlags(unum As Long)
  SetUserData unum, "sflags", DefaultSockFlags(unum)
End Sub

Sub SetMasterSockFlags(UsNum As Long, Give As Boolean)
  If Give Then
    SetUserSockFlag UsNum, SF_PrivToBot, SF_YES
    SetUserSockFlag UsNum, SF_UserCommands, SF_YES
  Else
    SetUserSockFlag UsNum, SF_PrivToBot, SF_NO
    SetUserSockFlag UsNum, SF_UserCommands, SF_NO
  End If
End Sub

Sub ReadSockFlags(TheSock As Long)
Dim u As Byte, TheFlags As String, SockFlag As String
  If Not IsValidSocket(TheSock) Then Exit Sub
  TheFlags = GetUserData(SocketItem(TheSock).UserNum, "sflags", DefaultSockFlags(SocketItem(TheSock).UserNum))
  For u = SavedStart To SavedEnd
    SetSockFlag TheSock, u, Mid(TheFlags, u - SavedStart + 1, 1)
  Next u
End Sub

Function DefaultSockFlags(unum As Long) As String
  DefaultSockFlags = String(7, SF_YES) + String(2, IIf(MatchFlags(BotUsers(unum).Flags, "+m"), SF_YES, SF_NO))
End Function

Sub SetUserSockFlag(UsNum As Long, ByVal SF As Byte, Value As String)
Dim TheSockFlags As String
  If UsNum > BotUserCount Then Exit Sub
  SF = SF - SavedStart + 1
  TheSockFlags = GetUserData(UsNum, "sflags", "")
  If Len(TheSockFlags) < SF Then
    TheSockFlags = TheSockFlags + String(SF - Len(TheSockFlags), " ")
  End If
  Mid(TheSockFlags, SF, 1) = Value
  SetUserData UsNum, "sflags", TheSockFlags
End Sub

Sub SaveSockFlags(TheSock As Long)
Dim u As Long, TheFlags As String, SockFlag As String
  If Not IsValidSocket(TheSock) Then Exit Sub
  TheFlags = SocketItem(TheSock).SocketFlags
  SetUserData SocketItem(TheSock).UserNum, "sflags", Mid(TheFlags, SavedStart, SavedEnd - SavedStart + 1)
  For u = 1 To SocketCount
    If u <> TheSock Then
      If IsValidSocket(u) Then
        If (SocketItem(u).OnBot = BotNetNick) And (SocketItem(u).RegNick = SocketItem(TheSock).RegNick) Then
          ReadSockFlags u
        End If
      End If
    End If
  Next u
End Sub

