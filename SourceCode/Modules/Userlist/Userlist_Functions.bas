Attribute VB_Name = "Userlist_Functions"
',-======================- ==-- -  -
'|   AnGeL - Userlist - Functions
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public Function GetPosFlags(Flags As String) As String
  Dim u As Long, Positif As Boolean, NewFlags As String
  Positif = True
  For u = 1 To Len(Flags)
    Select Case Mid(Flags, u, 1)
      Case "+"
        Positif = True
      Case "-"
        Positif = False
      Case Else
        If Positif Then NewFlags = NewFlags & Mid(Flags, u, 1)
    End Select
  Next u
  GetPosFlags = NewFlags
End Function

Public Function MatchFlags(Flags As String, MatchString As String) As Boolean
  MatchFlags = (CombineFlags(Flags, MatchString) = CombineFlags("", Flags))
End Function

Public Function CombineFlags(Flags As String, ChangeLine As String) As String
  Dim NewFlags As String, u As Long, Plus As Boolean
  NewFlags = Flags
  Plus = True
  For u = 1 To Len(ChangeLine)
    Select Case LCase(Mid(ChangeLine, u, 1))
      Case "+"
        Plus = True
      Case "-"
        Plus = False
      Case Else
        If Plus Then NewFlags = AddFlag(NewFlags, Mid(ChangeLine, u, 1))
        If Not Plus Then NewFlags = RemFlag(NewFlags, Mid(ChangeLine, u, 1))
    End Select
  Next u
  CombineFlags = NewFlags
End Function

Public Function AddFlag(ByVal Flags As String, Flag As String) As String
  Dim u As Byte, NewFlags As String
  Flags = Flags & Flag
  For u = 1 To 26
    If InStr(Flags, Chr(96 + u)) > 0 Then NewFlags = NewFlags & Chr(96 + u)
    If InStr(Flags, Chr(64 + u)) > 0 Then NewFlags = NewFlags & Chr(64 + u)
  Next u
  AddFlag = NewFlags
End Function

Public Function RemFlag(ByVal Flags As String, Flag As String) As String
  Dim u As Byte, NewFlags As String
  For u = 1 To 26
    If InStr(Flags, Chr(96 + u)) > 0 Then If Chr(96 + u) <> Flag Then NewFlags = NewFlags & Chr(96 + u)
    If InStr(Flags, Chr(64 + u)) > 0 Then If Chr(64 + u) <> Flag Then NewFlags = NewFlags & Chr(64 + u)
  Next u
  RemFlag = NewFlags
End Function


Public Function CheckForInvalidFlags(ByVal Flags As String) As String
  Dim InvalidOnes As String, u As Long
  For u = 1 To Len(Flags)
    Select Case Mid(Flags, u, 1)
      Case "+", "-", "a", "b", "d", "f", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "v", "w", "x", "A" To "Z"
        InvalidOnes = InvalidOnes
      Case Else
        If InvalidOnes <> "" Then InvalidOnes = InvalidOnes & "," & Mid(Flags, u, 1) Else InvalidOnes = Mid(Flags, u, 1)
    End Select
  Next u
  CheckForInvalidFlags = InvalidOnes
End Function

Function SearchUserFromHostmask2(Hostmask As String) As String
  Dim u As Long, u2 As Long
  Dim Nick As String, Ident As String, Domain As String
  SearchUserFromHostmask2 = ""
  SplitHostmask Hostmask, Nick, Ident, Domain
  For u = 1 To BotUserCount
    For u2 = 1 To BotUsers(u).HostMaskCount
      If MatchWM2Ex(BotUsers(u).HostMasks(u2), Nick, Ident, Domain) Then
        SearchUserFromHostmask2 = BotUsers(u).Name
        Exit Function
      End If
    Next u2
  Next u
End Function

Function SearchHigherMatchingUser(ThisUser As Long, NewHostmask As String) As String
  Dim u As Long, u2 As Long, ThisLev As Long, OtherLev As Long
  Dim Nick As String, Ident As String, Domain As String
  SearchHigherMatchingUser = ""
  ThisLev = LevelNumber(BotUsers(ThisUser).Flags)
  SplitHostmask NewHostmask, Nick, Ident, Domain
  For u = 1 To BotUserCount
    OtherLev = LevelNumber(BotUsers(u).Flags)
    If OtherLev > ThisLev Then
      For u2 = 1 To BotUsers(u).HostMaskCount
        If MatchWM2Ex(BotUsers(u).HostMasks(u2), Nick, Ident, Domain) Then
          If HostRating(NewHostmask) > HostRating(BotUsers(u).HostMasks(u2)) Then
            SearchHigherMatchingUser = BotUsers(u).Name
            Exit Function
          End If
        End If
      Next u2
    End If
  Next u
End Function

Function SearchUserFromHostmask3(Hostmask As String) As String
  Dim u As Long, u2 As Long
  SearchUserFromHostmask3 = ""
  For u2 = 1 To BotUserCount
    For u = 1 To BotUsers(u2).HostMaskCount
      If LCase(BotUsers(u2).HostMasks(u)) = LCase(Hostmask) Then
        SearchUserFromHostmask3 = BotUsers(u2).Name
        Exit Function
      End If
    Next u
  Next u2
End Function

Function SearchUserFromHostmask(Hostmask As String) As String
  Dim FullHM As String, u As Long, u2 As Long, curRating As Long, bestRating As Long
  Dim Nick As String, Ident As String, Domain As String
  SplitHostmask Hostmask, Nick, Ident, Domain
  If StrictHost = False Then
    If InStr("~-+^=", Left(Ident, 1)) > 0 And Len(Ident) > 1 Then Ident = Mid(Ident, 2)
    FullHM = Nick & "!" & Ident & "@" & Domain
  Else
    FullHM = Hostmask
  End If
  bestRating = 0: BotUserNum = 0
  For u = 1 To BotUserCount
    For u2 = 1 To BotUsers(u).HostMaskCount
      If SimpleMatch(BotUsers(u).HostMasks(u2), FullHM) Then
        If MatchHostEx(BotUsers(u).HostMasks(u2), Nick, Ident, Domain) Then
          curRating = HostRating(BotUsers(u).HostMasks(u2))
          If curRating > bestRating Then
            bestRating = curRating
            BotUserNum = u
          End If
        End If
      End If
    Next u2
  Next u
  If BotUserNum = 0 Then SearchUserFromHostmask = "" Else SearchUserFromHostmask = BotUsers(BotUserNum).Name
End Function

Function HostRating(Hostmask As String) As Long
  Dim Nick As String, Ident As String, Domain As String
  SplitHostmask Hostmask, Nick, Ident, Domain
  HostRating = Len(Replace(Domain, "*", "")) * 3 + Len(Replace(Ident, "*", "")) * 2 + Len(Replace(Nick, "*", ""))
End Function

Function SearchMatchedHostmask(Hostmask As String) As String
  Dim u As Long, u2 As Long, sNick As String, SIdent As String, SHost As String
  Dim curRating As Long, bestRating As Long
  bestRating = 0
  For u = 1 To BotUserCount
    For u2 = 1 To BotUsers(u).HostMaskCount
      If MatchHost(BotUsers(u).HostMasks(u2), Hostmask) Then
        curRating = HostRating(BotUsers(u).HostMasks(u2))
        If curRating > bestRating Then
          bestRating = curRating
          If StrictHost = False Then
            If MatchWM(BotUsers(u).HostMasks(u2), Hostmask) = False Then
              SearchMatchedHostmask = Mask(BotUsers(u).HostMasks(u2), 6)
            Else
              SearchMatchedHostmask = BotUsers(u).HostMasks(u2)
            End If
          Else
            SearchMatchedHostmask = BotUsers(u).HostMasks(u2)
          End If
        End If
      End If
    Next u2
  Next u
  If bestRating = 0 Then SearchMatchedHostmask = ""
End Function

Public Function MatchHost(WM As String, FullHM As String) As Boolean
  Dim Nick1 As String, Domain1 As String, Ident1 As String
  Dim Nick2 As String, Domain2 As String, Ident2 As String
  If WM = "" Or FullHM = "" Then MatchHost = False: Exit Function
  If InStr(WM, "!") = 0 Or InStr(FullHM, "!") = 0 Or InStr(WM, "@") = 0 Or InStr(FullHM, "@") = 0 Then MatchHost = False: Exit Function
  SplitHostmask WM, Nick1, Ident1, Domain1
  SplitHostmask FullHM, Nick2, Ident2, Domain2
  If StrictHost = False Then
    If InStr("~-+^=", Left(Ident2, 1)) > 0 Then If Len(Ident2) > 1 Then Ident2 = Mid(Ident2, 2)
  End If
  If SimpleMatch(Ident1, Ident2) = False Then MatchHost = False: Exit Function
  If SimpleMatch(Domain1, Domain2) = False Then MatchHost = False: Exit Function
  If SimpleMatch(Nick1, Nick2) = False Then MatchHost = False: Exit Function
  MatchHost = True
End Function

Public Function MatchHostEx(WM As String, Nick2 As String, Ident2 As String, Domain2 As String) As Boolean ' : AddStack "Base2_MatchHostEx(" & WM & ", " & Nick2$ & ", " & Ident2$ & ", " & Domain2$ & ")"
  Dim Nick1 As String, Ident1 As String, Domain1 As String
  If WM = "" Then MatchHostEx = False: Exit Function
  If InStr(WM, "!") = 0 Or InStr(WM, "@") = 0 Then MatchHostEx = False: Exit Function
  SplitHostmask WM, Nick1, Ident1, Domain1
  If SimpleMatch(Ident1, Ident2) = False Then MatchHostEx = False: Exit Function
  If SimpleMatch(Domain1, Domain2) = False Then MatchHostEx = False: Exit Function
  If SimpleMatch(Nick1, Nick2) = False Then MatchHostEx = False: Exit Function
  MatchHostEx = True
End Function

Public Function MatchWM(WM As String, FullHM As String) As Boolean
  Dim Nick1 As String, Domain1 As String, Ident1 As String
  Dim Nick2 As String, Domain2 As String, Ident2 As String
  If WM = "" Or FullHM = "" Then MatchWM = False: Exit Function
  If InStr(WM, "!") = 0 Or InStr(FullHM, "!") = 0 Or InStr(WM, "@") = 0 Or InStr(FullHM, "@") = 0 Then MatchWM = False: Exit Function
  SplitHostmask WM, Nick1, Ident1, Domain1
  SplitHostmask FullHM, Nick2, Ident2, Domain2
  If SimpleMatch(Ident1, Ident2) = False Then MatchWM = False: Exit Function
  If SimpleMatch(Domain1, Domain2) = False Then MatchWM = False: Exit Function
  If SimpleMatch(Nick1, Nick2) = False Then MatchWM = False: Exit Function
  MatchWM = True
End Function

Public Function MatchWM2(WM As String, FullHM As String) As Boolean
  Dim Nick1 As String, Domain1 As String, Ident1 As String
  Dim Nick2 As String, Domain2 As String, Ident2 As String
  If WM = "" Or FullHM = "" Then MatchWM2 = False: Exit Function
  If InStr(WM, "!") = 0 Or InStr(FullHM, "!") = 0 Or InStr(WM, "@") = 0 Or InStr(FullHM, "@") = 0 Then MatchWM2 = False: Exit Function
  SplitHostmask WM, Nick1, Ident1, Domain1
  SplitHostmask FullHM, Nick2, Ident2, Domain2
  If (SimpleMatch(Ident1, Ident2) = False) And (SimpleMatch(Ident2, Ident1) = False) Then MatchWM2 = False: Exit Function
  If (SimpleMatch(Domain1, Domain2) = False) And (SimpleMatch(Domain2, Domain1) = False) Then MatchWM2 = False: Exit Function
  If (SimpleMatch(Nick1, Nick2) = False) And (SimpleMatch(Nick2, Nick1) = False) Then MatchWM2 = False: Exit Function
  MatchWM2 = True
End Function

Public Function MatchWM2Ex(WM As String, Nick2 As String, Ident2 As String, Domain2 As String) As Boolean
  Dim Nick1 As String, Domain1 As String, Ident1 As String
  If WM = "" Then MatchWM2Ex = False: Exit Function
  If InStr(WM, "!") = 0 Or InStr(WM, "@") = 0 Then MatchWM2Ex = False: Exit Function
  SplitHostmask WM, Nick1, Ident1, Domain1
  If (SimpleMatch(Ident1, Ident2) = False) And (SimpleMatch(Ident2, Ident1) = False) Then MatchWM2Ex = False: Exit Function
  If (SimpleMatch(Domain1, Domain2) = False) And (SimpleMatch(Domain2, Domain1) = False) Then MatchWM2Ex = False: Exit Function
  If (SimpleMatch(Nick1, Nick2) = False) And (SimpleMatch(Nick2, Nick1) = False) Then MatchWM2Ex = False: Exit Function
  MatchWM2Ex = True
End Function

Public Function SimpleMatch(Wild As String, Full As String) As Boolean
  Dim Wild2 As String, Full2 As String
  Wild2 = LCase(MakeININick(Wild))
  Full2 = LCase(MakeININick(Full))
  SimpleMatch = Full2 Like Wild2
End Function

Function IsValidNick(Nick As String) As Boolean
  Dim u As Long
  If Nick = "" Then IsValidNick = False: Exit Function
  If Asc(Left(Nick, 1)) < 65 Or Asc(Left(Nick, 1)) > 125 Then IsValidNick = False: Exit Function
  For u = 1 To Len(Nick)
    If (Asc(Mid(Nick, u, 1)) < 65 Or Asc(Mid(Nick, u, 1)) > 125) And InStr("-0123456789", Mid(Nick, u, 1)) = 0 Then IsValidNick = False: Exit Function
  Next u
  IsValidNick = True
End Function

Function IsValidIdent(Ident As String) As Boolean
  Dim u As Long
  If Ident = "" Then IsValidIdent = False: Exit Function
  If Asc(Left(Ident, 1)) < 65 Or Asc(Left(Ident, 1)) > 125 Then IsValidIdent = False: Exit Function
  For u = 1 To Len(Ident)
    If (Asc(Mid(Ident, u, 1)) < 65 Or Asc(Mid(Ident, u, 1)) > 122) And InStr("-0123456789", Mid(Ident, u, 1)) = 0 Then IsValidIdent = False: Exit Function
  Next u
  If IsNumeric(Mid(Ident, 1, 1)) = True Then IsValidIdent = False: Exit Function
  IsValidIdent = True
End Function

Function IsIdentified(Nick As String, RegNick As String) As Boolean
  Dim u As Long, u2 As Long
  IsIdentified = False
'B10'
'  For u = 1 To ChanCount
'    For u2 = 1 To Channels(u).UserCount
'      If Channels(u).User(u2).RegNick = RegNick And Channels(u).User(u2).Nick = Nick Then
'        IsIdentified = Channels(u).User(u2).Identified
'        Exit For
'      End If
'    Next u2
'  Next u
'  If IsIdentified = True Then Exit Function
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If SocketItem(u).RegNick = RegNick Then
        IsIdentified = True
        Exit For
      End If
    End If
  Next u
End Function


Function CombineAllFlags(UsNum As Long) As String
Dim TheFlags As String, u As Long
  If BotUsers(UsNum).Flags <> "" Then TheFlags = BotUsers(UsNum).Flags
  For u = 1 To BotUsers(UsNum).ChannelFlagCount
    If TheFlags <> "" Then TheFlags = TheFlags & ","
    TheFlags = TheFlags + BotUsers(UsNum).ChannelFlags(u).Flags & " " & BotUsers(UsNum).ChannelFlags(u).Channel
  Next u
  If BotUsers(UsNum).BotFlags <> "" Then
    If TheFlags <> "" Then TheFlags = TheFlags & ","
    TheFlags = TheFlags + BotUsers(UsNum).BotFlags & " bot"
  End If
  CombineAllFlags = TheFlags
End Function

Function CombineAllHosts(UsNum As Long) As String
Dim TheHosts As String, u As Long
  For u = 1 To BotUsers(UsNum).HostMaskCount
    If TheHosts <> "" Then TheHosts = TheHosts & " "
    TheHosts = TheHosts + BotUsers(UsNum).HostMasks(u)
  Next u
  CombineAllHosts = TheHosts
End Function

Public Sub ConvertMatchString(Line As String, Hostmask As String, GlobalFlags As String, ChannelFlags As String, ChannelMatch As String, BotFlags As String)
Dim u As Long, SavedPart As String
  Hostmask = ""
  GlobalFlags = ""
  ChannelFlags = ""
  ChannelMatch = ""
  BotFlags = ""
  For u = 1 To ParamCount(Line)
    If SavedPart = "" Then
      'Flags were specified... save flags and wait for more information
      If InStr("+-", Left(Param(Line, u), 1)) > 0 Then
        SavedPart = Param(Line, u)
      'Valid hostmask given...
      ElseIf IsValidHostmask(Param(Line, u)) = True Then
        Hostmask = Param(Line, u)
      'Nothing recognizable given... assume that it's a nick
      Else
        If Hostmask = "" Then Hostmask = Param(Line, u) & "!*@*"
      End If
    Else
      'New flags given... assume that last flags were global flags and save new flags
      If InStr("+-", Left(Param(Line, u), 1)) > 0 Then
        GlobalFlags = SavedPart
        SavedPart = Param(Line, u)
      'Channel given... last flags were channel flags
      ElseIf IsValidChannel(Left(Param(Line, u), 1)) = True Then
        ChannelFlags = SavedPart
        ChannelMatch = Param(Line, u) + IIf(IsValidChannel(Param(Line, u)), "*", "")
        SavedPart = ""
      'String "channel"... last flags were channel flags
      ElseIf LCase(Param(Line, u)) = "channel" Then
        ChannelFlags = SavedPart
        ChannelMatch = "#*"
        SavedPart = ""
      'String "bot"... last flags were bot flags
      ElseIf LCase(Param(Line, u)) = "bot" Then
        BotFlags = SavedPart
        SavedPart = ""
      'Valid hostmask given... assume that last flags were global flags
      ElseIf IsValidHostmask(Param(Line, u)) = True Then
        GlobalFlags = SavedPart
        Hostmask = Param(Line, u)
        SavedPart = ""
      'Nothing recognizable given... assume that last flags were global flags and that this is a nick
      Else
        If GlobalFlags = "" Then GlobalFlags = SavedPart
        If Hostmask = "" Then Hostmask = Param(Line, u) & "!*@*"
        SavedPart = ""
      End If
    End If
  Next u
  If SavedPart <> "" Then
    GlobalFlags = SavedPart
  End If
End Sub
'Returns whether the given Hostmask is valid
Function IsValidHostmask(HM As String) As Boolean
  Dim PosAus As Long, PosAt As Long
  PosAus = InStr(HM, "!")
  PosAt = InStr(HM, "@")
  If (PosAus > PosAt) Or (PosAus = 0) Or (PosAt = 0) Then IsValidHostmask = False: Exit Function
  IsValidHostmask = True
End Function


Public Function ChattrChanges(Rest As String) As String
Dim ChangeLine As String
  If MatchFlags(Rest, "+d") Then ChangeLine = ChangeLine & "-fo"
  If MatchFlags(Rest, "+j") Then ChangeLine = ChangeLine & "+i"
  If MatchFlags(Rest, "+k") Then ChangeLine = ChangeLine & "-fo"
  If MatchFlags(Rest, "+m") Then ChangeLine = ChangeLine & "+ipft-dk"
  If MatchFlags(Rest, "+n") Then ChangeLine = ChangeLine & "+ijpfomt-dk"
  If MatchFlags(Rest, "+o") Then ChangeLine = ChangeLine & "+f-dk"
  If MatchFlags(Rest, "+s") Then ChangeLine = ChangeLine & "+ijpfomnt-dk"
  ChattrChanges = ChangeLine
End Function

'Combines old flags to flags after CHATTR
Function GetChattrResult(FlagsBefore As String, ChangeLine As String) As String
Dim NewFlags As String, PlusOrMinus As Byte, u As Long
  NewFlags = FlagsBefore
  For u = 1 To Len(ChangeLine)
    Select Case Mid(ChangeLine, u, 1)
      Case "+": PlusOrMinus = 1
      Case "-": PlusOrMinus = 2
      Case "a", "b", "d", "f", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "v", "w", "x", "A" To "Z"
        If PlusOrMinus = 1 Then NewFlags = AddFlag(NewFlags, Mid(ChangeLine, u, 1))
        If PlusOrMinus = 2 Then NewFlags = RemFlag(NewFlags, Mid(ChangeLine, u, 1))
    End Select
  Next u
  GetChattrResult = NewFlags
End Function

'Combines old flags to flags after BOTATTR
Function GetBotattrResult(FlagsBefore As String, ChangeLine As String) As String
Dim NewFlags As String, PlusOrMinus As Byte, u As Long
  NewFlags = FlagsBefore
  For u = 1 To Len(ChangeLine)
    Select Case LCase(Mid(ChangeLine, u, 1))
      Case "+": PlusOrMinus = 1
      Case "-": PlusOrMinus = 2
      Case "a", "h", "l", "s", "r", "i"
        If PlusOrMinus = 1 Then NewFlags = AddFlag(NewFlags, LCase(Mid(ChangeLine, u, 1)))
        If PlusOrMinus = 2 Then NewFlags = RemFlag(NewFlags, LCase(Mid(ChangeLine, u, 1)))
    End Select
  Next u
  GetBotattrResult = NewFlags
End Function

'Extracts nick, ident and host out of a hostmask
Public Sub SplitHostmask(Text As String, Nick As String, Ident As String, Host As String) 'NoTrap' : AddStack "Conversions_SplitHostmask(" & Text & ", " & Nick & ", " & Ident & ", " & Host & ")"
Dim XP As Long, ParP As Long, LastP As Long
  ParP = 1: LastP = 1
  Nick = "": Ident = "": Host = ""
  If (InStr(Text, "@") = 0) And (InStr(Text, "!") = 0) Then
    Host = Text
  Else
    If InStr(Text, "@") > 0 Then
      Do
        XP = InStr(ParP, Text & "@", "@") + 1
        If XP = 1 Then Host = Mid(Text, LastP): Exit Do
        LastP = ParP
        ParP = XP
      Loop
    End If
    If InStr(Text, "!") > 0 Then
      Nick = Left(Text, InStr(Text, "!") - 1)
    End If
    If Nick <> "" Then
      If Host <> "" Then
        Ident = Mid(Text, Len(Nick) + 2, Len(Text) - Len(Host) - Len(Nick$) - 2)
      Else
        Ident = Mid(Text, Len(Nick) + 2, Len(Text) - Len(Nick$) - 1)
      End If
    Else
      If Host <> "" Then
        Ident = Mid(Text, 1, Len(Text) - Len(Host) - 1)
      Else
        Ident = Mid(Text, 1, Len(Text) - Len(Host))
      End If
    End If
  End If
  If Nick = "" Then Nick = "*"
  If Ident = "" Then Ident = "*"
  If Host = "" Then Host = "*"
End Sub

'Converts a full hostmask to a wildcard mask
Public Function Mask(Hostmask As String, Nr As Long) As String
  Dim Nick As String, Domain As String, StarDomain As String, Host As String, Ident As String, u As Long, Count As Long, lastpoint As Long, IsValidIP As Boolean
  If Hostmask = "" Then Mask = "x!x@x": Exit Function
  SplitHostmask Hostmask, Nick, Ident, Domain
  Host = Left(Domain, InStr(Domain, "."))
  If InStr(InStr(Domain, ".") + 1, Domain, ".") = 0 Then Host = ""
  IsValidIP = True
  For u = 1 To Len(Domain)
    Select Case Mid(Domain, u, 1)
      Case "0" To "9"
      Case ".": Count = Count + 1: lastpoint = u
      Case Else: IsValidIP = False
    End Select
  Next u
  If Count <> 3 Then IsValidIP = False
  Domain = Right(Domain, Len(Domain) - Len(Host))
  If Host <> "" Then
    If Not IsValidIP Then
      StarDomain = "*." & Domain
    Else
      StarDomain = Left(Host + Domain, lastpoint) & "*"
    End If
  Else
    StarDomain = Domain
  End If
  Select Case Nr
    Case 1, 3, 6, 8
      If InStr("~-+^=", Left(Ident, 1)) > 0 Then If Len(Ident) > 1 Then Ident$ = Mid(Ident, 2)
    Case 21, 23
      If InStr("~-+^=", Left(Ident, 1)) > 0 Then If Len(Ident) > 1 Then Ident$ = Mid(Ident, 2)
      If StrictHost = True Then Ident = "*" & Ident
  End Select
  If Nick = "" Then Nick = "*"
  Select Case Nr
    Case 0: Mask = "*!" & Ident & "@" & Host + Domain                           ' *!~ident@1234.test.de
    Case 1: Mask = "*!*" & Right(Ident, 9) & "@" & Host + Domain                ' *!*ident@1234.test.de
    Case 2: Mask = "*!*@" & Host & Domain                                       ' *!*@1234.test.de
    Case 3: Mask = "*!*" & Right(Ident, 9) & "@" & StarDomain                   ' *!*ident@*.test.de
    Case 4: Mask = "*!*@" & StarDomain                                          ' *!*@*.test.de
    Case 5: Mask = Nick$ & "!" & Ident & "@" & Host + Domain                    ' nick!~ident@1234.test.de
    Case 6: Mask = Nick$ & "!*" & Right(Ident, 9) & "@" & Host + Domain         ' nick!*ident@1234.test.de
    Case 7: Mask = Nick$ & "!*@" & Host & Domain                                ' nick!*@1234.test.de
    Case 8: Mask = Nick$ & "!*" & Right(Ident, 9) & "@" & StarDomain            ' nick!*ident@*.test.de
    Case 9: Mask = Nick$ & "!*@" & StarDomain                                   ' nick!*@*.test.de
    Case 10: Mask = Ident$ & "@" & Host + Domain                                ' ~ident@1234.test.de
    Case 11: Mask = Host & Domain                                               ' 1234.test.de
    Case 12: Mask = Ident                                                       ' ~ident
    Case 13: Mask = Nick                                                        ' nick
    Case 14: Mask = Nick & "!" & Ident & "@"                                    ' nick!~ident@
    Case 21: Mask = "*!" & Ident & "@" & Host & Domain                          ' *!*ident@1234.test.de / *!ident@1234.test.de
    Case 23: Mask = "*!" & Ident & "@" & StarDomain                             ' *!*ident@*.test.de    / *!ident@*.test.de
  End Select
End Function

