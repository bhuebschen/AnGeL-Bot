Attribute VB_Name = "Partyline_Functions"
Option Explicit

Sub SpreadPartylineChanEvent(TheEvent As String, Channel As String, WhoNick As String, WhoRegNick As String, WhoFlags As String, TarNick As String, TarRegNick As String, TarFlags As String, InfoLine As String)
  Dim u As Long
  Dim Z As Long
  Select Case TheEvent
    Case "nick"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMJP", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 7*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " is now known as " & TarNick
              End If
            Next Z
          End If
        End If
      Next u
    Case "join"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMJP", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 3*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " has joined"
              End If
            Next Z
          End If
        End If
      Next u
    Case "part"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMJP", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 3*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " has left" & IIf(InfoLine <> "", " (" & InfoLine & ")", "")
              End If
            Next Z
          End If
        End If
      Next u
    Case "quit"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMJP", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMJP", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 3*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " has quit" & IIf(InfoLine <> "", " (" & InfoLine & ")", "")
              End If
            Next Z
          End If
        End If
      Next u
    Case "kick"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMKB", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMKB", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMKB", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 7*** " & TarNick & IIf(TarRegNick <> "", " (" & TarRegNick & ")", "") & " was kicked by " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & IIf(InfoLine <> "", " - (" & InfoLine & ")", "")
              End If
            Next Z
          End If
        End If
      Next u
    Case "mode"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMKB", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMKB", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMKB", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 6*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " sets mode: " & InfoLine
              End If
            Next Z
          End If
        End If
      Next u
    Case "textflood"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMFP", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMFP", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMFP", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 4*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " is textflooding"
              End If
            Next Z
          End If
        End If
      Next u
    Case "nickflood"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMFP", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMFP", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMFP", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 4*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " is nickflooding"
              End If
            Next Z
          End If
        End If
      Next u
    Case "chantalk"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMPT", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMPT", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMPT", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") <" & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & "> " & InfoLine
              End If
            Next Z
          End If
        End If
      Next u
    Case "chanaction"
      For u = 1 To SocketCount
        If IsValidSocket(u) Then
          If GetUserData(SocketItem(u).UserNum, "BMPT", "°none°") <> "°none°" And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, Channel), "+o")) Then
            For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMPT", "")))
              If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMPT", "")), Z) = LCase(Channel) Then
                If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, "14(" & Channel & ") 6*** " & WhoNick & IIf(WhoRegNick <> "", " (" & WhoRegNick & ")", "") & " " & InfoLine
              End If
            Next Z
          End If
        End If
      Next u
    Case Else
      Trace "SpreadPartylineChanEvent ->", TheEvent, WhoNick, TarNick, InfoLine
  End Select
End Sub


Public Sub SpreadMessage(vsock As Long, PLChannel As Long, Line As String) ' : AddStack "Socks_SpreadMessage(" & vsock & ", " & PLChannel & ", " & Line & ")"
Dim u As Long, u2 As Long, OnBot As String, Nick As String, Message As String
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If (vsock <> u) Then
        If ((PLChannel > -1) And (SocketItem(u).PLChannel = PLChannel)) Or (PLChannel = -1) Then
          If (GetSockFlag(u, SF_Status) = SF_Status_Party) And (SocketItem(u).OnBot = BotNetNick) Then TU u, Line
        End If
      End If
    End If
  Next u
  If Left(Strip(Line), 8) = "[" & Time & "] " Then PutLog "| " & Mid(Strip(Line), 8) Else PutLog "|  " & Strip(Line)
End Sub

Public Sub SpreadMessageEx(vsock As Long, PLChannel As Long, SockFlag As Byte, Line As String) ' : AddStack "Socks_SpreadMessageEx(" & vsock & ", " & PLChannel & ", " & SockFlag & ", " & Line & ")"
Dim u As Long, u2 As Long, OnBot As String, Nick As String, Message As String
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If (vsock <> u) Then
        If ((PLChannel > -1) And (SocketItem(u).PLChannel = PLChannel)) Or (PLChannel = -1) Then
          If (GetSockFlag(u, SF_Status) = SF_Status_Party) And (SocketItem(u).OnBot = BotNetNick) Then TUEx u, SockFlag, Line
        End If
      End If
    End If
  Next u
  If Left(Strip(Line), 8) = "[" & Time & "] " Then PutLog "| " & Mid(Strip(Line), 8) Else PutLog "|  " & Strip(Line)
End Sub

Public Sub SpreadFlagMessage(vsock As Long, NeededFlags As String, Line As String) '' : AddStack "Socks_SpreadFlagMessage(" & vsock & ", " & NeededFlags & ", " & Line & ")"
Dim u As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If vsock <> u Then
        If (GetSockFlag(u, SF_Status) = SF_Status_Party) And (SocketItem(u).OnBot = BotNetNick) Then
          If NeededFlags <> "" Then
            If GetChattrResult(SocketItem(u).Flags, NeededFlags) = SocketItem(u).Flags Then TU u, Line
          Else
            TU u, Line
          End If
        End If
      End If
    End If
  Next u
  If Left(Strip(Line), 8) = "[" & Time & "] " Then PutLog "| " & Mid(Strip(Line), 8) Else PutLog "|  " & Strip(Line)
  Output Strip(Line) + vbCrLf
End Sub


Public Sub SpreadChanMessage(NeedDefChan As String, Line As String) '' : AddStack "Socks_SpreadChanMessage(" & NeededFlags & ", " & Line & ")"
  Dim u As Long
  Dim Z As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If (GetUserData(SocketItem(u).UserNum, "BMSS", "") = "1") And (SocketItem(u).OnBot = BotNetNick) And (MatchFlags(GetUserChanFlags(SocketItem(u).RegNick, NeedDefChan), "+o")) Then
        If NeedDefChan <> "" Then
          For Z = 1 To ParamCount(LCase(GetUserData(SocketItem(u).UserNum, "BMDF", "")))
            If Param(LCase(GetUserData(SocketItem(u).UserNum, "BMDF", "")), Z) = LCase(NeedDefChan) Then
              'SpreadFlagMessage 0, "+m", Param(LCase(GetUserData(SocketItem(u).UserNum, "BMDF", "")), Z) & "=" & LCase(NeedDefChan) & "(" & (Param(LCase(GetUserData(SocketItem(u).UserNum, "BMDF", "")), Z) = LCase(NeedDefChan)) & ")|" & GetSockFlag(u, SF_Status) & "=" & SF_Status_Party & "(" & (GetSockFlag(u, SF_Status) = SF_Status_Party) & ")"
              If (GetSockFlag(u, SF_Status) = SF_Status_Party) Then TU u, Line
            End If
          Next Z
        End If
      End If
    End If
  Next u
End Sub


Public Sub SpreadFlagMessageEx(vsock As Long, NeededFlags As String, SockFlag As Byte, Line As String) ' : AddStack "Socks_SpreadFlagMessageEx(" & vsock & ", " & NeededFlags & ", " & SockFlag & ", " & Line & ")"
Dim u As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If vsock <> u Then
        If (GetSockFlag(u, SF_Status) = SF_Status_Party) And (SocketItem(u).OnBot = BotNetNick) Then
          If NeededFlags <> "" Then
            If GetChattrResult(SocketItem(u).Flags, NeededFlags) = SocketItem(u).Flags Then TUEx u, SockFlag, Line
          Else
            TU u, Line
          End If
        End If
      End If
    End If
  Next u
  If Left(Strip(Line), 8) = "[" & Time & "] " Then PutLog "| " & Mid(Strip(Line), 8) Else PutLog "|  " & Strip(Line)
  Output Strip(Line) + vbCrLf
End Sub

'Converts flags to a Cl number (for help system); better flags -> higher number
Function CLevel(GFlags As String, CFlags As String) As Integer
Dim LevNum As Integer
  LevNum = Cl_User
  If MatchFlags(GFlags, "+w") Then LevNum = (LevNum Or Cl_What)
  If MatchFlags(CFlags, "+o") Then LevNum = (LevNum Or Cl_Op)
  'If MatchFlags(CFlags, "+h") Then LevNum = (LevNum Or Cl_HOp)
  If MatchFlags(GFlags, "+t") Then LevNum = (LevNum Or Cl_Net)
  If MatchFlags(GFlags, "+m") Then
    LevNum = (LevNum Or Cl_Mas)
  Else
    If MatchFlags(CFlags, "+m") Then LevNum = (LevNum Or Cl_CMas)
  End If
  If MatchFlags(GFlags, "+n") Then
    LevNum = (LevNum Or Cl_Own)
  Else
    If MatchFlags(CFlags, "+n") Then LevNum = (LevNum Or Cl_COwn)
  End If
  If MatchFlags(GFlags, "+s") Then LevNum = (LevNum Or Cl_SOwn)
  CLevel = LevNum
End Function

'Converts flags to a level number; better flags -> higher number
Function LevelNumber(Flags As String) As Byte
  If MatchFlags(Flags, "+n") Then LevelNumber = 3: Exit Function
  If MatchFlags(Flags, "+m") Then LevelNumber = 2: Exit Function
  If MatchFlags(Flags, "+o") Then LevelNumber = 1: Exit Function
  LevelNumber = 0
End Function

'Converts flags to a string
Function LevelString(GFlags As String, CFlags As String) As String
Dim Lev As String
  Lev = "3User"
  If MatchFlags(CFlags, "+o") Then Lev = "3Channel op"
  If MatchFlags(GFlags, "+t") Then Lev = "10Botnet master"
  If MatchFlags(GFlags, "+m") Then Lev = "12Master"
  If MatchFlags(GFlags, "+n") Then Lev = "4Owner"
  If MatchFlags(GFlags, "+s") Then Lev = "4Super owner"
  LevelString = Lev
End Function

Public Function MakeTelnetColor(ByVal Line As String, vsock As Long) As String  ' Makes Telnet-ANSI Colors =)' : AddStack "SCMsExtensions_MakeTelnetColor(" & Line & ")"
  Dim x As Long, r As String, R2 As String, TelnetCol As String
  Dim NLine As String
  NLine = Line
  x = 0
  Do
    x = x + 1
    x = InStr(x, NLine, "")
    If x > 0 Then
      r = Mid(NLine, x + 1, 1)
      If IsNumeric(r) Then
      If IsNumeric(Mid(NLine, x + 1, 2)) Then r = Mid(NLine, x + 1, 2)
      If InStr(r, ",") > 0 Then r = Mid(NLine, x + 1, 1)
      If InStr(r, " ") > 0 Then r = Mid(NLine, x + 1, 1)
      If InStr(r, ".") > 0 Then r = Mid(NLine, x + 1, 1)
      If InStr(r, "-") > 0 Then r = Mid(NLine, x + 1, 1)
      If IsNumeric(r) Then
        If CInt(r) < 0 Then r = CInt(r) * -1
        Select Case CInt(r)
          Case 0: TelnetCol = "[1;37m"
          Case 1: TelnetCol = "[2;30m"
          Case 2: TelnetCol = "[2;34m"
          Case 3: TelnetCol = "[2;32m"
          Case 4: TelnetCol = "[1;31m"
          Case 5: TelnetCol = "[2;31m"
          Case 6: TelnetCol = "[2;35m"
          Case 7: TelnetCol = "[2;33m"
          Case 8: TelnetCol = "[1;33m"
          Case 9: TelnetCol = "[1;32m"
          Case 10: TelnetCol = "[2;36m"
          Case 11: TelnetCol = "[1;36m"
          Case 12: TelnetCol = "[1;34m"
          Case 13: TelnetCol = "[1;35m"
          Case 14: TelnetCol = "[1;30m"
          Case 15: TelnetCol = "[2;37m"
        End Select
        If Mid(NLine, x + Len(r) + 1, 1) = "," Then
          R2 = Mid(NLine, x + Len(r) + 2, 1)
          If IsNumeric(Mid(NLine, x + Len(r) + 2, 2)) Then R2 = Mid(NLine, x + Len(r) + 2, 2)
          If InStr(R2, ",") > 0 Then R2 = Mid(NLine, x + Len(r) + 2, 1)
          If IsNumeric(R2) Then
            Select Case CByte(R2)
              Case 0: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";47m"
              Case 1: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";40m"
              Case 2: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";44m"
              Case 3: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";42m"
              Case 4: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";41m"
              Case 5: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";41m"
              Case 6: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";45m"
              Case 7: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";43m"
              Case 8: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";43m"
              Case 9: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";42m"
              Case 10: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";46m"
              Case 11: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";46m"
              Case 12: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";44m"
              Case 13: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";45m"
              Case 14: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";47m"
              Case 15: TelnetCol = Mid(TelnetCol, 1, Len(TelnetCol) - 1) & ";40m"
            End Select
            If r = R2 Then TelnetCol = "[1;" & Mid(TelnetCol, 3)
            NLine = Mid(NLine, 1, x - 1) & TelnetCol & Mid(NLine, x + Len(r) + 2 + Len(R2))
          Else
            NLine = Mid(NLine, 1, x - 1) & TelnetCol & Mid(NLine, x + 1 + Len(r))
          End If
        Else
          NLine = Mid(NLine, 1, x - 1) & TelnetCol & Mid(NLine, x + Len(r) + 1)
        End If
      Else
        NLine = Mid(NLine, 1, x - 1) & "[2;37;40m" & Mid(NLine, x + Len(r))
      End If
    Else
      NLine = Mid(NLine, 1, x - 1) & "[2;37;40m" & Mid(NLine, x + 1)
    End If
    End If
    If Len(NLine) > 0 And x > 0 Then If Len(NLine) = InStr(x, NLine, "") Then NLine = Mid(NLine, 1, Len(NLine) - 1): Exit Do
  Loop Until InStr(x + 1, NLine, "") = 0
  MakeTelnetColor = Replace(Replace(Replace(NLine, "", ""), "", ""), "", "")
  MakeTelnetColor = MakeTelnetColor & "[2;37;40m"  '[2;37m '[24;1H[0;0;0m[K[0B[23;1H[0;0;0m" & MakeTelnetColor & "[24;1H[0;0;m[K" & vbCrLf & "[1;44m" & StatusText(vsock) & "[25;1H[0;0;0m"   '& vbCrLf
End Function
Public Function MakeDCCColor(ByVal Line As String, vsock As Long) As String  ' Makes Telnet-ANSI Colors =)' : AddStack "SCMsExtensions_MakeTelnetColor(" & Line & ")"
  Dim Dark As Byte, Foreground As Byte, Background As Byte
  Dim NLine As String, NCol As String
  Dim x As Integer, x2 As Integer, X3 As Long
  NLine = Line
  x = InStr(1, NLine, "[")
  While x <> 0
    x2 = InStr(1, Mid(NLine, x + 1, 9), "m")
    If x2 <> 0 Then
      Dark = 255
      Foreground = 255
      Background = 255
      For X3 = 1 To 3
        If IsNumeric(ParamX(Mid(NLine, x + 2, (x2 + x) - (x + 2)), ";", X3)) Then
          Select Case CByte(ParamX(Mid(NLine, x + 2, (x2 + x) - (x + 2)), ";", X3))
            Case 1, 0
              Dark = 0
            Case 2
              Dark = 1
            Case 30
              Foreground = 1
            Case 31
              Foreground = 5
            Case 32
              Foreground = 3
            Case 33
              Foreground = 7
            Case 34
              Foreground = 2
            Case 35
              Foreground = 6
            Case 36
              Foreground = 10
            Case 37
              Foreground = 0
            Case 40
              Background = 1
          End Select
        End If
      Next X3
      Select Case Foreground
        Case 1
          If Dark = 1 Then NCol = "1" Else NCol = "14"
        Case 2
          If Dark = 1 Then NCol = "2" Else NCol = "12"
        Case 3
          If Dark = 1 Then NCol = "3" Else NCol = "9"
        Case 5
          If Dark = 1 Then NCol = "5" Else NCol = "4"
        Case 6
          If Dark = 1 Then NCol = "6" Else NCol = "13"
        Case 7
          If Dark = 1 Then NCol = "7" Else NCol = "8"
        Case 10
          If Dark = 1 Then NCol = "10" Else NCol = "11"
        Case 0
          If Dark = 1 Then NCol = "0" Else NCol = "15"
        Case 255
          NCol = ""
      End Select
      If Dark <> 255 Then NLine = Left(NLine, x - 1) & NCol & IIf(Background = 255, "", "," & Background) & Mid(NLine, x + x2 + 1)
    Else
      x = x + 9
      NLine = Replace(NLine, "[", "")
    End If
    x = InStr(1, NLine, "[")
  Wend
  MakeDCCColor = NLine
End Function

Function ChanStatus(Channel As String) As String
  Dim ChNum As Long, u As Long, Chan As String, u2 As Long
  Dim FUsers() As Long, FUserCount As Long, u3 As Long, ClNick As String, Host As String
  Dim NUserStat As Long, OUserStat As Long, SortedIn As Boolean, CountOps As Long
  Dim CountVoices As Long, CountUsers As Long, ChanStats As String, CountHOps As Long
  ReDim Preserve FUsers(5)
  ChNum = FindChan(Channel)
  For u = 1 To Channels(ChNum).UserCount
    Select Case Channels(ChNum).User(u).Status
      Case "@", "@%", "@%+", "@+": NUserStat = 4
      Case "%", "%+": NUserStat = 3
      Case "+": NUserStat = 2
      Case Else: NUserStat = 0
    End Select
    SortedIn = False
    For u2 = 1 To FUserCount
      Select Case Channels(ChNum).User(FUsers(u2)).Status
        Case "@", "@%", "@%+", "@+": OUserStat = 4
        Case "%", "%+": OUserStat = 2
        Case "+": OUserStat = 1
        Case Else: OUserStat = 0
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
    ChanStatus = "14[10" & Channels(ChNum).Name & "14] 3(" & Trim(Channels(ChNum).Mode) & "3) (m/" & Channels(ChNum).UserCount & " o/" & CountOps & " h/" & CountHOps & " v/" & CountVoices & " n/" & ((((Channels(ChNum).UserCount - CountVoices) - CountOps) - CountHOps)) & "3)"
  Else
    ChanStatus = "14[10" & Channels(ChNum).Name & "14] 3(" & Trim(Channels(ChNum).Mode) & "3) (m/" & Channels(ChNum).UserCount & " o/" & CountOps & " v/" & CountVoices & " n/" & ((((Channels(ChNum).UserCount - CountVoices) - CountOps) - CountHOps)) & "3)"
  End If
End Function

Function StatusText(vsock As Long) As String
  Dim sText As String, CNick As String
  sText = "[K[1;1m(" & IIf(SocketItem(vsock).RegNick = "", "", SocketItem(vsock).RegNick & "@") & BotNetNick & ") "
  Select Case GetSockFlag(vsock, SF_Status)
    Case SF_Status_UserGetName: sText = sText & " (Username)"
    Case SF_Status_UserGetPass: sText = sText & " (Password)"
    Case SF_Status_BotSetup: sText = sText & " [AWAY: BotSetup]"
    Case SF_Status_ChanSetup: sText = sText & " [AWAY: ChanSetup]"
    Case SF_Status_DCCInit
    Case SF_Status_DCCWaiting
    Case SF_Status_File: sText = sText & " [FILE: " & SocketItem(vsock).FileName & "]"
    Case SF_Status_FileArea: sText = sText & " [FA:" & SocketItem(vsock).FileAreaDir & "]"
    Case SF_Status_FileWaiting
    Case SF_Status_KISetup: sText = sText & " [AWAY: KISetup]"
    Case SF_Status_NETSetup: sText = sText & " [AWAY: NETSetup]"
    Case SF_Status_POLSetup: sText = sText & " [AWAY: POLICYSetup]"
    Case SF_Status_AUTHSetup: sText = sText & " [AWAY: AUTHSetup]"
    Case SF_Status_Party
      Select Case True
        Case MatchFlags(SocketItem(vsock).Flags, "+s"): sText = sText & " [[1;31mOWN[0m[1;1;44m (" & SocketItem(vsock).Flags & ")]"
        Case MatchFlags(SocketItem(vsock).Flags, "+n"): sText = sText & " [[1;31mOWN[0m[1;1;44m (" & SocketItem(vsock).Flags & ")]"
        Case MatchFlags(SocketItem(vsock).Flags, "+m"): sText = sText & " [[34mMAS[0m[1;1;44m (" & SocketItem(vsock).Flags & ")]"
        Case MatchFlags(SocketItem(vsock).Flags, "+t"): sText = sText & " [[1;26mNET[0m[1;1;44m (" & SocketItem(vsock).Flags & ")]"
        Case Else: sText = sText & " [[32mUSR[0m[1;1;44m (" & SocketItem(vsock).Flags & ")]"
      End Select
    Case SF_Status_PersonalSetup: sText = sText & " [AWAY: PersonalSetup]"
    Case SF_Status_SendFile
    Case SF_Status_SendFileWaiting
    Case SF_Status_SharingSetup: sText = sText & " [AWAY: SharingSetup]"
  End Select
  StatusText = sText
End Function
Sub SetAway(vsock As Long, Reason As String) ' : AddStack "BotNetRoutines_SetAway(" & vsock & ", " & Reason & ")"
Dim NoteCount As Long
  If Reason = "" Then
    NoteCount = NotesCount(SocketItem(vsock).RegNick)
    If NoteCount > 0 Then
      If NoteCount = 1 Then
        TU vsock, MakeMsg(MSG_PLNote, SocketItem(vsock).RegNick)
      Else
        TU vsock, MakeMsg(MSG_PLNotes, CStr(NoteCount), SocketItem(vsock).RegNick)
      End If
    End If
    SpreadMessage 0, SocketItem(vsock).PLChannel, MakeMsg(MSG_PLBack, SocketItem(vsock).RegNick, SocketItem(vsock).AwayMessage)
    SocketItem(vsock).LastEvent = Now
    SocketItem(vsock).AwayMessage = ""
  Else
    SpreadMessage 0, SocketItem(vsock).PLChannel, MakeMsg(MSG_PLAway, SocketItem(vsock).RegNick, Reason)
    SocketItem(vsock).LastEvent = Now
    SocketItem(vsock).AwayMessage = Reason
  End If
  ToBotNet 0, "aw " & BotNetNick & " " & SocketItem(vsock).OrderSign & " " & SocketItem(vsock).AwayMessage
End Sub

