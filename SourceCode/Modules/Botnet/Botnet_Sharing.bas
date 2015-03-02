Attribute VB_Name = "Botnet_Sharing"
Option Explicit

Sub CheckSharing(vsock As Long)
  'Only the connecting bot requests sharing
  If SocketItem(vsock).SocketDirection = SD_Out Then Exit Sub
  'Check BotFlags for sharing Flags and winsock2_send request if so.
  If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") Then TU vsock, "s rqst mix"
End Sub

'Handle a sharing message ("s ....") received from botnet
Sub SharingMessage(vsock As Long, Line As String)
  Select Case Param(Line, 2)
    Case "rqst" ' Request (mix)
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") And Param(Line, 3) = "mix" Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " requested mixing..."
        RTU vsock, "s ackr mix"
        SharingSendMixList vsock
        RTU vsock, "s acks mix"
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "ackr" ' Request accepted, prepare to receive userlist.
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") And Param(Line, 3) = "mix" Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " accepted mixing request..."
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "acks" ' Request accepted, prepare to winsock2_send userlist.
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") And Param(Line, 3) = "mix" Then
        SharingSendMixList vsock
        RTU vsock, "s ackf mix"
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "ackf" ' Request accepted, prefare to finish.
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") And Param(Line, 3) = "mix" Then
'        SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " completed..."
        RTU vsock, "s ackm mix"
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "ackm" ' Request accepted, finshed.
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") And Param(Line, 3) = "mix" Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " completed..."
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "+u" ' AddUser <name> <ALL flags>
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") Then
        SharingMixUser vsock, Param(Line, 3), GetRest(Line, 4)
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "+p" ' Password <name> <password>
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") Then
        SharingMixPassword vsock, Param(Line, 3), GetRest(Line, 4)
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "+h" ' AddHost(s) <name> <hostmask(s)>
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") Then
        SharingMixHost vsock, Param(Line, 3), GetRest(Line, 4)
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "+d" ' AddUserdata <name> <userdata>
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") Then
        SharingMixData vsock, Param(Line, 3), GetRest(Line, 4)
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "cmd" ' Execute Command <line>
      If MatchFlags(BotUsers(SocketItem(vsock).UserNum).BotFlags, "+s") Then
        SharingCommand vsock, GetRest(Line, 3)
      Else
        SpreadFlagMessage 0, "+t", "14*** Sharing: Rejecting " & SocketItem(vsock).RegNick & "'s request... Configuration missmatch."
        RTU vsock, "s abrt Configuration missmatch."
      End If
    Case "abrt"
      SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " aborted... " & GetRest(Line, 3)
  End Select
End Sub

Sub SharingSendMixList(vsock As Long)
  Dim u As Long
  For u = 1 To BotUserCount
    RTU vsock, "s +u " & BotUsers(u).Name & " " & CombineAllFlags(u)
    RTU vsock, "s +p " & BotUsers(u).Name & " " & BotUsers(u).Password
    If BotUsers(u).HostMaskCount > 0 Then RTU vsock, "s +h " & BotUsers(u).Name & " " & CombineAllHosts(u)
    If BotUsers(u).UserData <> "" Then RTU vsock, "s +d " & BotUsers(u).Name & " " & BotUsers(u).UserData
  Next u
End Sub

Sub SharingMixUser(vsock As Long, NewUser As String, NewFlags As String)
  Dim u As Long, u2 As Long, u3 As Long
  Dim FoundHim As Boolean, FoundIt As Boolean
  Dim FlagValue As String
  Dim FlagName As String
  Dim FlagGroup As String
  Dim ChangedSomething As Boolean, ChangeLine As String
  
  ChangedSomething = False
  FoundHim = False
  For u = 1 To BotUserCount
    If LCase(NewUser) = LCase(BotUsers(u).Name) Then
      FoundHim = True
      Exit For
    End If
  Next u
  If FoundHim = False Then
    BotUserCount = BotUserCount + 1
    If BotUserCount > UBound(BotUsers) Then ReDim Preserve BotUsers(BotUserCount + 4)
    u = BotUserCount
    BotUsers(u).Name = NewUser
    SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " added user " & NewUser & "."
    FoundHim = True
  End If
  If FoundHim = True Then
    If InStr(NewFlags, ",") = 0 Then
      ' Keine ChannelFlags
      FlagValue = Replace(NewFlags, "s", "")
      If Not Replace(BotUsers(u).Flags, "s", "") = FlagValue Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " changed " & NewUser & "'s userflags to " & CombineFlags(BotUsers(u).Flags, FlagValue)
        If InStr(BotUsers(u).Flags, "s") > 0 Then BotUsers(u).Flags = CombineFlags("ps", "+" & FlagValue) Else BotUsers(u).Flags = FlagValue
        ChangeLine = FlagValue
        ChangedSomething = True
      End If
    Else
      For u2 = 1 To ParamXCount(NewFlags, ",")
        FlagGroup = ParamX(NewFlags, ",", u2)
        If InStr(FlagGroup, " ") > 0 Then
          FlagValue = Param(FlagGroup, 1)
          FlagName = Param(FlagGroup, 2)
          If FlagName = "bot" Then
            FlagValue = Replace(Replace(FlagValue, "a", ""), "h", "")
            'BotFlags
            If Not MatchFlags(FlagValue, Replace(Replace(BotUsers(u).BotFlags, "a", ""), "h", "")) Then
              SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " changed " & NewUser & "'s botflags to " & CombineFlags(BotUsers(u).BotFlags, FlagValue)
              BotUsers(u).BotFlags = FlagValue
              ChangeLine = ChangeLine & FlagValue & " " & FlagName & ","
              ChangedSomething = True
            End If
          Else
            'ChannelFlags
            For u3 = 1 To BotUsers(u).ChannelFlagCount
              If LCase(BotUsers(u).ChannelFlags(u3).Channel) = LCase(FlagName) Then
                FoundIt = True
                If Not MatchFlags(FlagValue, BotUsers(u).ChannelFlags(u3).Flags) Then
                  SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " changed " & NewUser & "'s channelflags in " & FlagName & " to " & CombineFlags(BotUsers(u).ChannelFlags(u3).Flags, FlagValue)
                  BotUsers(u).ChannelFlags(u3).Flags = FlagValue
                  ChangeLine = ChangeLine & FlagValue & " " & FlagName & ","
                  ChangedSomething = True
                End If
                Exit For
              End If
            Next u3
            If FoundIt = False Then
              ' Nisch gefindet... Tu dranhängen
              BotUsers(u).ChannelFlagCount = BotUsers(u).ChannelFlagCount + 1
'              If BotUsers(u).ChannelFlagCount > UBound(BotUsers(u).ChannelFlags()) Then ReDim Preserve BotUsers(u).ChannelFlags(1 To UBound(BotUsers(u).ChannelFlags) + 5)
              SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " added " & NewUser & "'s channelflags in " & FlagName & " with " & CombineFlags(BotUsers(u).ChannelFlags(u3).Flags, FlagValue)
              BotUsers(u).ChannelFlags(BotUsers(u).ChannelFlagCount).Channel = FlagName
              BotUsers(u).ChannelFlags(BotUsers(u).ChannelFlagCount).Flags = FlagValue
              ChangeLine = ChangeLine & FlagValue & " " & FlagName & ","
              ChangedSomething = True
            End If
          End If
        Else
          ' Global
          FlagValue = Replace(FlagGroup, "s", "")
          If Not MatchFlags(Replace(BotUsers(u).Flags, "s", ""), FlagValue) Then
            SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " changed " & NewUser & "'s userflags to " & CombineFlags(BotUsers(u).Flags, FlagValue)
            If InStr(BotUsers(u).Flags, "s") > 0 Then BotUsers(u).Flags = CombineFlags(FlagValue, "+sp") Else BotUsers(u).Flags = FlagValue
            ChangeLine = ChangeLine & FlagValue & ","
            ChangedSomething = True
          End If
        End If
      Next u2
    End If
  End If
  If Right(ChangeLine, 1) = "," Then ChangeLine = Left(ChangeLine, Len(ChangeLine) - 1)
  ChangeLine = Trim(ChangeLine)
  If ChangedSomething = False Or ChangeLine = "" Then Exit Sub
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If MatchFlags(BotUsers(SocketItem(u).UserNum).BotFlags, "+s") And u <> vsock Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: Forwarding request to " & SocketItem(u).RegNick
        RTU u, "s +u " & NewUser & " " & NewFlags
      End If
    End If
  Next u
End Sub

Sub SharingMixPassword(vsock As Long, NewUser As String, NewPassword As String)
  Dim u As Long
  Dim FoundHim As Boolean
  Dim ChangedSomething As Boolean
  FoundHim = False
  ChangedSomething = False
  For u = 1 To BotUserCount
    If LCase(NewUser) = LCase(BotUsers(u).Name) Then
      FoundHim = True
      Exit For
    End If
  Next u
  If FoundHim = True Then
    If BotUsers(u).Password <> NewPassword Then
      ChangedSomething = True
      SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " changed user " & NewUser & "'s password." ' to " & NewPassword
      BotUsers(u).Password = NewPassword
    Else
      Exit Sub
    End If
  Else
    Exit Sub
  End If
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If MatchFlags(BotUsers(SocketItem(u).UserNum).BotFlags, "+s") And u <> vsock Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: Forwarded request to " & SocketItem(u).RegNick
        RTU u, "s +p " & NewUser & " " & NewPassword
      End If
    End If
  Next u
End Sub

Sub SharingMixHost(vsock As Long, NewUser As String, NewHosts As String)
  Dim u As Long, u2 As Long
  Dim FoundHim As Boolean
  Dim ChangedSomething As Boolean
  Dim ChangeLine As String
  ChangedSomething = False
  FoundHim = False
  For u = 1 To BotUserCount
    If LCase(NewUser) = LCase(BotUsers(u).Name) Then
      FoundHim = True
      Exit For
    End If
  Next u
  If FoundHim = True Then
    ChangeLine = NewHosts
    For u2 = 1 To BotUsers(u).HostMaskCount
      If InStr(LCase(NewHosts), LCase(BotUsers(u).HostMasks(u2))) <> 0 Then
        ChangeLine = Replace(LCase(ChangeLine), LCase(BotUsers(u).HostMasks(u2)), "")
      End If
    Next u2
    ChangeLine = Trim(ChangeLine)
    If ParamCount(ChangeLine) > 0 Then
      For u2 = 1 To ParamCount(ChangeLine)
        BotUsers(u).HostMaskCount = BotUsers(u).HostMaskCount + 1
        BotUsers(u).HostMasks(BotUsers(u).HostMaskCount) = Param(ChangeLine, u2)
        SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " added hostmask " & Param(ChangeLine, u2) & " to user " & NewUser & "."
        UpdateRegUsers "A " & BotUsers(u).Name & " " & BotUsers(u).HostMasks(BotUsers(u).HostMaskCount)
      Next u2
    Else
      Exit Sub
    End If
  Else
    Exit Sub
  End If
  If ChangeLine = "" Then Exit Sub
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If MatchFlags(BotUsers(SocketItem(u).UserNum).BotFlags, "+s") And u <> vsock Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: Forwarded request to " & SocketItem(u).RegNick
        RTU u, "s +p " & NewUser & " " & ChangeLine
      End If
    End If
  Next u
End Sub

Sub SharingMixData(vsock As Long, NewUser As String, NewData As String)
  Dim u As Long, u2 As Long
  Dim FoundHim As Boolean
  Dim ChangedSomething As Boolean
  Dim OldIndex As String, OldData As String
  Dim CurPiece As String, CurEntry As String, CurLen As Long, CurPos As Long
  
  FoundHim = False
  For u = 1 To BotUserCount
    If LCase(NewUser) = LCase(BotUsers(u).Name) Then
      FoundHim = True
      Exit For
    End If
  Next u
  ChangedSomething = False
  If FoundHim = True Then
    OldIndex = Param(NewData, 1)
    OldData = GetRest(NewData, 2)
    CurPos = 1
    For u2 = 1 To ParamXCount(OldIndex, ",")
      CurPiece = ParamX(OldIndex, ",", u2)
      CurEntry = LCase(ParamX(CurPiece, ":", 1))
      CurLen = CLng(ParamX(CurPiece, ":", 2))
      If GetUserData(u, CurEntry, "") <> Mid(OldData, CurPos, CurLen) Then
        SetUserData u, CurEntry, Mid(OldData, CurPos, CurLen)
        ChangedSomething = True
      End If
      CurPos = CurPos + CurLen
    Next u2
  End If
  If ChangedSomething = False Then Exit Sub
  SpreadFlagMessage 0, "+t", "14*** Sharing: " & SocketItem(vsock).RegNick & " mixed user " & NewUser & "'s data." ' with " & NewData
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If MatchFlags(BotUsers(SocketItem(u).UserNum).BotFlags, "+s") And u <> vsock Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: Forwarded request to " & SocketItem(u).RegNick
        RTU u, "s +p " & NewUser & " " & NewData
      End If
    End If
  Next u
End Sub

Sub SharingCommand(vsock As Long, Line As String)
  Dim NewSock As Long
  If ParamCount(Line) > 1 Then
    NewSock = AddSocket
    SocketItem(NewSock).RegNick = SocketItem(vsock).RegNick
    SocketItem(NewSock).Flags = "fijmnoprstvwx"
    SocketItem(NewSock).IRCNick = "²sharing²"
    SocketItem(NewSock).CurrentQuestion = ""
    SocketItem(NewSock).SetupChan = SocketItem(vsock).RegNick
    SocketItem(NewSock).SocketNumber = 0
    SocketItem(NewSock).UserNum = SocketItem(vsock).UserNum
    SocketItem(NewSock).OnBot = SocketItem(vsock).OnBot
    SetSockFlag NewSock, SF_Status, SF_Status_Party
    Party NewSock, 0, Line
    SocketItem(NewSock).IsInternalSocket = False
    RemoveSocket NewSock, 0, "", True
'    For NewSock = 1 To SocketCount
'      If IsValidSocket(NewSock) Then
'        If MatchFlags(BotUsers(SocketItem(NewSock).UserNum).BotFlags, "+s") And NewSock <> vsock Then
'          SpreadFlagMessage 0, "+t", "14*** Sharing: Forwarded request to " & SocketItem(NewSock).RegNick
'          RTU NewSock, "s cmd " & Line
'        End If
'      End If
'    Next NewSock
  End If
End Sub

Sub SharingSpreadMessage(WhoDid As String, Line As String)
  Dim u As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If MatchFlags(BotUsers(SocketItem(u).UserNum).BotFlags, "+s") And SocketItem(u).RegNick <> WhoDid And SocketItem(u).IRCNick <> "²sharing²" Then
        SpreadFlagMessage 0, "+t", "14*** Sharing: Forwarded changes to " & SocketItem(u).RegNick
        RTU u, "s " & Line
      End If
    End If
  Next u
End Sub

