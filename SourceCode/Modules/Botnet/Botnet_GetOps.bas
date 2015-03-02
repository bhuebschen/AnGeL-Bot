Attribute VB_Name = "Botnet_GetOps"
',-======================- ==-- -  -
'|   AnGeL - Botnet - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


'Request ops from the other bots in a channel
Public Sub RequestOps(ByVal ChNum As Long)
Dim Requestable() As String, ReqNum As Long, RequestedFrom As String, u As Long, u2 As Long, u3 As Long
Dim AlreadyIn As Boolean, Bot1 As String, Bot2 As String, Bot3 As String
Dim BPW As String, TempStr As String
  'Check whether it's possible and necessary to request ops
  If (ChNum = 0) Or (ChNum > ChanCount) Then Exit Sub
  If (Channels(ChNum).GotOPs = True) Or (Channels(ChNum).CompletedWHO = False) Then Exit Sub
  If IsOrdered("gop " & Channels(ChNum).Name) Then Exit Sub
  
  ReDim Preserve Requestable(5)
  For u2 = 1 To Channels(ChNum).UserCount
    If (InStr(Channels(ChNum).User(u2).Status, "@") > 0) And (Channels(ChNum).User(u2).RegNick <> "") Then
      If InStr("-+=", Left(Mask(Channels(ChNum).User(u2).Hostmask, 12), 1)) = 0 Then
        For u3 = 2 To BotCount
          If u3 > UBound(Bots()) Then Exit For
          If LCase(Channels(ChNum).User(u2).RegNick) = LCase(Bots(u3).Nick) Then
            If MatchFlags(GetUserChanFlags(Bots(u3).Nick, Channels(ChNum).Name), "+o") Then
              AlreadyIn = False
              For u = 1 To ReqNum
                If Requestable(u) = Bots(u3).Nick Then AlreadyIn = True: Exit For
              Next u
              If Not AlreadyIn Then
                ReqNum = ReqNum + 1: If ReqNum > UBound(Requestable()) Then ReDim Preserve Requestable(UBound(Requestable()) + 5)
                Requestable(ReqNum) = Bots(u3).Nick
              End If
            End If
            Exit For
          End If
        Next u3
      End If
    End If
  Next u2
  
  'Choose 3 random bots
  If ReqNum > 3 Then
    Bot1 = Requestable(Int(Rnd * ReqNum) + 1)
    Do: Bot2 = Requestable(Int(Rnd * ReqNum) + 1): Loop Until Bot2 <> Bot1
    Do: Bot3 = Requestable(Int(Rnd * ReqNum) + 1): Loop Until Bot3 <> Bot1 And Bot3 <> Bot2
  Else
    If ReqNum >= 1 Then Bot1 = Requestable(1)
    If ReqNum >= 2 Then Bot2 = Requestable(2)
    If ReqNum = 3 Then Bot3 = Requestable(3)
  End If
  If Bot1 <> "" Then SendToBot Bot1, "z " & BotNetNick & " " & Bot1 & " gop op " & Channels(ChNum).Name & " " & MyNick: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & Bot1 Else RequestedFrom = Bot1
  If Bot2 <> "" Then SendToBot Bot2, "z " & BotNetNick & " " & Bot2 & " gop op " & Channels(ChNum).Name & " " & MyNick: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & Bot2 Else RequestedFrom = Bot2
  If Bot3 <> "" Then SendToBot Bot3, "z " & BotNetNick & " " & Bot3 & " gop op " & Channels(ChNum).Name & " " & MyNick: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & Bot3 Else RequestedFrom = Bot3

  If RequestedFrom <> "" Then
    SpreadFlagMessage 0, "+t", "14[" & Time & "] *** GetOps: Requested op for " & Channels(ChNum).Name & " from " & RequestedFrom & "."
    Order "gop " & Channels(ChNum).Name, 20
  End If
  If ReqNum < 3 Then
    'No bots or only few bots in botnet found? -> Try to request op via MSG!
    ReqNum = 0: Bot1 = "": Bot2 = "": Bot3 = "": RequestedFrom = ""
    For u2 = 1 To Channels(ChNum).UserCount
      If (InStr(Channels(ChNum).User(u2).Status, "@") > 0) And (Channels(ChNum).User(u2).RegNick <> "") Then
        If InStr("-+=", Left(Mask(Channels(ChNum).User(u2).Hostmask, 12), 1)) = 0 Then
          BPW = BotUsers(Channels(ChNum).User(u2).UserNum).Password
          If BPW <> "" Then
            TempStr = GetUserChanFlags(Channels(ChNum).User(u2).RegNick, Channels(ChNum).Name)
            If MatchFlags(TempStr, "+bo") Then
              'Don't request op via MSG when a bot is in botnet :)
              If GetBotPos(Channels(ChNum).User(u2).RegNick) = 0 Then
                AlreadyIn = False
                For u = 1 To ReqNum
                  If Requestable(u) = Channels(ChNum).User(u2).Hostmask & " " & BPW Then AlreadyIn = True: Exit For
                Next u
                If Not AlreadyIn Then
                  ReqNum = ReqNum + 1: If ReqNum > UBound(Requestable()) Then ReDim Preserve Requestable(UBound(Requestable()) + 5)
                  Requestable(ReqNum) = Channels(ChNum).User(u2).Hostmask & " " & BPW
                End If
              End If
            End If
          End If
        End If
      End If
    Next u2
    
    'Choose 3 random bots
    If ReqNum > 3 Then
      Bot1 = Requestable(Int(Rnd * ReqNum) + 1)
      Do: Bot2 = Requestable(Int(Rnd * ReqNum) + 1): Loop Until Bot2 <> Bot1
      Do: Bot3 = Requestable(Int(Rnd * ReqNum) + 1): Loop Until Bot3 <> Bot1 And Bot3 <> Bot2
    Else
      If ReqNum >= 1 Then Bot1 = Requestable(1)
      If ReqNum >= 2 Then Bot2 = Requestable(2)
      If ReqNum = 3 Then Bot3 = Requestable(3)
    End If
    If ServerUseFullAdress = False Then
      If Bot1 <> "" Then SendLine "PRIVMSG " & GetNick(Param(Bot1, 1)) & " :OP " & Param(Bot1, 2) & " " & Channels(ChNum).Name, 1: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & GetNick(Param(Bot1, 1)) Else RequestedFrom = GetNick(Param(Bot1, 1))
      If Bot2 <> "" Then SendLine "PRIVMSG " & GetNick(Param(Bot2, 1)) & " :OP " & Param(Bot2, 2) & " " & Channels(ChNum).Name, 1: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & GetNick(Param(Bot2, 1)) Else RequestedFrom = GetNick(Param(Bot2, 1))
      If Bot3 <> "" Then SendLine "PRIVMSG " & GetNick(Param(Bot3, 1)) & " :OP " & Param(Bot3, 2) & " " & Channels(ChNum).Name, 1: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & GetNick(Param(Bot3, 1)) Else RequestedFrom = GetNick(Param(Bot3, 1))
    Else
      If Bot1 <> "" Then SendLine "PRIVMSG " & Param(Bot1, 1) & " :OP " & Param(Bot1, 2) & " " & Channels(ChNum).Name, 1: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & GetNick(Param(Bot1, 1)) Else RequestedFrom = GetNick(Param(Bot1, 1))
      If Bot2 <> "" Then SendLine "PRIVMSG " & Param(Bot2, 1) & " :OP " & Param(Bot2, 2) & " " & Channels(ChNum).Name, 1: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & GetNick(Param(Bot2, 1)) Else RequestedFrom = GetNick(Param(Bot2, 1))
      If Bot3 <> "" Then SendLine "PRIVMSG " & Param(Bot3, 1) & " :OP " & Param(Bot3, 2) & " " & Channels(ChNum).Name, 1: If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & GetNick(Param(Bot3, 1)) Else RequestedFrom = GetNick(Param(Bot3, 1))
    End If
    If RequestedFrom <> "" Then
      SpreadFlagMessage 0, "+t", "14[" & Time & "] *** GetOps: Requested MSG op for " & Channels(ChNum).Name & " from " & RequestedFrom & "."
      Order "gop " & Channels(ChNum).Name, 20
    End If
  End If
End Sub


'Offers ops to bots with +o on a channel
Public Sub OfferOps(ByVal ChNum As Long)
Dim Requestable() As String, RBotNick() As String, ReqNum As Long, RequestedFrom As String, u As Long, u2 As Long, u3 As Long
Dim AlreadyIn As Boolean, Bot1 As String, Bot2 As String, Bot3 As String
  'Check whether it's possible to offer ops
  If (ChNum = 0) Or (ChNum > ChanCount) Then Exit Sub
  If (Channels(ChNum).GotOPs = False) Or (Channels(ChNum).CompletedWHO = False) Then Exit Sub
  If IsOrdered("xgop " & Channels(ChNum).Name) Then Exit Sub
  ReDim Preserve Requestable(5)
  ReDim Preserve RBotNick(5)
  
  For u2 = 1 To Channels(ChNum).UserCount
    If (InStr(Channels(ChNum).User(u2).Status, "@") = 0) And (Channels(ChNum).User(u2).RegNick <> "") Then
      If InStr("-+=", Left(Mask(Channels(ChNum).User(u2).Hostmask, 12), 1)) = 0 Then
        For u3 = 2 To BotCount
          If u3 > UBound(Bots()) Then Exit For
          If LCase(Channels(ChNum).User(u2).RegNick) = LCase(Bots(u3).Nick) Then
            If MatchFlags(GetUserChanFlags(Bots(u3).Nick, Channels(ChNum).Name), "+bo") Then
              AlreadyIn = False
              For u = 1 To ReqNum
                If Requestable(u) = Bots(u3).Nick Then AlreadyIn = True: Exit For
              Next u
              If (AlreadyIn = False) And (Bots(u3).SendRequests = True) Then
                ReqNum = ReqNum + 1
                If ReqNum > UBound(Requestable()) Then ReDim Preserve Requestable(UBound(Requestable()) + 5)
                If ReqNum > UBound(RBotNick()) Then ReDim Preserve RBotNick(UBound(RBotNick()) + 5)
                Requestable(ReqNum) = Bots(u3).Nick
                RBotNick(ReqNum) = Channels(ChNum).User(u2).Nick
              End If
            End If
            Exit For
          End If
        Next u3
      End If
    End If
  Next u2
  
  u3 = 0
  For u = 1 To ReqNum
    If IsOrdered("xgop2 " & Channels(ChNum).Name & " " & Requestable(u)) = False Then
      u3 = u3 + 1: If u3 > 3 Then Exit For
      SendToBot Requestable(u), "z " & BotNetNick & " " & Requestable(u) & " xgop wantops? " & Channels(ChNum).Name & " " & MyNick & " " & RBotNick(u): If RequestedFrom <> "" Then RequestedFrom = RequestedFrom & ", " & Requestable(u) Else RequestedFrom = Requestable(u)
      Bots(GetBotPos(Requestable(u))).SendRequests = False
      Order "xgop2 " & Channels(ChNum).Name & " " & Requestable(u), 10
    End If
  Next u

  If RequestedFrom <> "" Then
    SpreadFlagMessage 0, "+t", "14[" & Time & "] *** GetOps: Offered op for " & Channels(ChNum).Name & " to " & RequestedFrom + IIf(u3 > 3, "...", ".")
    If u3 > 3 Then
      TimedEvent "xgop " & Channels(ChNum).Name, 1
    Else
      Order "xgop " & Channels(ChNum).Name, 20
    End If
  End If
End Sub


Public Sub SendGetOps(Action As String, Chan As String, LastParam As String)
Dim u As Long, SNum As Long
  For u = 1 To BotCount
    If MatchFlags(GetUserChanFlags(Bots(u).Nick, Chan), "+bo") Then
      SendToBot Bots(u).Nick, "z " & BotNetNick & " " & Bots(u).Nick & " gop " & Action & " " & Chan & " " & LastParam
    End If
  Next u
End Sub

