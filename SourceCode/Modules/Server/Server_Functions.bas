Attribute VB_Name = "Server_Functions"
',-======================- ==-- -  -
'|   AnGeL - Server - Functions
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit




Public Function Connected() As Boolean
  Connected = (MyNick <> "")
End Function


Public Function BotChannels() As String
  Dim Rest As String, u As Long
  Rest = ""
  For u = 1 To ChanCount
    If (Left(Channels(u).Name, 1) <> "&") And (Not Channels(u).Secret) Then
      If Rest <> "" Then Rest = Rest & ", " & Channels(u).Name Else Rest = Channels(u).Name
    End If
  Next u
  If Rest = "" Then Rest = "no channels"
  BotChannels = Rest
End Function



Public Function FunnyKick(Nick As String) As String
  Randomize WinTickCount
  Select Case Int(Rnd * 24) + 1
    Case 1
      FunnyKick = "Did you know the difference between INSIDE and OUTSIDE? Now you do."
    Case 2
      FunnyKick = "...and don't come back."
    Case 3
      FunnyKick = "¯`·.¸ whack! ¸.·´¯"
    Case 4
      FunnyKick = "¯`·.¸ BAM! ¸.·´¯"
    Case 5
      FunnyKick = "ERROR: Intelligence missing."
    Case 6
      FunnyKick = "Connection reset by ME"
    Case 7
      FunnyKick = "C U... but not here."
    Case 8
      FunnyKick = "Press any key to continue or any other key to quit..."
    Case 9
      FunnyKick = "@@@ POOF! @@@"
    Case 10
      FunnyKick = "Do you like being kicked? I hope so."
    Case 11
      FunnyKick = "Mess with the best, die like the rest."
    Case 12
      FunnyKick = "...at least my foot likes you."
    Case 13
      FunnyKick = "Oops, missed."
    Case 14
      FunnyKick = "See you later. MUCH later I hope."
    Case 15
      FunnyKick = "A kick every day keeps the " & Nick & " away!"
    Case 16
      FunnyKick = "Sayonara Scumbag!"
    Case 17
      FunnyKick = "Get outta here!"
    Case 18
      FunnyKick = "'Move your ass!' ;)"
    Case 19
      FunnyKick = "To be here or not to be here..."
    Case 20
      FunnyKick = "No comment!"
    Case 21
      FunnyKick = "My foot seems to like it."
    Case 22
      FunnyKick = "*yawn* Whoops, sorry ;)"
    Case 23
      FunnyKick = "That was fun, let's do it again!!!"
    Case 24
      FunnyKick = "3, 2, 1, ...lift off!"
  End Select
End Function

Public Function IsBanned(ChNum As Long, Hostmask As String) As Boolean
  Dim u As Long
  IsBanned = False
  If ChNum > 0 Then
    For u = 1 To Channels(ChNum).BanCount
      If LCase(Channels(ChNum).BanList(u).Mask) = LCase(Hostmask) Then
        IsBanned = True
        Exit Function
      End If
    Next u
  End If
End Function

Public Function IsInvited(ChNum As Long, Hostmask As String) As Boolean
  Dim u As Long
  IsInvited = False
  If ChNum > 0 Then
    For u = 1 To Channels(ChNum).InviteCount
      If LCase(Channels(ChNum).InviteList(u).Mask) = LCase(Hostmask) Then
        IsInvited = True
        Exit Function
      End If
    Next u
  End If
End Function

Public Function IsExceptd(ChNum As Long, Hostmask As String) As Boolean ' : AddStack "Base_IsBanned(" & ChNum & ", " & Hostmask & ")"
  Dim u As Long
  IsExceptd = False
  If ChNum > 0 Then
    For u = 1 To Channels(ChNum).ExceptCount
      If LCase(Channels(ChNum).ExceptList(u).Mask) = LCase(Hostmask) Then
        IsExceptd = True
        Exit Function
      End If
    Next u
  End If
End Function

Public Function IsFloodNick(Nick As String) As Boolean
  Dim lres As Boolean
  lres = (LCase(Nick) Like "[a-z{}]#[a-z{}]#[a-z{}]#[a-z{}]#[a-z{}]")
  If lres Then IsFloodNick = True: Exit Function
  lres = (Nick Like "[[\-^}{_`][[\-^}{_`][[\-^}{_`][[\-^}{_`][[\-^}{_`][[\-^}{_`][[\-^}{_`][[\-^}{_`][[\-^}{_`]")
  IsFloodNick = lres
End Function

Public Sub AddMassMode(Action As String, Target As String)
  Select Case LCase(Action)
    Case "+o": If MassModeOps = "" Then MassModeOps = Target Else If InStr(" " & MassModeOps & " ", " " & Target & " ") = 0 Then MassModeOps = MassModeOps & " " & Target
    Case "-o": If MassModeDeops = "" Then MassModeDeops = Target Else If InStr(" " & MassModeDeops & " ", " " & Target & " ") = 0 Then MassModeDeops = MassModeDeops & " " & Target
    Case "+b": If MassModeBans = "" Then MassModeBans = Target Else If InStr(" " & MassModeBans & " ", " " & Target & " ") = 0 Then MassModeBans = MassModeBans & " " & Target
    Case "-b": If MassModeUnbans = "" Then MassModeUnbans = Target Else If InStr(" " & MassModeUnbans & " ", " " & Target & " ") = 0 Then MassModeUnbans = MassModeUnbans & " " & Target
    Case "+I": If MassModeInvite = "" Then MassModeInvite = Target Else If InStr(" " & MassModeInvite & " ", " " & Target & " ") = 0 Then MassModeInvite = MassModeInvite & " " & Target
    Case "-I": If MassModeUnInvite = "" Then MassModeUnInvite = Target Else If InStr(" " & MassModeUnInvite & " ", " " & Target & " ") = 0 Then MassModeUnInvite = MassModeUnInvite & " " & Target
    Case "+e": If MassModeExcept = "" Then MassModeExcept = Target Else If InStr(" " & MassModeExcept & " ", " " & Target & " ") = 0 Then MassModeExcept = MassModeExcept & " " & Target
    Case "-e": If MassModeUnExcept = "" Then MassModeUnExcept = Target Else If InStr(" " & MassModeExcept & " ", " " & Target & " ") = 0 Then MassModeExcept = MassModeUnExcept & " " & Target
    Case "+v": If MassModeVoice = "" Then MassModeVoice = Target Else If InStr(" " & MassModeVoice & " ", " " & Target & " ") = 0 Then MassModeVoice = MassModeVoice & " " & Target
    Case "-v": If MassModeDevoice = "" Then MassModeDevoice = Target Else If InStr(" " & MassModeDevoice & " ", " " & Target & " ") = 0 Then MassModeDevoice = MassModeDevoice & " " & Target
    Case "+h": If MassModeHOps = "" Then MassModeHOps = Target Else If InStr(" " & MassModeHOps & " ", " " & Target & " ") = 0 Then MassModeHOps = MassModeHOps & " " & Target
    Case "-h": If MassModeDehops = "" Then MassModeDehops = Target Else If InStr(" " & MassModeDehops & " ", " " & Target & " ") = 0 Then MassModeDehops = MassModeDehops & " " & Target
  End Select
End Sub

Public Sub AddMassMode2(Action As String, Target As String)
  Select Case LCase(Action)
    Case "+o": If MassModeOps2 = "" Then MassModeOps2 = Target Else If InStr(" " & MassModeOps2 & " ", " " & Target & " ") = 0 Then MassModeOps2 = MassModeOps2 & " " & Target
    Case "-o": If MassModeDeops2 = "" Then MassModeDeops2 = Target Else If InStr(" " & MassModeDeops2 & " ", " " & Target & " ") = 0 Then MassModeDeops2 = MassModeDeops2 & " " & Target
    Case "+b": If MassModeBans2 = "" Then MassModeBans2 = Target Else If InStr(" " & MassModeBans2 & " ", " " & Target & " ") = 0 Then MassModeBans2 = MassModeBans2 & " " & Target
    Case "-b": If MassModeUnbans2 = "" Then MassModeUnbans2 = Target Else If InStr(" " & MassModeUnbans2 & " ", " " & Target & " ") = 0 Then MassModeUnbans2 = MassModeUnbans2 & " " & Target
    Case "+I": If MassModeInvite2 = "" Then MassModeInvite2 = Target Else If InStr(" " & MassModeInvite2 & " ", " " & Target & " ") = 0 Then MassModeInvite2 = MassModeInvite2 & " " & Target
    Case "-I": If MassModeUnInvite2 = "" Then MassModeUnInvite2 = Target Else If InStr(" " & MassModeUnInvite2 & " ", " " & Target & " ") = 0 Then MassModeUnInvite2 = MassModeUnInvite2 & " " & Target
    Case "+e": If MassModeExcept2 = "" Then MassModeExcept2 = Target Else If InStr(" " & MassModeExcept2 & " ", " " & Target & " ") = 0 Then MassModeExcept2 = MassModeExcept2 & " " & Target
    Case "-e": If MassModeUnExcept2 = "" Then MassModeUnExcept2 = Target Else If InStr(" " & MassModeUnExcept2 & " ", " " & Target & " ") = 0 Then MassModeUnExcept2 = MassModeUnExcept2 & " " & Target
    Case "+v": If MassModeVoice2 = "" Then MassModeVoice2 = Target Else If InStr(" " & MassModeVoice2 & " ", " " & Target & " ") = 0 Then MassModeVoice2 = MassModeVoice2 & " " & Target
    Case "-v": If MassModeDevoice2 = "" Then MassModeDevoice2 = Target Else If InStr(" " & MassModeDevoice2 & " ", " " & Target & " ") = 0 Then MassModeDevoice2 = MassModeDevoice2 & " " & Target
    Case "+h": If MassModeHOps2 = "" Then MassModeHOps2 = Target Else If InStr(" " & MassModeHOps2 & " ", " " & Target & " ") = 0 Then MassModeHOps2 = MassModeHOps2 & " " & Target
    Case "-h": If MassModeDehops2 = "" Then MassModeDehops2 = Target Else If InStr(" " & MassModeDehops2 & " ", " " & Target & " ") = 0 Then MassModeDehops2 = MassModeDehops2 & " " & Target
  End Select
End Sub

Public Sub DoMassMode(Channel As String)
  Dim Action As String, Multiple As String, MultCount As Long, LastOne As Long
  Dim u2 As Long, TargetNick As String, Result As String
  If MassModeDeops <> "" Then
    For u2 = 1 To ParamCount(MassModeDeops)
      TargetNick = Param(MassModeDeops, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "o" Else Action = Action & "-o"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeOps <> "" Then
    For u2 = 1 To ParamCount(MassModeOps)
      TargetNick = Param(MassModeOps, u2)
      If IsOrdered("giveop " & Channel & " " & TargetNick) = False Then
        Order "giveop " & Channel & " " & TargetNick, 10
        MultCount = MultCount + 1
        If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
        If LastOne = 2 Then Action = Action & "o" Else Action = Action & "+o"
        LastOne = 2
        If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
      End If
    Next u2
  End If
  If MassModeHOps <> "" Then
    For u2 = 1 To ParamCount(MassModeHOps)
      TargetNick = Param(MassModeHOps, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "h" Else Action = Action & "+h"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeDehops <> "" Then
    For u2 = 1 To ParamCount(MassModeDehops)
      TargetNick = Param(MassModeDehops, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "h" Else Action = Action & "-h"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeVoice <> "" Then
    For u2 = 1 To ParamCount(MassModeVoice)
      TargetNick = Param(MassModeVoice, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "v" Else Action = Action & "+v"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeDevoice <> "" Then
    For u2 = 1 To ParamCount(MassModeDevoice)
      TargetNick = Param(MassModeDevoice, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "v" Else Action = Action & "-v"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeBans <> "" Then
    For u2 = 1 To ParamCount(MassModeBans)
      TargetNick = Param(MassModeBans, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "b" Else Action = Action & "+b"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeUnbans <> "" Then
    For u2 = 1 To ParamCount(MassModeUnbans)
      TargetNick = Param(MassModeUnbans, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "b" Else Action = Action & "-b"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeInvite <> "" Then
    For u2 = 1 To ParamCount(MassModeInvite)
      TargetNick = Param(MassModeInvite, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "I" Else Action = Action & "+I"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeUnInvite <> "" Then
    For u2 = 1 To ParamCount(MassModeUnInvite)
      TargetNick = Param(MassModeUnInvite, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "I" Else Action = Action & "-I"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeExcept <> "" Then
    For u2 = 1 To ParamCount(MassModeExcept)
      TargetNick = Param(MassModeExcept, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "e" Else Action = Action & "+e"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeUnExcept <> "" Then
    For u2 = 1 To ParamCount(MassModeUnExcept)
      TargetNick = Param(MassModeUnExcept, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "e" Else Action = Action & "-e"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MultCount > 0 Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = ""
  MassModeOps = "": MassModeDeops = "": MassModeBans = "": MassModeUnbans = ""
End Sub

Public Sub DoMassMode2(Channel As String)
  Dim Action As String, Multiple As String, MultCount As Long, LastOne As Long
  Dim u2 As Long, TargetNick As String, Result As String
  If MassModeDeops2 <> "" Then
    For u2 = 1 To ParamCount(MassModeDeops2)
      TargetNick = Param(MassModeDeops2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "o" Else Action = Action & "-o"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeOps2 <> "" Then
    For u2 = 1 To ParamCount(MassModeOps2)
      TargetNick = Param(MassModeOps2, u2)
      If IsOrdered("giveop " & Channel & " " & TargetNick) = False Then
        Order "giveop " & Channel & " " & TargetNick, 10
        MultCount = MultCount + 1
        If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
        If LastOne = 2 Then Action = Action & "o" Else Action = Action & "+o"
        LastOne = 2
        If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
      End If
    Next u2
  End If
  If MassModeHOps2 <> "" Then
    For u2 = 1 To ParamCount(MassModeHOps2)
      TargetNick = Param(MassModeHOps2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "h" Else Action = Action & "+h"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeDehops2 <> "" Then
    For u2 = 1 To ParamCount(MassModeDehops2)
      TargetNick = Param(MassModeDehops2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "h" Else Action = Action & "-h"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeVoice2 <> "" Then
    For u2 = 1 To ParamCount(MassModeVoice2)
      TargetNick = Param(MassModeVoice2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "v" Else Action = Action & "+v"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeDevoice2 <> "" Then
    For u2 = 1 To ParamCount(MassModeDevoice2)
      TargetNick = Param(MassModeDevoice2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "v" Else Action = Action & "-v"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeBans2 <> "" Then
    For u2 = 1 To ParamCount(MassModeBans2)
      TargetNick = Param(MassModeBans2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "b" Else Action = Action & "+b"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeUnbans2 <> "" Then
    For u2 = 1 To ParamCount(MassModeUnbans2)
      TargetNick = Param(MassModeUnbans2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "b" Else Action = Action & "-b"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeInvite2 <> "" Then
    For u2 = 1 To ParamCount(MassModeInvite2)
      TargetNick = Param(MassModeInvite2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "I" Else Action = Action & "+I"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeUnInvite2 <> "" Then
    For u2 = 1 To ParamCount(MassModeUnInvite2)
      TargetNick = Param(MassModeUnInvite2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "I" Else Action = Action & "-I"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeExcept2 <> "" Then
    For u2 = 1 To ParamCount(MassModeExcept2)
      TargetNick = Param(MassModeExcept2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 2 Then Action = Action & "e" Else Action = Action & "+e"
      LastOne = 2
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MassModeUnExcept2 <> "" Then
    For u2 = 1 To ParamCount(MassModeUnExcept2)
      TargetNick = Param(MassModeUnExcept2, u2)
      MultCount = MultCount + 1
      If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
      If LastOne = 1 Then Action = Action & "e" Else Action = Action & "-e"
      LastOne = 1
      If MultCount = ServerNumberOfModes Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = "": LastOne = 0
    Next u2
  End If
  If MultCount > 0 Then SendLine "mode " & Channel & " " & Action & " " & Multiple, 1: MultCount = 0: Multiple = "": Action = ""
  MassModeOps2 = "": MassModeDeops2 = "": MassModeBans2 = "": MassModeUnbans2 = "": MassModeExcept2 = "": MassModeUnExcept2 = "": MassModeInvite2 = "": MassModeUnInvite2 = "": MassModeHOps2 = "": MassModeDehops2 = "": MassModeVoice2 = "": MassModeDevoice2 = "":
End Sub

Function FindChan(Name As String) As Long
  Dim u As Long
  FindChan = 0
  If Name <> "" Then
    For u = 1 To ChanCount
      If LCase(Channels(u).Name) = LCase(Name) Then
        FindChan = u
        Exit Function
      End If
    Next u
  End If
End Function

Public Function DifferNick(OldOne As String) As String
  Dim NewOne As String, u As Long
  NewOne = Left(OldOne, 5)
  For u = 1 To 4
    NewOne = NewOne + Choose(Int(Rnd * 9) + 1, "-", "\", "`", "_", "|", "}", "{", "[", "]")
  Next u
  DifferNick = NewOne
End Function

Public Function RandNick() As String
  RandNick = Chr(Asc("a") + Int(Rnd * 27)) & Chr(Asc("0") + Int(Rnd * 10)) & Chr(Asc("a") + Int(Rnd * 27)) & Chr(Asc("0") + Int(Rnd * 10)) & Chr(Asc("a") + Int(Rnd * 27)) & Chr(Asc("0") + Int(Rnd * 10)) & Chr(Asc("a") + Int(Rnd * 27)) & Chr(Asc("0") + Int(Rnd * 10)) & Chr(Asc("a") + Int(Rnd * 27))
End Function

Public Function RandString() As String
  Dim u As Long, RStr As String
  For u = 1 To Int(Rnd * 20) + 1
    RStr = RStr + Chr(Asc("a") + Int(Rnd * 26))
  Next u
  RandString = RStr
End Function

Public Function GetNick(Hostmask) As String
  Dim u As Long
  u = InStr(Hostmask, "!")
  If u = 0 Then GetNick = "" Else GetNick = Left(Hostmask, u - 1)
End Function

Public Function StripDP(Strip As String) As String
  If Left(Strip, 1) = ":" Then StripDP = Mid(Strip, 2) Else StripDP = Strip
End Function

Function CombineModes(ByVal First As String, ByVal Second As String)
  Dim Mix As String, ModeKombi As String, ParamKombi As String
  If First = "" Then CombineModes = Second: Exit Function
  If Second = "" Then CombineModes = First: Exit Function
  If InStr(First, Chr(13)) > 0 Then First = Left(First, InStr(First, Chr(13)) - 1)
  If InStr(Second, Chr(13)) > 0 Then Second = Left(Second, InStr(Second, Chr(13)) - 1)
  ModeKombi = Param(First, 3) & Param(Second, 3)
  Mix = Param(First, 1) & " " & Param(First, 2) & " " & CleanModes(ModeKombi)
  If Param(First, 4) <> "" Then
    If Param(Second, 4) <> "" Then
      Mix = Mix & " " & Right(First, Len(First) - Len(Param(First, 1) & " " & Param(First, 2) & " " & Param(First, 3)) - 1) & " "
      Mix = Mix & Right(Second, Len(Second) - Len(Param(Second, 1) & " " & Param(Second, 2) & " " & Param(Second, 3)) - 1)
    Else
      Mix = Mix & " " & Right(First, Len(First) - Len(Param(First, 1) & " " & Param(First, 2) & " " & Param(First, 3)) - 1)
    End If
  Else
    If Param(Second, 4) <> "" Then
      Mix = Mix & " " & Right(Second, Len(Second) - Len(Param(Second, 1) & " " & Param(Second, 2) & " " & Param(Second, 3)) - 1)
    End If
  End If
  CombineModes = Mix + vbCrLf
End Function

Function CombineKicks(ByVal First As String, ByVal Second As String)
  Dim Mix As String, Buff As String
  If First = "" Then CombineKicks = Second: Exit Function
  If Second = "" Then CombineKicks = First: Exit Function
  If InStr(First, Chr(13)) > 0 Then First = Left(First, InStr(First, Chr(13)) - 1)
  Mix = "KICK " & Param(First, 2) & " " & Param(First, 3) & "," & Param(Second, 3)
  Buff = "KICK " & Param(First, 2) & " " & Param(First, 3)
  Mix = Mix & Right(First, Len(First) - Len(Buff))
  CombineKicks = Mix & vbCrLf
End Function

Sub PatternKickBan(Nick As String, Chan As String, Reason As String, TakeBack As Boolean, ShortTime As Currency, LongTime As Currency)
  Dim u2 As Long, ChNum As Long, UsNum As Long, NumOfKicks As Long
  Dim TheBanMask As String, DidTheBan As Boolean
  ChNum = FindChan(Chan)
  If Not (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) Then Exit Sub
  UsNum = FindUser(Nick, ChNum)
  For u2 = 1 To Channels(ChNum).UserCount
    If MatchWM(Mask(Channels(ChNum).User(UsNum).Hostmask, 2), Channels(ChNum).User(u2).Hostmask) And (Channels(ChNum).User(u2).Nick <> MyNick) Then
      NumOfKicks = NumOfKicks + 1
      If InStr(Channels(ChNum).User(u2).Status, "@") > 0 Then AddMassMode2 "-o", Channels(ChNum).User(u2).Nick
    End If
  Next u2
  If NumOfKicks > 1 Then
    TheBanMask = Mask(Channels(ChNum).User(UsNum).Hostmask, 2)
    If Not IsBanned(ChNum, TheBanMask) Then
      If Not IsOrdered("ban " & Chan & " " & TheBanMask) Then
        Order "ban " & Chan & " " & TheBanMask, 30
        AddMassMode2 "+b", TheBanMask
        DidTheBan = True
      End If
    End If
    DoMassMode2 Chan
    SendLine "kick " & Chan & " " & Nick & " :" & Reason, 1
    For u2 = 1 To Channels(ChNum).UserCount
      If MatchWM(Mask(Channels(ChNum).User(UsNum).Hostmask, 2), Channels(ChNum).User(u2).Hostmask) And (Channels(ChNum).User(u2).Nick <> MyNick) Then
        If Channels(ChNum).User(u2).Nick <> Nick Then SendLine "kick " & Chan & " " & Channels(ChNum).User(u2).Nick & " :Pattern kick of " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), 1
      End If
    Next u2
    If TakeBack And DidTheBan Then TimedEvent "UnBan " & Chan & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 2), LongTime
  Else
    MassModeOps2 = "": MassModeDeops2 = "": MassModeBans2 = "": MassModeUnbans2 = ""
    If InStr(Channels(ChNum).User(UsNum).Status, "@") > 0 Then
      SendLine "mode " & Chan & " -o+b " & Nick & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 3), 1
    Else
      SendLine "mode " & Chan & " +b " & Mask(Channels(ChNum).User(UsNum).Hostmask, 3), 1
    End If
    SendLine "kick " & Chan & " " & Nick & " :" & Reason, 1
    If TakeBack And DidTheBan Then TimedEvent "UnBan " & Chan & " " & Mask(Channels(ChNum).User(UsNum).Hostmask, 3), ShortTime
  End If
End Sub
'Returns a percentage from 0 to 100 about how much the given strings match
Function MatchGrade(Str1 As String, Str2 As String) As Long
Dim u As Long, u2 As Long, AtPos As Long, SliceSize As Long
Dim CStr1 As String, CStr2 As String, TempVar As String, OneSlice As String
Dim CStrL1 As Long, CStrL2 As Long, TempNum As Long
  CStr1 = LCase(Str1)
  CStr2 = LCase(Str2)
  If CStr1 = CStr2 Then MatchGrade = 100: Exit Function
  CStr1 = MakeEasy(CStr1)
  CStr2 = MakeEasy(CStr2)
  CStrL1 = Len(CStr1)
  CStrL2 = Len(CStr2)
  If CStrL2 > CStrL1 Then
    TempVar = CStr1: CStr1 = CStr2: CStr2 = TempVar
    TempNum = CStrL1: CStrL1 = CStrL2: CStrL2 = TempNum
  End If
  If InStr(CStr1, CStr2) > 0 Then MatchGrade = Int(CStrL2 / CStrL1 * 100): MatchGrade = MatchGrade + ((100 - MatchGrade) \ 4): Exit Function
  
  'Calculate the size of the used 'slices'
  SliceSize = Int(CStrL1 * 0.35): If SliceSize < 1 Then SliceSize = 1
  
  For u = 1 To CStrL1 Step SliceSize
    OneSlice = Mid(CStr2, u, SliceSize)
    For u2 = 1 To SliceSize
      If (u + u2 - 1) <= CStrL1 Then
        TempNum = InStr(OneSlice, Mid(CStr1, u + u2 - 1, 1))
        If TempNum > 0 Then Mid(OneSlice, TempNum, 1) = "?": AtPos = AtPos + 1
      End If
    Next u2
  Next u
  MatchGrade = Int(95 / CStrL1 * AtPos)
End Function

'Removes double chars (lowercase string needed)
Function MakeEasy(Str1 As String) As String
Dim Result As String, Char As String, u As Long
  For u = 1 To Len(Str1)
    Char = Mid(Str1, u, 1)
    If InStr("\[{}]", Char) > 0 Then Char = "|"
    Result = Result + Char
  Next u
  MakeEasy = Result
End Function

'Strips all color codes from a line
Public Function Strip(ByVal Line As String) As String
  Dim SPos As Long, OneChar As String, lng As Long, leng As Long, StripAway As Boolean
  Do
    StripAway = True
    SPos = InStr(Line, "")
    If SPos = 0 Then SPos = InStr(Line, ""): StripAway = False
    If SPos = 0 Then SPos = InStr(Line, ""): StripAway = False
    If SPos = 0 Then Exit Do
    If StripAway Then
      For lng = SPos To Len(Line) + 1
        '- 12.07.2004: Fixed bug: Strip("12345abc") = "abc"
        'OneChar = Mid$(Line, lng, 1)
        'If OneChar = "" Then leng = lng - SPos: Exit For
        'If (OneChar < "0" Or OneChar > "9") Or (SPos = (lng - 3)) Then
        '  If OneChar = "," Then
        '    If Mid$(Line, lng + 1, 1) < "0" Or Mid$(Line, lng + 1, 1) > "9" Then leng = lng - SPos: Exit For
        '  Else
        '    leng = lng - SPos: Exit For
        '  End If
        'End If
        '- /12.07.2004
        '+ 12.07.2004: Fixed bug: Strip("12345abc") = "abc"
        Select Case True
          Case Mid$(Line, lng, 6) Like "##,##"
            Mid$(Line, lng, 6) = ""
          Case Mid$(Line, lng, 5) Like "##,#"
            Mid$(Line, lng, 5) = ""
          Case Mid$(Line, lng, 4) Like "##,"
            Mid$(Line, lng, 4) = ""
          Case Mid$(Line, lng, 3) Like "##"
            Mid$(Line, lng, 3) = ""
          Case Mid$(Line, lng, 5) Like "#,##"
            Mid$(Line, lng, 5) = ""
          Case Mid$(Line, lng, 4) Like "#,#"
            Mid$(Line, lng, 4) = ""
          Case Mid$(Line, lng, 2) Like "#"
            Mid$(Line, lng, 2) = ""
          Case Mid$(Line, lng, 1) Like ""
            Mid$(Line, lng, 1) = ""
        End Select
        '+ /12.07.2004
      Next lng
    Else
      leng = 1
    End If
    If SPos = 1 Then Line = Right$(Line, Len(Line) - leng)
    If SPos = Len(Line) Then
      Line = Left$(Line, SPos - 1)
    ElseIf SPos > 1 Then
      Line = Left$(Line, SPos - 1) + Right$(Line, Len(Line) - SPos - leng + 1)
    End If
  Loop
  Strip = Replace(Line, "", "")
End Function

Public Function CompressedModes(ByVal What As String, BufNum As Long) As Boolean ' : AddStack "mdlWinsock_CompressedModes(" & What & ", " & BufNum & ")"
Dim u As Long, u2 As Long, Line As String, Count As Long, Piece As String, GotMode As String
  If LCase(Param(What, 1)) = "mode" And (Len(Param(What, 3)) = 2 Or Param(What, 3) = "+smtin" Or Param(What, 3) = "+stin" Or Param(What, 3) = "+m") Then
    For u = 1 To Buffer(BufNum).BufferedLines
      Line = Buffer(BufNum).LineBuffer(u)
      If LCase(Param(Line, 1)) = "mode" Then
        If LCase(Param(Line, 2)) = LCase(Param(What, 2)) Then
          If InStr(LCase(Param(Line, 3)), "k") = 0 Then
            Piece = LCase(Param(Line, 3))
            Count = 0
            For u2 = 1 To Len(Piece)
              Select Case Mid(Piece, u2, 1)
                Case "o", "v", "b", "I", "e": Count = Count + 1
              End Select
            Next u2
            If Count < ServerNumberOfModes Or Param(What, 3) = "+smtin" Or Param(What, 3) = "+stin" Or Param(What, 3) = "+m" Then
              'Don't allow MODE +b (banlist check) to be combined with other things
              If InStr(Line, Chr(13)) > 0 Then Piece = Left(Line, InStr(Line, Chr(13)) - 1) Else Piece = Line
              If LCase(Param(Piece, 3)) = "+b" And Param(Piece, 4) = "" Then CompressedModes = False: Exit Function
              If LCase(Param(Piece, 3)) = "+e" And Param(Piece, 4) = "" Then CompressedModes = False: Exit Function
              If LCase(Param(Piece, 3)) = "+I" And Param(Piece, 4) = "" Then CompressedModes = False: Exit Function
              GotMode = CombineModes(Line, What)
              If GotMode <> "" Then
                Buffer(BufNum).LineBuffer(u) = GotMode
              Else
                Buffer(BufNum).BufferedLines = Buffer(BufNum).BufferedLines + 1
                If Buffer(BufNum).BufferedLines > UBound(Buffer(BufNum).LineBuffer()) Then ReDim Preserve Buffer(BufNum).LineBuffer(UBound(Buffer(BufNum).LineBuffer()) + 5)
                Buffer(BufNum).LineBuffer(Buffer(BufNum).BufferedLines) = What + vbCrLf
              End If
              CompressedModes = True
              Exit Function
            End If
          End If
        End If
      End If
    Next u
  End If
  CompressedModes = False
End Function

Public Function CompressedKicks(ByVal What As String, BufNum As Long) As Boolean ' : AddStack "mdlWinsock_CompressedKicks(" & What & ", " & BufNum & ")"
  Dim u As Long, u2 As Long, Line As String, Count As Long, Piece As String
  Dim OneNick As String
  If ServerInfo.SupportsMultiKicks = False Then CompressedKicks = False: Exit Function
  If LCase(Param(What, 1)) = "kick" Then
    For u = 1 To Buffer(BufNum).BufferedLines
      Line = Buffer(BufNum).LineBuffer(u)
      If LCase(Param(Line, 1)) = "kick" Then
        If LCase(Param(Line, 2)) = LCase(Param(What, 2)) Then
          Piece = LCase(Param(Line, 3))
          Count = 1
          OneNick = ""
          For u2 = 1 To Len(Piece)
            If Mid(Piece, u2, 1) = "," Then
              Count = Count + 1
              If OneNick = Param(What, 3) Then CompressedKicks = True: Exit Function
              OneNick = ""
            Else
              OneNick = OneNick + Mid(Piece, u2, 1)
            End If
          Next u2
          If OneNick = Param(What, 3) Then CompressedKicks = True: Exit Function
          If Count < 4 Then
            Buffer(BufNum).LineBuffer(u) = CombineKicks(Line, What)
            CompressedKicks = True
            Exit Function
          End If
        End If
      End If
    Next u
  End If
  CompressedKicks = False
End Function

Public Sub UpdateBufferNicks(OldNick As String, NewNick As String) ' : AddStack "mdlWinsock_UpdateBufferNicks(" & OldNick & ", " & NewNick & ")"
Dim u As Long, u2 As Long, NewBuf As String, TheNicks As String, Nick As String, ONick  As String, nonick As String
Dim NewNicks As String, ExchangedOne As Boolean, BufNum As Long, PosInONick As Long
Dim OpString As String, PlusOrMinus As Byte
  For BufNum = 1 To 3
    For u = 1 To Buffer(BufNum).BufferedLines
      If LCase(Param(Buffer(BufNum).LineBuffer(u), 1)) = "kick" Then
        TheNicks = Param(Buffer(BufNum).LineBuffer(u), 3)
        If InStr(TheNicks, Chr(13)) > 0 Then TheNicks = Left(TheNicks, InStr(TheNicks, Chr(13)) - 1)
        Nick = "": NewNicks = "": ExchangedOne = False
        For u2 = 1 To Len(TheNicks)
          Select Case Mid(TheNicks, u2, 1)
            Case ","
              If LCase(Nick) = LCase(OldNick) Then Nick = NewNick: ExchangedOne = True
              If NewNicks <> "" Then NewNicks = NewNicks & "," & Nick Else NewNicks = Nick
              Nick = ""
            Case Else: Nick = Nick + Mid(TheNicks, u2, 1)
          End Select
        Next u2
        If LCase(Nick) = LCase(OldNick) Then Nick = NewNick: ExchangedOne = True
        If Nick <> "" Then If NewNicks <> "" Then NewNicks = NewNicks & "," & Nick Else NewNicks = Nick
        If ExchangedOne Then
          Buffer(BufNum).LineBuffer(u) = "KICK " & Param(Buffer(BufNum).LineBuffer(u), 2) & " " & NewNicks + Right(Buffer(BufNum).LineBuffer(u), Len(Buffer(BufNum).LineBuffer(u)) - Len("KICK " & Param(Buffer(BufNum).LineBuffer(u), 2) & " " & Param(Buffer(BufNum).LineBuffer(u), 3)))
          If Right(Buffer(BufNum).LineBuffer(u), 2) <> vbCrLf Then Buffer(BufNum).LineBuffer(u) = Buffer(BufNum).LineBuffer(u) + vbCrLf
        End If
      End If
      If LCase(Param(Buffer(BufNum).LineBuffer(u), 1)) = "mode" Then
        TheNicks = Buffer(BufNum).LineBuffer(u)
        If InStr(TheNicks, Chr(13)) > 0 Then TheNicks = Left(TheNicks, InStr(TheNicks, Chr(13)) - 1)
        OpString = Param(TheNicks, 3)
        nonick = "": ONick = Param(TheNicks, 4): ExchangedOne = False
        If Param(TheNicks, 5) <> "" Then ONick = ONick & " " & Param(TheNicks, 5)
        If Param(TheNicks, 6) <> "" Then ONick = ONick & " " & Param(TheNicks, 6)
        If Param(TheNicks, 7) <> "" Then ONick = ONick & " " & Param(TheNicks, 7)
        PosInONick = 0
        For u2 = 1 To Len(OpString)
          Select Case Mid(OpString, u2, 1)
            Case "+": PlusOrMinus = 1
            Case "-": PlusOrMinus = 2
            Case "l": If GetModeChar(OpString, "l") = 1 Then PosInONick = PosInONick + 1: nonick = nonick & " " & Param(ONick, PosInONick)
            Case "k": PosInONick = PosInONick + 1: nonick = nonick & " " & Param(ONick, PosInONick)
            Case "b", "I", "e": PosInONick = PosInONick + 1: nonick = nonick & " " & Param(ONick, PosInONick)
            Case "v", "o": PosInONick = PosInONick + 1
              If LCase(Param(ONick, PosInONick)) = LCase(OldNick) Then
                nonick = nonick & " " & NewNick: ExchangedOne = True
              Else
                nonick = nonick & " " & Param(ONick, PosInONick)
              End If
          End Select
        Next u2
        If ExchangedOne Then
          Buffer(BufNum).LineBuffer(u) = "mode " & Param(TheNicks, 2) & " " & OpString & " " & nonick
          If Right(Buffer(BufNum).LineBuffer(u), 2) <> vbCrLf Then Buffer(BufNum).LineBuffer(u) = Buffer(BufNum).LineBuffer(u) + vbCrLf
        End If
      End If
    Next u
  Next BufNum
End Sub

Public Sub UpdateBufferKicks(KickedNick As String, Chan As String) ' : AddStack "mdlWinsock_UpdateBufferKicks(" & KickedNick & ", " & Chan & ")"
Dim u As Long, u2 As Long, NewBuf As String, TheNicks As String, Nick As String, ONick As String, nonick As String
Dim NewNicks As String, ExchangedOne As Boolean, BufNum As Long, PosInONick As Long
Dim OpString As String, PlusOrMinus As Byte, ModeChars As String
  For BufNum = 1 To 3
    For u = 1 To Buffer(BufNum).BufferedLines
      If LCase(Param(Buffer(BufNum).LineBuffer(u), 1)) = "kick" And LCase(Param(Buffer(BufNum).LineBuffer(u), 2)) = LCase(Chan) Then
        TheNicks = Param(Buffer(BufNum).LineBuffer(u), 3)
        If InStr(TheNicks, Chr(13)) > 0 Then TheNicks = Left(TheNicks, InStr(TheNicks, Chr(13)) - 1)
        Nick = "": NewNicks = "": ExchangedOne = False
        For u2 = 1 To Len(TheNicks)
          Select Case Mid(TheNicks, u2, 1)
            Case ","
              If Nick = KickedNick Then Nick = "": ExchangedOne = True
              If Nick <> "" Then If NewNicks <> "" Then NewNicks = NewNicks & "," & Nick Else NewNicks = Nick
              Nick = ""
            Case Else: Nick = Nick + Mid(TheNicks, u2, 1)
          End Select
        Next u2
        If Nick = KickedNick Then Nick = "": ExchangedOne = True
        If Nick <> "" Then If NewNicks <> "" Then NewNicks = NewNicks & "," & Nick Else NewNicks = Nick
        If ExchangedOne Then
          If NewNicks <> "" Then
            Buffer(BufNum).LineBuffer(u) = "KICK " & Param(Buffer(BufNum).LineBuffer(u), 2) & " " & NewNicks + Right(Buffer(BufNum).LineBuffer(u), Len(Buffer(BufNum).LineBuffer(u)) - Len("KICK " & Param(Buffer(BufNum).LineBuffer(u), 2) & " " & Param(Buffer(BufNum).LineBuffer(u), 3)))
            If Right(Buffer(BufNum).LineBuffer(u), 2) <> vbCrLf Then Buffer(BufNum).LineBuffer(u) = Buffer(BufNum).LineBuffer(u) + vbCrLf
          Else
            For u2 = u To Buffer(BufNum).BufferedLines - 1
              Buffer(BufNum).LineBuffer(u2) = Buffer(BufNum).LineBuffer(u2 + 1)
            Next u2
            Buffer(BufNum).BufferedLines = Buffer(BufNum).BufferedLines - 1
          End If
        End If
      End If
      If LCase(Param(Buffer(BufNum).LineBuffer(u), 1)) = "mode" And LCase(Param(Buffer(BufNum).LineBuffer(u), 2)) = LCase(Chan) Then
        TheNicks = Buffer(BufNum).LineBuffer(u)
        If InStr(TheNicks, Chr(13)) > 0 Then TheNicks = Left(TheNicks, InStr(TheNicks, Chr(13)) - 1)
        OpString = Param(TheNicks, 3): ModeChars = ""
        nonick = "": ONick = Param(TheNicks, 4): ExchangedOne = False
        If Param(TheNicks, 5) <> "" Then ONick = ONick & " " & Param(TheNicks, 5)
        If Param(TheNicks, 6) <> "" Then ONick = ONick & " " & Param(TheNicks, 6)
        If Param(TheNicks, 7) <> "" Then ONick = ONick & " " & Param(TheNicks, 7)
        If Param(TheNicks, 8) <> "" Then ONick = ONick & " " & Param(TheNicks, 8)
        PosInONick = 0
        For u2 = 1 To Len(OpString)
          Select Case Mid(OpString, u2, 1)
            Case "+": PlusOrMinus = 1: ModeChars = ModeChars + Mid(OpString, u2, 1)
            Case "-": PlusOrMinus = 2: ModeChars = ModeChars + Mid(OpString, u2, 1)
            Case "l": ModeChars = ModeChars + Mid(OpString, u2, 1): If GetModeChar(OpString, "l") = 1 Then PosInONick = PosInONick + 1: nonick = nonick & " " & Param(ONick, PosInONick)
            Case "k": PosInONick = PosInONick + 1: nonick = nonick & " " & Param(ONick, PosInONick): ModeChars = ModeChars + Mid(OpString, u2, 1)
            Case "b", "I", "e": PosInONick = PosInONick + 1: nonick = nonick & " " & Param(ONick, PosInONick): ModeChars = ModeChars + Mid(OpString, u2, 1)
            Case "v", "o": PosInONick = PosInONick + 1
              If LCase(Param(ONick, PosInONick)) = LCase(KickedNick) Then
                ExchangedOne = True
              Else
                nonick = nonick & " " & Param(ONick, PosInONick): ModeChars = ModeChars + Mid(OpString, u2, 1)
              End If
            Case Else: ModeChars = ModeChars + Mid(OpString, u2, 1)
          End Select
        Next u2
        If ExchangedOne Then
          Buffer(BufNum).LineBuffer(u) = "mode " & Param(TheNicks, 2) & " " & ModeChars & " " & nonick
          If Right(Buffer(BufNum).LineBuffer(u), 2) <> vbCrLf Then Buffer(BufNum).LineBuffer(u) = Buffer(BufNum).LineBuffer(u) + vbCrLf
        End If
      End If
    Next u
  Next BufNum
End Sub

Sub UpdateRegUsers(Action As String) ' : AddStack "mdlWinsock_UpdateRegUsers(" & Action & ")"
Dim u As Long, u2 As Long, Rest As String
  Select Case Param(Action, 1)
    Case "A" 'Add
        Rest = Param(Action, 2)
        For u = 1 To ChanCount
          For u2 = 1 To Channels(u).UserCount
            'Only check for RegUser change if hostmask matches newly added host
            If (MatchHost(Rest, Channels(u).User(u2).Hostmask) = True) Or (Rest = "") Then
              Channels(u).User(u2).RegNick = SearchUserFromHostmask(Channels(u).User(u2).Hostmask)
              Channels(u).User(u2).UserNum = BotUserNum
            End If
          Next u2
        Next u
    Case "I" 'Ident
        Rest = Param(Action, 2)
        For u = 1 To ChanCount
          For u2 = 1 To Channels(u).UserCount
            If Channels(u).User(u2).Nick = Rest Then
              Channels(u).User(u2).RegNick = Param(Action, 3)
              Channels(u).User(u2).UserNum = GetUserNum(Param(Action, 3))
            End If
          Next u2
        Next u
    Case "R" 'Remove
        For u = 1 To ChanCount
          For u2 = 1 To Channels(u).UserCount
            If Channels(u).User(u2).RegNick = Param(Action, 2) Then
              Channels(u).User(u2).RegNick = SearchUserFromHostmask(Channels(u).User(u2).Hostmask)
              Channels(u).User(u2).UserNum = BotUserNum
            End If
          Next u2
        Next u
    Case Else
        For u = 1 To ChanCount
          For u2 = 1 To Channels(u).UserCount
            Channels(u).User(u2).RegNick = SearchUserFromHostmask(Channels(u).User(u2).Hostmask)
            Channels(u).User(u2).UserNum = BotUserNum
          Next u2
        Next u
  End Select
End Sub

Sub MakeSessionValid(UserNum As Long)
  Dim i As Long, j As Long
  BotUsers(UserNum).ValidSession = True
  For i = 0 To ChanCount
    For j = 0 To Channels(i).UserCount
      If Channels(i).User(j).RegNick = BotUsers(UserNum).Name Then
        BotUsers(UserNum).OldFullAddress = Mid(Channels(i).User(j).Hostmask, InStr(Channels(i).User(j).Hostmask, "@") + 1)
        Exit Sub
      End If
    Next j
  Next i
End Sub

Sub MakeSessionInValid(UserNum As Long)
  BotUsers(UserNum).ValidSession = False
  BotUsers(UserNum).OldFullAddress = "srtpohwqu45908qvj059q"
End Sub

Function IrcGetAscIp(ByVal IPL As String) As String
  Dim Inn As Long
  Dim Char(0 To 3) As Byte
  If Val(IPL) > 2147483647 Then
      Inn = Val(IPL) - 4294967296#
  Else
      Inn = Val(IPL)
  End If
  kernel32_RtlMoveMemory Char(0), Inn, 4
  IrcGetAscIp = Char(3) & "." & Char(2) & "." & Char(1) & "." & Char(0)
End Function

Function IrcGetLongIp(ByVal AscIp As String) As String
  Dim Char(0 To 3) As Byte
  Dim Inn As Long
  If ParamXCount(AscIp, ".") = 4 Then
    On Error GoTo IrcGetLongIpError
    Char(3) = ParamX(AscIp, ".", 1)
    Char(2) = ParamX(AscIp, ".", 2)
    Char(1) = ParamX(AscIp, ".", 3)
    Char(0) = ParamX(AscIp, ".", 4)
    kernel32_RtlMoveMemory Inn, Char(0), 4
    If Inn < 0 Then
      IrcGetLongIp = CVar(Inn + 4294967296#)
      Exit Function
    Else
      IrcGetLongIp = CVar(Inn)
      Exit Function
    End If
  Else
    IrcGetLongIp = "0"
  End If
  Exit Function
IrcGetLongIpError:
  IrcGetLongIp = "0"
  Err.Clear
  Exit Function
End Function

Public Function FormatMode(OldMode As String, ModeChanges As String) As String ' : AddStack "Routines_FormatMode(" & OldMode & ", " & ModeChanges & ")"
Dim u As Long, NewMode As String, CurMode As Long, CurPos As Long
Dim LimitPos As Long, KeyPos As Long
  NewMode = "+"
  CurMode = GetModeChar(OldMode, "s")
  Select Case GetModeChar(ModeChanges, "s")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "s"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "s"
  End Select
  CurMode = GetModeChar(OldMode, "p")
  Select Case GetModeChar(ModeChanges, "p")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "p"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "p"
  End Select
  CurMode = GetModeChar(OldMode, "m")
  Select Case GetModeChar(ModeChanges, "m")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "m"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "m"
  End Select
  CurMode = GetModeChar(OldMode, "t")
  Select Case GetModeChar(ModeChanges, "t")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "t"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "t"
  End Select
  CurMode = GetModeChar(OldMode, "i")
  Select Case GetModeChar(ModeChanges, "i")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "i"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "i"
  End Select
  CurMode = GetModeChar(OldMode, "n")
  Select Case GetModeChar(ModeChanges, "n")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "n"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "n"
  End Select
  CurMode = GetModeChar(OldMode, "l")
  Select Case GetModeChar(ModeChanges, "l")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "l"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "l"
  End Select
  CurMode = GetModeChar(OldMode, "k")
  Select Case GetModeChar(ModeChanges, "k")
    Case 1: If CurMode <> -1 Then NewMode = NewMode & "k"
    Case 0: If CurMode = 1 Then NewMode = NewMode & "k"
  End Select
  If InStr(ServerChannelModes, "c") Then
    CurMode = GetModeChar(OldMode, "c")
    Select Case GetModeChar(ModeChanges, "c")
      Case 1: If CurMode <> -1 Then NewMode = NewMode & "c"
      Case 0: If CurMode = 1 Then NewMode = NewMode & "c"
    End Select
  End If
  If InStr(ServerChannelModes, "C") Then
    CurMode = GetModeChar(OldMode, "C")
    Select Case GetModeChar(ModeChanges, "C")
      Case 1: If CurMode <> -1 Then NewMode = NewMode & "C"
      Case 0: If CurMode = 1 Then NewMode = NewMode & "C"
    End Select
  End If
  For u = 1 To Len(ModeChanges)
    Select Case Mid(ModeChanges, u, 1)
      Case "o": CurPos = CurPos + 1
      Case "b": CurPos = CurPos + 1
      Case "I": CurPos = CurPos + 1
      Case "e": CurPos = CurPos + 1
      Case "l": If GetModeChar(ModeChanges, "l") = 1 Then CurPos = CurPos + 1: LimitPos = CurPos + 1
      Case "k": CurPos = CurPos + 1: KeyPos = CurPos + 1
      Case " ": Exit For
    End Select
  Next u
  If InStr(NewMode, "l") > 0 Then
    If InStr(OldMode, "l") > 0 And GetModeChar(ModeChanges, "l") = 0 Then NewMode = NewMode & " " & Param(OldMode, 2)
    If GetModeChar(ModeChanges, "l") > 0 Then NewMode = NewMode & " " & Param(ModeChanges, LimitPos)
  End If
  If InStr(NewMode, "k") > 0 Then
    If InStr(OldMode, "k") > 0 Then
      NewMode = NewMode & " " & Param(OldMode, ParamCount(OldMode))
    ElseIf InStr(ModeChanges, "k") > 0 Then
      NewMode = NewMode & " " & Param(ModeChanges, KeyPos)
    End If
  End If
  FormatMode = NewMode
End Function

Public Function ChangeMode(Should As String, Current As String) ' : AddStack "Routines_ChangeMode(" & Should & ", " & Current & ")"
Dim u As Long, CurMode As Long, Changes As String, InsertWhat As String
Dim CurPos As Long, LimitPos As Long, KeyPos As Long
  CurMode = GetModeChar(Current, "p")
  Select Case GetModeChar(Should, "p")
    Case -1: If CurMode = 1 Then Changes = Changes & "-p"
    Case 1: If CurMode = 0 Then Changes = Changes & "+p"
  End Select
  CurMode = GetModeChar(Current, "s")
  Select Case GetModeChar(Should, "s")
    Case -1: If CurMode = 1 Then Changes = Changes & "-s"
    Case 1: If CurMode = 0 Then Changes = Changes & "+s"
  End Select
  CurMode = GetModeChar(Current, "m")
  Select Case GetModeChar(Should, "m")
    Case -1: If CurMode = 1 Then Changes = Changes & "-m"
    Case 1: If CurMode = 0 Then Changes = Changes & "+m"
  End Select
  CurMode = GetModeChar(Current, "t")
  Select Case GetModeChar(Should, "t")
    Case -1: If CurMode = 1 Then Changes = Changes & "-t"
    Case 1: If CurMode = 0 Then Changes = Changes & "+t"
  End Select
  CurMode = GetModeChar(Current, "i")
  Select Case GetModeChar(Should, "i")
    Case -1: If CurMode = 1 Then Changes = Changes & "-i"
    Case 1: If CurMode = 0 Then Changes = Changes & "+i"
  End Select
  CurMode = GetModeChar(Current, "n")
  Select Case GetModeChar(Should, "n")
    Case -1: If CurMode = 1 Then Changes = Changes & "-n"
    Case 1: If CurMode = 0 Then Changes = Changes & "+n"
  End Select
  If InStr(ServerChannelModes, "c") Then
    CurMode = GetModeChar(Current, "c")
    Select Case GetModeChar(Should, "c")
      Case -1: If CurMode = 1 Then Changes = Changes & "-c"
      Case 1: If CurMode = 0 Then Changes = Changes & "+c"
    End Select
  End If
  If InStr(ServerChannelModes, "C") Then
    CurMode = GetModeChar(Current, "C")
    Select Case GetModeChar(Should, "C")
      Case -1: If CurMode = 1 Then Changes = Changes & "-C"
      Case 1: If CurMode = 0 Then Changes = Changes & "+C"
    End Select
  End If
  
  For u = 1 To Len(Should)
    Select Case Mid(Should, u, 1)
      Case "l": If GetModeChar(Should, "l") = 1 Then CurPos = CurPos + 1: LimitPos = CurPos + 1
      Case "k": CurPos = CurPos + 1: KeyPos = CurPos + 1
      Case " ": Exit For
    End Select
  Next u
  
  CurMode = GetModeChar(Current, "l")
  Select Case GetModeChar(Should, "l")
    Case -1: If CurMode = 1 Then Changes = Changes & "-l"
    Case 1
      If CurMode = 0 Then Changes = Changes & "+l": InsertWhat = " " & Param(Should, LimitPos)
      If CurMode = 1 Then If Param(Current, 2) <> Param(Should, LimitPos) Then Changes = Changes & "+l": InsertWhat = " " & Param(Should, LimitPos)
  End Select
  CurMode = GetModeChar(Current, "k")
  Select Case GetModeChar(Should, "k")
    Case -1: If CurMode = 1 Then Changes = Changes & "-k" & InsertWhat & " " & Param(Current, ParamCount(Current)): InsertWhat = ""
    Case 1: If CurMode = 0 Then Changes = Changes & "+k" & InsertWhat & " " & Param(Should, KeyPos): InsertWhat = ""
  End Select
    
  Changes = CleanModes(Changes)
  ChangeMode = Changes & InsertWhat
End Function

Public Function CleanModes(OldM As String) As String
  Dim Pos As Byte
  Dim Index As Integer
  Dim NewM As String
  If OldM = "" Then Exit Function
  Pos = 2
  NewM = ""
  For Index = 1 To Len(OldM)
    Select Case Mid(OldM, Index, 1)
      Case "+"
        If Pos <> 1 Then NewM = NewM & "+": Pos = 1
      Case "-"
        If Pos <> 0 Then NewM = NewM & "-": Pos = 0
      Case Else
        NewM = NewM & Mid(OldM, Index, 1)
    End Select
  Next Index
  CleanModes = NewM
End Function

Public Function GetModeChar(strMode As String, Char As String) As Long ' : AddStack "Routines_GetModeChar(" & strMode & ", " & Char & ")"
Dim u As Long, Positive As Boolean
  For u = 1 To Len(strMode)
    Select Case Mid(strMode, u, 1)
      Case "+": Positive = True
      Case "-": Positive = False
      Case " ": Exit For
      Case Char
        If Positive Then GetModeChar = 1 Else GetModeChar = -1
        Exit Function
    End Select
  Next u
  GetModeChar = 0
End Function

'Connects to an IRC server
Sub ConnectServer(Seconds As Currency, SpecialServer As String) ' : AddStack "Server_ConnectServer(" & Seconds & ", " & SpecialServer & ")"
Dim TheEntry As String, TheServer As String, ThePort As Long, blubb As Long
Dim t As Single, Tim As Single, u As Long, Proxy As String
On Error GoTo Err2
  If (DontConnect = True) Or (HubBot = True) Then Exit Sub
  If Seconds > 0 Then
    TimedEvent "ConnectServer " & SpecialServer, Seconds
    Exit Sub
  End If
  blubb = 1
  GUI_frmWinsock.ConnectTimeOut.Enabled = False
  
  'CheckRestart  'No connection since 15 tries? Restart...
  If IsTimed("RESTART") Then Exit Sub
    
  LastEvent = Now
    
  blubb = 2
  'Get server:port to connect to
  If SpecialServer = "" Then
    UsedServer = UsedServer + 1
    TheEntry = GetPPString("Server", IIf(UsedServer = 1, "Server", "Server" & CStr(UsedServer)), "", AnGeL_INI)
    If TheEntry = "" Then UsedServer = 1: TheEntry = MakeNormalNick(GetPPString("Server", "Server", "", AnGeL_INI))
    JustJumped = False
  Else
    TheEntry = SpecialServer
    JustJumped = True
  End If

  'Get proxy, if specified
  ProxyToConnect = GetPPString("Identification", "Proxy", "", AnGeL_INI)
  If GetProxy(TheEntry) <> "" Then ProxyToConnect = GetProxy(TheEntry)
  ServerToConnect = GetServer(TheEntry)
  SentLogin = False
  
  'Parse <server>:<port>
  On Error GoTo Err2
  If ProxyToConnect = "" Then
    TheServer = GetAddr(ServerToConnect)
    ThePort = GetPort(ServerToConnect, 6667)
  Else
    TheServer = GetAddr(ProxyToConnect)
    ThePort = GetPort(ProxyToConnect, 1080)
  End If
    
  If (GetCacheIP(TheServer, True) = "255.255.255.255") Or (TheServer = "") Then
    SpreadFlagMessage 0, "+m", MakeMsg(ERR_ServerFailed, TheServer, "Unable to resolve address.")
    SetTrayIcon SI_Offline
    ConnectServer CurrentConnectDelay, ""
    Exit Sub
  End If

  blubb = 3
  BlockSends = False
  Disconnect
  
  'Trying to get Ident Socket
  CloseIdentSockets
  If ProxyToConnect = "" And UseIDENTD = True Then
    GetIdentSocket 0, TheServer, ThePort
  Else
    SpreadFlagMessage 0, "+m", "3*** Disabled Ident (" & IIf(ProxyToConnect = "", "IDENTD disabled", "using proxy server") & ")..."
    ConnectNow TheServer, ThePort
  End If

Exit Sub
Err2:
  Err.Clear
  MsgBox "2.Fehler ist an Stelle Nr. " & CStr(blubb) & " aufgetreten."
  End
End Sub

Public Sub GetIdentSocket(TimesLeft As Long, TheServer As String, ThePort As Long) ' : AddStack "Server_GetIdentSocket(" & TimesLeft & ", " & TheServer & ", " & ThePort & ")"
  Dim u As Long
  
  'Ident Socket already open?! (shouldn't happen) - Close it!
  If IdentSocket > 0 Then CloseIdentSockets
  
  'Try to get Ident Socket
  IdentSocket = AddSocket
  u = ListenTCP(IdentSocket, 113)
  If u = 0 Then
    'Success -> set vsock -> connect to server
    SocketItem(IdentSocket).Used = True
    SocketItem(IdentSocket).IsInternalSocket = True
    SocketItem(IdentSocket).RegNick = "<IDENTD>"
    SetSockFlag IdentSocket, SF_Status, SF_Status_IdentListen
    SpreadFlagMessage 0, "+m", "3*** Got the Ident Socket (" & CStr(IdentSocket) & ")"
    ConnectNow TheServer, ThePort
  Else
    DisconnectSocket IdentSocket
    IdentSocket = -1
    'Error
    If TimesLeft <= 0 Then
      'First time this fails -> tell user, try again 5 times
      SpreadFlagMessage 0, "+m", "3*** Trying to get the Ident Socket... (max 90 secs)"
      TimedEvent "GetIdentSocket 30 " & TheServer & " " & Trim(Str(ThePort)), 3
      Exit Sub
    ElseIf TimesLeft = 1 Then
      'Failed 5 times -> give up
      SpreadFlagMessage 0, "+m", "4*** I couldn't get the Ident Socket!"
      IdentSocket = 0
      ConnectNow TheServer, ThePort
    Else
      'Failed, decrease TimesLeft counter and try again in 3 seconds
      TimedEvent "GetIdentSocket " & Trim(Str(TimesLeft - 1)) & " " & TheServer & " " & Trim(Str(ThePort)), 3
    End If
  End If
End Sub

Public Sub ConnectNow(TheServer As String, ThePort As Long) ' : AddStack "Server_ConnectNow(" & TheServer & ", " & ThePort & ")"
  Dim Result As Long
  ServerSocket = AddSocket
  
  Result = ConnectTCP(ServerSocket, TheServer, ThePort)
  If Result = 0 Then
    SetSockFlag ServerSocket, SF_Status, SF_Status_Server
    SocketItem(ServerSocket).RegNick = "<SERVER>"
    ServerPort = SocketItem(ServerSocket).LocalPort
    SpreadFlagMessage 0, "+m", "3*** Trying to connect to " & TheServer & " 10(" & SocketItem(ServerSocket).RemoteHost & ")3..."
    Status "*** Connecting to " & IIf(ProxyToConnect = "", "server", "proxy") & "..." & vbCrLf
    SetTrayIcon SI_Connecting
    GUI_frmWinsock.cmdConnect.Enabled = False
    GotServerPong = True
  Else
    SpreadFlagMessage 0, "+m", "3*** Connect to " & TheServer & " 4failed3... " & Result & " - " & WSAGetErrorString(Result)
    Status "*** Connect to " & IIf(ProxyToConnect = "", "server", "proxy") & " failed..." & vbCrLf
    DisconnectSocket ServerSocket
    ServerSocket = -1
  End If
  GUI_frmWinsock.ConnectTimeOut.Enabled = True
End Sub

'Disconnects from IRC server and resets all variables
Public Sub Disconnect() ' : AddStack "Server_Disconnect()"
  Dim u As Long, TmpStr As String
  If ServerSocket > 0 Then
    DisconnectSocket ServerSocket
    ServerSocket = -1
  End If
  CloseIdentSockets
  MyNick = ""
  ServerName = ""
  BlockSends = False
  OrderCount = 0: ChanCount = 0
  'Erase server based events
  For u = 1 To EventCount
    TmpStr = Param(Events(u).DoThis, 1)
    Select Case TmpStr
      Case "CallScript", "PutSendQ", "winsock2_shutdown", "RESTART", "ConnectServer", "GetIdentSocket", "RemKI", "UnIgnore", "KickOldBot", "AddTrayIcon", "FinalBotNetLogin", "UPTIME"
        Events(u).DoThis = Events(u).DoThis
      Case Else
        Events(u).DoThis = ""
    End Select
  Next u
  IgnoredUserCount = 0: Buffer(1).BufferedLines = 0: Buffer(2).BufferedLines = 0: Buffer(3).BufferedLines = 0: BytesSent = 0
  LineToComplete = ""
  GUI_frmWinsock.lstChannels.Clear
  GUI_frmWinsock.cmdConnect.Caption = "Connect"
  GUI_frmWinsock.ConnectTimeOut.Enabled = False: GUI_frmWinsock.cmdConnect.Enabled = True
  SetTrayIcon SI_Offline
End Sub

'Automatically connect to next server in server list after a connect failure
Public Sub AutoConnect() ' : AddStack "Server_AutoConnect()"
  If JustJumped = True Then
    'If just jumped - reconnect in (i.e. 20) seconds
    ConnectServer JumpConnectDelay, ""
  Else
    'Otherwise - reconnect in (i.e. 5) seconds
    ConnectServer CurrentConnectDelay, ""
  End If
End Sub

'When the bot couldn't connect to a server since 15 tries -> restart
Public Sub CheckRestart() ' : AddStack "Server_CheckRestart()"
Dim u As Long, InUse As Boolean
  ConnectTryCounter = ConnectTryCounter + 1
  'Set connect delay according to number of connect failures
  If ConnectTryCounter >= 3 Then CurrentConnectDelay = FailureConnectDelay Else CurrentConnectDelay = DefaultConnectDelay
  
  If ConnectTryCounter = 15 Then
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If SocketItem(u).LastEvent + CDate("00:03:00") > Now Then InUse = True: ConnectTryCounter = 0: Exit For
      End If
    Next u
    If Not InUse Then
      PutLog "---> RESTARTING due to 15 connect failures... <----------------------------------"
      Status "*** Restarting due to 15 connect failures"
      TimedEvent "RESTART", 1
    End If
  End If
End Sub

Public Sub BufferEx() ' : AddStack "Server_BufferEx()"
Dim u As Long, u2 As Long, AnzBytes As Long, ChNum As Long
Dim SendThis As String, BufNum As Long, Kick(1 To 4) As Long, KickNick(1 To 4) As String
Dim KickNum As Long, DontAdd As Boolean, TempStr As String
  If BlockSends Or MaxBytesToServer - BytesSent <= 0 Then Exit Sub
  
  'Complete last line
  If LineToComplete <> "" Then
    SendThis = GetNextPart(LineToComplete)
    LineToComplete = GetNextPartRest
    If SendThis <> "" Then Output SendThis: SendTCP ServerSocket, SendThis: LastSendTime = Now
    If LineToComplete <> "" Then Exit Sub
  End If
  
  'Process ToBanList and KickList
  If (Buffer(1).BufferedLines = 0) And (Buffer(2).BufferedLines < 4) Then
    For ChNum = 1 To ChanCount
      If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then
        'Choose (max allowed) random bans from ToBanList (ban only if not +i)
        If (InStr(Param(Channels(ChNum).Mode, 1), "i") = 0) Or (Channels(ChNum).KickCount = 0) Then
          For u = IIf(Channels(ChNum).ToBanCount > ServerNumberOfModes, ServerNumberOfModes, Channels(ChNum).ToBanCount) To 1 Step -1
            TempStr = Channels(ChNum).ToBanList(u)
            DontAdd = False
            'Remove disturbing bans
            For u2 = 1 To Channels(ChNum).BanCount
              If LCase(TempStr) = LCase(Channels(ChNum).BanList(u2).Mask) Then
                DontAdd = True
              ElseIf MatchWM(Channels(ChNum).BanList(u2).Mask, TempStr) Then
                DontAdd = True
              ElseIf MatchWM(TempStr, Channels(ChNum).BanList(u2).Mask) Then
                WaitThisLine = True
                SendLine "mode " & Channels(ChNum).Name & " -b " & Channels(ChNum).BanList(u2).Mask, 2
              End If
            Next u2
            If DontAdd = False Then
              WaitThisLine = True
              SendLine "mode " & Channels(ChNum).Name & " +b " & TempStr, 2
              AddDesiredBan ChNum, TempStr
            End If
            RemToBan ChNum, TempStr
          Next u
        End If
        
        'Find users in KickList matching bans in the channel
        KickNum = 0
        For u2 = 1 To Channels(ChNum).BanCount
          For u = 1 To Channels(ChNum).KickCount
            If MatchWM(Channels(ChNum).BanList(u2).Mask, Channels(ChNum).KickList(u).Hostmask) Then
              KickNum = KickNum + 1: Kick(KickNum) = u
              If KickNum = 4 Then Exit For
            End If
          Next u
          If KickNum = 4 Then Exit For
        Next u2
        'Find users in KickList matching future bans in the channel
        If KickNum < 4 Then
          For u2 = 1 To Channels(ChNum).DesiredBanCount
            For u = 1 To Channels(ChNum).KickCount
              If MatchWM(Channels(ChNum).DesiredBanList(u2).Mask, Channels(ChNum).KickList(u).Hostmask) Then
                KickNum = KickNum + 1: Kick(KickNum) = u
                If KickNum = 4 Then Exit For
              End If
            Next u
            If KickNum = 4 Then Exit For
          Next u2
        End If
        'Find users in KickList not needing a ban
        If KickNum < 4 Then
          For u = 1 To Channels(ChNum).KickCount
            If Channels(ChNum).KickList(u).Hostmask = "" Then
              KickNum = KickNum + 1: Kick(KickNum) = u
              If KickNum = 4 Then Exit For
            End If
          Next u
        End If
        'Channel is +i: Kick the rest...
        If (KickNum < 4) And ((InStr(Param(Channels(ChNum).Mode, 1), "i") > 0) Or Channels(ChNum).InFlood) Then
          For u = 1 To Channels(ChNum).KickCount
            If Channels(ChNum).KickList(u).Hostmask <> "" Then
              DontAdd = False
              For u2 = 1 To KickNum
                If Kick(u2) = u Then DontAdd = True: Exit For
              Next u2
              If Not DontAdd Then
                KickNum = KickNum + 1: Kick(KickNum) = u
                If KickNum = 4 Then Exit For
              End If
            End If
          Next u
        End If
        For u = 1 To KickNum
          If Kick(u) > 0 Then
            If FindUser(Channels(ChNum).KickList(Kick(u)).Nick, ChNum) > 0 Then WaitThisLine = True: SendLine "kick " & Channels(ChNum).Name & " " & Channels(ChNum).KickList(Kick(u)).Nick & " :" & Channels(ChNum).KickList(Kick(u)).Message, 2
            KickNick(u) = Channels(ChNum).KickList(Kick(u)).Nick
          End If
        Next u
        For u = 1 To KickNum
          If Kick(u) > 0 Then RemKickUser ChNum, KickNick(u)
        Next u
      End If
    Next ChNum
  End If
  
  If (Buffer(1).BufferedLines = 0) And (Buffer(2).BufferedLines < 4) Then
    For ChNum = 1 To ChanCount
      If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then
        For u = IIf(Channels(ChNum).ToExceptCount > ServerNumberOfModes, ServerNumberOfModes, Channels(ChNum).ToExceptCount) To 1 Step -1
          TempStr = Channels(ChNum).ToExceptList(u)
          DontAdd = False
          'Remove disturbing excepts
          For u2 = 1 To Channels(ChNum).ExceptCount
            If LCase(TempStr) = LCase(Channels(ChNum).ExceptList(u2).Mask) Then
              DontAdd = True
            ElseIf MatchWM(Channels(ChNum).ExceptList(u2).Mask, TempStr) Then
              DontAdd = True
            ElseIf MatchWM(TempStr, Channels(ChNum).ExceptList(u2).Mask) Then
              WaitThisLine = True
              SendLine "mode " & Channels(ChNum).Name & " -e " & Channels(ChNum).ExceptList(u2).Mask, 2
            End If
          Next u2
          If DontAdd = False Then
            WaitThisLine = True
            SendLine "mode " & Channels(ChNum).Name & " +e " & TempStr, 2
            AddDesiredExcept ChNum, TempStr
          End If
          RemToExcept ChNum, TempStr
        Next u
      End If
    Next ChNum
  End If
  
  
  If (Buffer(1).BufferedLines = 0) And (Buffer(2).BufferedLines < 4) Then
    For ChNum = 1 To ChanCount
      If Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs Then
        For u = IIf(Channels(ChNum).ToInviteCount > ServerNumberOfModes, ServerNumberOfModes, Channels(ChNum).ToInviteCount) To 1 Step -1
          TempStr = Channels(ChNum).ToInviteList(u)
          DontAdd = False
          'Remove disturbing invites
          For u2 = 1 To Channels(ChNum).InviteCount
            If LCase(TempStr) = LCase(Channels(ChNum).InviteList(u2).Mask) Then
              DontAdd = True
            ElseIf MatchWM(Channels(ChNum).InviteList(u2).Mask, TempStr) Then
              DontAdd = True
            ElseIf MatchWM(TempStr, Channels(ChNum).InviteList(u2).Mask) Then
              WaitThisLine = True
              SendLine "mode " & Channels(ChNum).Name & " -I " & Channels(ChNum).InviteList(u2).Mask, 2
            End If
          Next u2
          If DontAdd = False Then
            WaitThisLine = True
            SendLine "mode " & Channels(ChNum).Name & " +I " & TempStr, 2
            AddDesiredInvite ChNum, TempStr
          End If
          RemToInvite ChNum, TempStr
        Next u
      End If
    Next ChNum
  End If
  
  'winsock2_send next line(s)
  For BufNum = 1 To 3
    Do
      If Buffer(BufNum).BufferedLines > 0 Then
        SendThis = GetNextPart(Buffer(BufNum).LineBuffer(1))
        LineToComplete = GetNextPartRest
        For u = 1 To Buffer(BufNum).BufferedLines - 1
          Buffer(BufNum).LineBuffer(u) = Buffer(BufNum).LineBuffer(u + 1)
        Next u
        Buffer(BufNum).BufferedLines = Buffer(BufNum).BufferedLines - 1
        If SendThis <> "" Then Output SendThis: SendTCP ServerSocket, SendThis: LastSendTime = Now
        If LineToComplete <> "" Then Output "->!": Exit Sub
        If LinesSent > 5 Then LinesSent = 0: SendTCP ServerSocket, "PING 1" & vbCrLf: BlockSends = True: Output "*** Blocked Sends" & vbCrLf: LastSendTime = Now: Exit Sub
      Else
        Exit Do
      End If
    Loop
  Next BufNum
End Sub

Public Sub SendLine(ByVal What As String, BufNum As Long) ' : AddStack "Server_SendLine(" & What & ", " & BufNum & ")"
Dim u As Long, u2 As Long, GotMode As String
  If Not Connected Then Exit Sub
  If CompressedModes(What, BufNum) Then Exit Sub
  If CompressedKicks(What, BufNum) Then Exit Sub
  If CompressedModes(What, BufNum) Then Exit Sub
  If CompressedKicks(What, BufNum) Then Exit Sub
  Buffer(BufNum).BufferedLines = Buffer(BufNum).BufferedLines + 1
  If Buffer(BufNum).BufferedLines > UBound(Buffer(BufNum).LineBuffer()) Then ReDim Preserve Buffer(BufNum).LineBuffer(UBound(Buffer(BufNum).LineBuffer()) + 5)
  Buffer(BufNum).LineBuffer(Buffer(BufNum).BufferedLines) = What & vbCrLf
  If WaitThisLine Then WaitThisLine = False: Exit Sub
  If Not (LCase(Param(What, 1) & " " & Param(What, 3)) = "mode -b") And Not (LCase(Param(What, 1) & " " & Param(What, 3)) = "mode +b") And Not (LCase(Param(What, 1)) = "kick") Then BufferEx
End Sub

Public Function GetNextPart(Line As String) As String ' : AddStack "Server_GetNextPart(" & Line & ")"
Dim ByteNum As Long
  If Line = "" Then Exit Function
  ByteNum = MaxBytesToServer - BytesSent
  If ByteNum >= Len(Line) Then
    GetNextPartRest = ""
    BytesSent = BytesSent + Len(Line)
    If Line <> "" Then LinesSent = LinesSent + 1
    GetNextPart = Line
  Else
    GetNextPartRest = Right(Line, Len(Line) - ByteNum)
    BytesSent = BytesSent + ByteNum
    GetNextPart = Left(Line, ByteNum)
  End If
End Function

Sub CloseIdentSockets()
  Dim u As Long
  If IdentSocket > 0 Then
    DisconnectSocket IdentSocket
    IdentSocket = -1
  End If
  RemoveTimedEvent "GetIdentSocket"
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If GetSockFlag(u, SF_Status) = SF_Status_Ident Then
        DisconnectSocket u
      End If
    End If
  Next u
End Sub

Sub DoAutoStuff(ChNum As Long) ' : AddStack "ServerRoutines_DoAutoStuff(" & ChNum & ")"
Dim PosInONick As Long, u2 As Long, OpUser As String, VoiceUser As String
Dim TargetNick As String, Multiple As String, MultCount As Long, OUserFlags As String
Dim DeopUser As String, RegUser As String, NoteCount As Long, KUserFlags As String
Dim GaveOps As Boolean, KickedUser As Boolean
  
  For u2 = 1 To Channels(ChNum).UserCount
    RegUser = Channels(ChNum).User(u2).RegNick
    GaveOps = False: KickedUser = False
    If (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) And (Channels(ChNum).User(u2).Nick <> MyNick) Then
      'Check if user is banned
      If CheckPermBans(ChNum, u2) Then
        If InStr(Channels(ChNum).User(u2).Status, "@") > 0 Then AddMassMode "-o", Channels(ChNum).User(u2).Nick
        KickedUser = True
      End If
    End If
    If RegUser <> "" Then
      If (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) And (Channels(ChNum).User(u2).Nick <> MyNick) Then
        If KickedUser = False Then
          OUserFlags = GetUserChanFlags2(Channels(ChNum).User(u2).UserNum, Channels(ChNum).Name)
          If MatchFlags(OUserFlags, "+k") Then
            If InStr(Channels(ChNum).User(u2).Status, "@") > 0 Then SendLine "mode " & Channels(ChNum).Name & " -o+b " & Channels(ChNum).User(u2).Nick & " " & Mask(Channels(ChNum).User(u2).Hostmask, 1), 2
            If InStr(Channels(ChNum).User(u2).Status, "@") = 0 Then SendLine "mode " & Channels(ChNum).Name & " +b " & Mask(Channels(ChNum).User(u2).Hostmask, 1), 2
            SendLine "kick " & Channels(ChNum).Name & " " & Channels(ChNum).User(u2).Nick & " :Banned: requested", 2
            KickedUser = True
          End If
        End If
        If KickedUser = False Then
          'AutoDeop
          If MatchFlags(OUserFlags, "+d") And Channels(ChNum).GotOPs Then
            If InStr(Channels(ChNum).User(u2).Status, "@") > 0 Then AddMassMode "-o", Channels(ChNum).User(u2).Nick
          'AutoOp
          ElseIf MatchFlags(OUserFlags, "+a") And Channels(ChNum).GotOPs Then
            If InStr(Channels(ChNum).User(u2).Status, "@") = 0 Then AddMassMode "+o", Channels(ChNum).User(u2).Nick: GaveOps = True
          End If
        End If
        If (GaveOps = False) And (KickedUser = False) Then
          'AutoVoice
          If MatchFlags(OUserFlags, "+v") And Channels(ChNum).User(u2).Status = "" Then
            If VoiceUser = "" Then VoiceUser = Channels(ChNum).User(u2).Nick Else VoiceUser = VoiceUser & " " & Channels(ChNum).User(u2).Nick
          End If
        End If
      End If
      If (Not IsIgnored(Channels(ChNum).User(u2).Hostmask)) And (KickedUser = False) Then
        NoteCount = NotesCount(Channels(ChNum).User(u2).RegNick)
        If NoteCount > 0 Then
          AddIgnore Mask(Channels(ChNum).User(u2).Hostmask, 2), 20, 1
          If NoteCount = 1 Then
            SendLine "notice " & Channels(ChNum).User(u2).Nick & " :Hi! I've got 1 note waiting for you.", 3
            SendLine "notice " & Channels(ChNum).User(u2).Nick & " :To get it, type: /msg " & MyNick & " notes <pass>", 3
          Else
            SendLine "notice " & Channels(ChNum).User(u2).Nick & " :Hi! I've got " & CStr(NoteCount) & " notes waiting for you.", 3
            SendLine "notice " & Channels(ChNum).User(u2).Nick & " :To get them, type: /msg " & MyNick & " notes <pass>", 3
          End If
        End If
      End If
    Else
      'Deop unknown users
      If (Channels(ChNum).GotOPs = True) And (Channels(ChNum).User(u2).Nick <> MyNick) Then
        If Channels(ChNum).DeopUnknownUsers > 0 Then
          If InStr(Channels(ChNum).User(u2).Status, "@") > 0 Then AddMassMode "-o", Channels(ChNum).User(u2).Nick
        End If
      End If
    End If
  Next u2
  If Not (Channels(ChNum).GotOPs Or Channels(ChNum).GotHOPs) Then Exit Sub
  DoMassMode Channels(ChNum).Name
  
  'Remove bans matching bots, masters, owners and superowners
  RemDisturbingBans ChNum
  RemDisturbingExcepts ChNum
  RemDisturbingInvites ChNum
  
  'Give voice to unvoiced and unopped +v users
  u2 = 0: MultCount = 0: Multiple = ""
  Do
    u2 = u2 + 1
    TargetNick = Param(VoiceUser, u2)
    If TargetNick = "" Then Exit Do
    MultCount = MultCount + 1
    If Multiple = "" Then Multiple = TargetNick Else Multiple = Multiple & " " & TargetNick
    If MultCount = ServerNumberOfModes Then SendLine "mode " & Channels(ChNum).Name & " +" & String(ServerNumberOfModes, "v") & " " & Multiple, 2: MultCount = 0: Multiple = """"
  Loop
  If MultCount > 0 Then SendLine "mode " & Channels(ChNum).Name & " +" & String(MultCount, "v") & " " & Multiple, 2
End Sub

Function GiveOp(Channel As String, Nick As String) As Boolean ' : AddStack "ServerRoutines_GiveOp(" & Channel & ", " & Nick & ")"
  If IsOrdered("giveop " & Channel & " " & Nick) = False Then
    Order "giveop " & Channel & " " & Nick, 10
    WaitThisLine = True
    SendLine "mode " & Channel & " +o " & Nick, 1
    GiveOp = True
  Else
    GiveOp = False
  End If
End Function

'Sets all variables according to my hostmask
Sub GotMyHost(TheHostmask As String) ' : AddStack "ServerRoutines_GotMyHost(" & TheHostmask & ")"
Dim Rest As String, HostOnly As String, ScNum As Long
  MyHostmask = TheHostmask
  SpreadFlagMessage 0, "+m", "3*** Connected as " & MyHostmask
  If IsValidIP(Mask(MyHostmask, 11)) Then
    MyIPmask = ""
  Else
    Rest = GetCacheIP(Mask(MyHostmask, 11), False)
    If Rest <> "" Then MyIPmask = Mask(MyHostmask, 14) & Rest Else MyIPmask = ""
  End If
  Status "Hostmask: " & MyHostmask + vbCrLf
  If MyIPmask <> "" Then Status "IP    mask: " & MyIPmask + vbCrLf
  
  Select Case ResolveIP
    Case "1"
        MyIP = IrcGetLongIp(GetLastLocalIP)
    Case "2"
        MyIP = IrcGetLongIp(GetLastLocalIP)
        HostOnly = Mask(MyHostmask, 11)
        HostOnly = IrcGetLongIp(GetCacheIP(HostOnly, True))
        If HostOnly <> MyIP Then
          If HostOnly <> "4294967295" Then MyIP = HostOnly
        End If
    Case Else
        MyIP = IrcGetLongIp(GetCacheIP(ResolveIP, True))
  End Select
    
  'Check script hooks
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.Srv_Connect Then
      RunScriptX ScNum, "Srv_Connect"
    End If
  Next ScNum
End Sub


'Gets bot users from hostmasks
Sub GetRegUsers(ChNum As Long) ' : AddStack "ServerRoutines_GetRegUsers(" & ChNum & ")"
Dim u As Long
  For u = 1 To Channels(ChNum).UserCount
    Channels(ChNum).User(u).RegNick = SearchUserFromHostmask(Channels(ChNum).User(u).Hostmask)
    Channels(ChNum).User(u).UserNum = BotUserNum
  Next u
End Sub

Public Sub FloodCheck(ChNum As Long, Modes As String, AllowCount As Long) ' : AddStack "ServerRoutines_FloodCheck(" & ChNum & ", " & Modes & ", " & AllowCount & ")"
  Channels(ChNum).FloodEvents = Channels(ChNum).FloodEvents + 1
  If (Channels(ChNum).InFlood = False) And (Channels(ChNum).FloodEvents >= AllowCount) Then
    SpreadFlagMessage 0, "+m", MakeMsg(MSG_PLFloodStart, Channels(ChNum).Name, Modes)
    SendLine "mode " & Channels(ChNum).Name & " " & Modes, 2
    Channels(ChNum).InFlood = True
    TimedEvent "checkforflood " & Channels(ChNum).Name, 60
  End If
End Sub


Public Sub InitiateBotChat(ByVal UserNum As Long, AutoLink As Boolean) ' : AddStack "Routines_InitiateBotChat(" & UserNum & ", " & AutoLink & ")"
  Dim Host As String, Port As String, strRemoteIP As String, NewSock As Long
  Dim FullLine As String, Result As Long
  FullLine = GetUserData(UserNum, UD_LinkAddr, "")
  If InStr(FullLine, ":") > 0 Then
    Host = ParamX(FullLine, ":", 1)
    Port = ParamX(FullLine, ":", 2)
    If Not IsNumeric(Port) Then
      If Not AutoLink Then SpreadFlagMessage 0, "+t", "14[" & Time & "] *** Couldn't link to " & BotUsers(UserNum).Name & ": Port isn't numeric!"
      Exit Sub
    End If
  Else
    Exit Sub
  End If
  
  NewSock = AddSocket
  Result = ConnectTCP(NewSock, Host, CLng(Port))
  If Result = 0 Then
    SocketItem(NewSock).Hostmask = Host
    SocketItem(NewSock).RegNick = BotUsers(UserNum).Name
    SocketItem(NewSock).IRCNick = ""
    SocketItem(NewSock).Flags = BotUsers(UserNum).Flags
    SocketItem(NewSock).UserNum = UserNum
    SetSockFlag NewSock, SF_Colors, SF_YES
    SetSockFlag NewSock, SF_Echo, SF_NO
    SetSockFlag NewSock, SF_Status, SF_Status_InitBotLink
    SetSockFlag NewSock, SF_LoggedIn, SF_NO
    SocketItem(NewSock).OnBot = BotNetNick
    SocketItem(NewSock).OrderSign = CStr(Timer)
    SocketItem(NewSock).LinkStatus = IIf(AutoLink, "AutoLink", "")
    SocketItem(NewSock).PLChannel = 0
    SocketItem(NewSock).IsInternalSocket = True
  Else
    RemoveSocket NewSock, 0, "", True
    If Not AutoLink Then SpreadFlagMessage 0, "+t", "14[" & Time & "] *** Couldn't link to " & BotUsers(UserNum).Name & ": " & WSAGetErrorString(Result)
  End If
End Sub

Public Sub InitiateDCCChat(Nick As String, Hostmask As String, RegUser As String, UserFlags As String) ' : AddStack "Routines_InitiateDCCChat(" & Nick & ", " & Hostmask & ", " & RegUser & ", " & UserFlags & ")"
  Dim Port As Long, NewSock As Long, Index As Integer
  On Local Error Resume Next
  
  NewSock = AddSocket
  Index = 0
  Do
    Index = Index + 1
    Port = WSARandomPort
    If ListenTCP(NewSock, Port) = 0 Then Exit Do
    If Index >= 20 Then
      RemoveSocket NewSock, 0, "", True
      Exit Sub
    End If
  Loop
  
  'Router Fix
  Dim HostOnly As String
  HostOnly = Mask(Hostmask, 11)
  HostOnly = IrcGetLongIp(GetCacheIP(HostOnly, True))
  If HostOnly = MyIP And RouterWorkAround = True Then
    If LocalAddress = "" Then
      HostOnly = IrcGetLongIp(GetLastLocalIP())
    Else
      If IsValidIP(LocalAddress) Then
        HostOnly = IrcGetLongIp(LocalAddress)
      Else
        HostOnly = MyIP
      End If
    End If
  Else
    HostOnly = MyIP
  End If
  
  SocketItem(NewSock).Hostmask = Hostmask
  SocketItem(NewSock).RegNick = RegUser
  SocketItem(NewSock).IRCNick = Nick
  SocketItem(NewSock).Flags = UserFlags
  SocketItem(NewSock).UserNum = GetUserNum(RegUser)
  SetSockFlag NewSock, SF_Status, SF_Status_DCCInit
  SetSockFlag NewSock, SF_Colors, GetUserData(SocketItem(NewSock).UserNum, "colors", SF_NO)
  SetSockFlag NewSock, SF_Echo, SF_NO
  SetSockFlag NewSock, SF_DCC, SF_YES
  SetSockFlag NewSock, SF_LF_ONLY, SF_YES
  SocketItem(NewSock).OnBot = BotNetNick
  SocketItem(NewSock).PLChannel = BotUsers(SocketItem(NewSock).UserNum).PLChannel
  SocketItem(NewSock).IsInternalSocket = True
  SendLine "PRIVMSG " & Nick & " :DCC CHAT chat " & HostOnly & " " & CStr(Port) & "", 1
End Sub

Public Function InitiateDCCSend(Nick As String, RegUser As String, FileName As String) As Boolean ' : AddStack "Routines_InitiateDCCSend(" & Nick & ", " & RegUser & ", " & FileName & ")"
  Dim Port As Long, NewSock As Long, FileNum As Integer, FileLength As Long, ScNum As Long, TP As Long, Index As Long
  On Local Error Resume Next
  FileNum = FreeFile: Open FileName For Input Shared As #FileNum: FileLength = LOF(FileNum): Close #FileNum
  If Err.Number > 0 Then InitiateDCCSend = False: Exit Function
  HaltDefault = False
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.fa_downloadbegin Then
      RunScriptX ScNum, "fa_downloadbegin", Nick, RegUser, FileName, FileLength
    End If
  Next ScNum
  If HaltDefault = True Then
    InitiateDCCSend = False: Exit Function
  End If
  
  If Err.Number > 0 Then Err.Clear
  On Error GoTo 0
  
  'Port Range
  Index = 0
  NewSock = AddSocket
  Do
    Index = Index + 1
    Port = WSARandomPort
    If ListenTCP(NewSock, Port) = 0 Then Exit Do
    If Index >= 20 Then
      InitiateDCCSend = False
      Exit Function
    End If
  Loop
  
  'Router Fix
  Dim HostOnly As String
  For Index = 1 To ChanCount
    If FindUser(Nick, Index) > 0 Then
      HostOnly = Mask(Channels(Index).User(FindUser(Nick, Index)).Hostmask, 11)
      Exit For
    End If
  Next Index
  If HostOnly = "" Then HostOnly = "localhost"
  HostOnly = IrcGetLongIp(GetCacheIP(HostOnly, True))
  If HostOnly = MyIP And RouterWorkAround = True Then
    If LocalAddress = "" Then
      HostOnly = IrcGetLongIp(GetLastLocalIP())
    Else
      If IsValidIP(LocalAddress) Then
        HostOnly = IrcGetLongIp(LocalAddress)
      Else
        HostOnly = MyIP
      End If
    End If
  Else
    HostOnly = MyIP
  End If
  
  SocketItem(NewSock).Hostmask = ""
  SocketItem(NewSock).RegNick = RegUser
  SocketItem(NewSock).IRCNick = Nick
  SocketItem(NewSock).Flags = ""
  SocketItem(NewSock).FileSize = FileLength
  SocketItem(NewSock).FileName = FileName
  SocketItem(NewSock).TrafficIn = 0
  SocketItem(NewSock).UserNum = GetUserNum(RegUser)
  SetSockFlag NewSock, SF_Status, SF_Status_SendFileWaiting
  SetSockFlag NewSock, SF_Echo, SF_NO
  SocketItem(NewSock).OnBot = BotNetNick
  SocketItem(NewSock).IsInternalSocket = True
  SendLine "PRIVMSG " & Nick & " :DCC SEND " & MakeUnderscore(GetFileName(FileName)) & " " & HostOnly & " " & CStr(Port) & " " & CStr(FileLength) & "", 1
  InitiateDCCSend = True
End Function

