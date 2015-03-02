Attribute VB_Name = "Plugin_Seen"
',-======================- ==-- -  -
'|   AnGeL - Plugins - Seen
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Type SeenType
  Nick As String
  SeenTime As String
End Type

Public Const MaxExtSeenCount As Long = 5000

Public LastSeenOutput As String


Public SeenCacheCount As Long
Public SeenCache() As String
Public ExtSeenCacheCount As Long
Public ExtSeenCache() As String


Sub Seen_Load()
  ReDim Preserve SeenCache(5)
  ReDim Preserve ExtSeenCache(5)
End Sub


Sub Seen_Unload()
'
End Sub


Public Function LastSeen(ByVal SeenNick As String, Nick As String, Chan As String, RegNick As String, MatchedOne As Boolean) As String ' : AddStack "SeenRoutines_LastSeen(" & SeenNick & ", " & Nick & ", " & Chan & ", " & RegNick & ", " & MatchedOne & ")"
  Dim ChNum As Long, UsNum As Long, u As Long, u2 As Long, Alias As String, IsHostmask As Boolean
  Dim BotUser As Long, Message As String, Hostmask As String, FileNum As Integer, SeenLine As String
  Dim Rest As String, RegUser As String, Average As Boolean, Reply As String, SeenChan As String
  Dim ExtRest As String, ExtRegUser As String, ExtAverage As Boolean, UnReply As String, ExtSeenChan As String
  Dim SeenSort(15) As SeenType, SeenCount As Long, GotMatch As Boolean, SeenString As String
  Dim UseFile As Boolean, TempStr As String

  UsNum = GetUserNum(SeenNick)
  If UsNum = 0 Then
    On Local Error Resume Next
    UseFile = True
    If InStr(SeenNick, ",") > 0 Then SeenLine = Replace(SeenNick, ",", " "): UseFile = False
    If UseFile And Dir(FileAreaHome & GetFileName(SeenNick) & ".seen") <> "" Then
      FileNum = FreeFile
      Open FileAreaHome & GetFileName(SeenNick) & ".seen" For Input As #FileNum
    End If
    If Err.Number = 0 And Dir(FileAreaHome & GetFileName(SeenNick) & ".seen") <> "" Then
      If UseFile Then
        If LOF(FileNum) > 200 Then Close #FileNum: LastSeen = Nick & ": Sorry, there's an error in the seen list - file is too long.": Exit Function
        Err.Clear
        MatchedOne = True
        Line Input #FileNum, SeenLine
        Close #FileNum
      End If
      For u = 1 To IIf(ParamCount(SeenLine) > 14, 14, ParamCount(SeenLine))
        Alias = Param(SeenLine, u)
        If (IsValidNick(Alias) = True) And (Len(Alias) <= ServerNickLen) Then
          GotMatch = False
          For ChNum = 1 To ChanCount
            For u2 = 1 To Channels(ChNum).UserCount
              If (LCase(Channels(ChNum).User(u2).RegNick) = LCase(Alias)) Or (LCase(Channels(ChNum).User(u2).Nick) = LCase(Alias)) Then
                GotMatch = True: If (LCase(Channels(ChNum).User(u2).RegNick) = LCase(Alias)) Then RegUser = Channels(ChNum).User(u2).RegNick Else RegUser = Channels(ChNum).User(u2).Nick
                Exit For
              End If
            Next u2
            If GotMatch Then Exit For
          Next ChNum
          If GotMatch Then
            SeenCount = SeenCount + 1
            SeenSort(SeenCount).Nick = RegUser
            SeenSort(SeenCount).SeenTime = "online!"
          Else
            Reply = SeenReply("Seen.txt", Alias, RegUser, SeenChan, Rest, False, False, Average)
            If Reply = "" Then Reply = SeenReply("ExtSeen.txt", Alias, RegUser, SeenChan, Rest, False, False, Average)
            SeenCount = SeenCount + 1
            If Reply <> "" Then
              SeenSort(SeenCount).Nick = RegUser
              SeenSort(SeenCount).SeenTime = Rest
            Else
              SeenSort(SeenCount).Nick = Param(SeenLine, u)
              SeenSort(SeenCount).SeenTime = ""
            End If
          End If
        End If
      Next u
      For u = 1 To SeenCount - 1
        For u2 = u + 1 To SeenCount
          If A_after_B(SeenSort(u2).SeenTime, SeenSort(u).SeenTime) Then
            SeenSort(15) = SeenSort(u): SeenSort(u) = SeenSort(u2): SeenSort(u2) = SeenSort(15)
          End If
        Next u2
      Next u
      For u = 1 To SeenCount
        Select Case SeenSort(u).SeenTime
          Case "":        SeenString = "not seen"
          Case "online!": SeenString = "online!"
          Case Else:      SeenString = TimeSpan3(SeenSort(u).SeenTime)
        End Select
        If LastSeen = "" Then
          LastSeen = Nick & ": " & SeenSort(u).Nick & " (" & SeenString & ")"
        Else
          LastSeen = LastSeen & ", " & SeenSort(u).Nick & " (" & SeenString & ")"
        End If
        MatchedOne = True
      Next u
      If SeenCount = 0 Then
        MatchedOne = False
        If UseFile Then
          LastSeen = Nick & ": Sorry, there's an error in the seen list."
        Else
          LastSeen = ""
        End If
      End If
      Exit Function
    Else
      Close #FileNum
    End If
    If Err.Number > 0 Then Err.Clear
    On Error GoTo 0
  End If
  
  If Chan <> "" Then ChNum = FindChan(Chan)
  IsHostmask = IsValidHostmask(SeenNick)
  If Not IsHostmask Then
    If (InStr(SeenNick, "?") > 0) Or (InStr(SeenNick, "*") > 0) Then
      If (InStr(SeenNick, "!") = 0) And (InStr(SeenNick, "@") = 0) Then
        If InStr(SeenNick, ".") > 0 Then
          SeenNick = "*!*@" & SeenNick
        Else
          SeenNick = SeenNick & "!*@*"
        End If
      ElseIf (InStr(SeenNick, "@") > 0) Then
        SeenNick = "*!" & SeenNick
      ElseIf (InStr(SeenNick, "!") > 0) Then
        SeenNick = SeenNick & "@*"
      End If
    End If
    IsHostmask = IsValidHostmask(SeenNick)
  End If
  MatchedOne = True
  If Not IsHostmask Then
    If (Len(SeenNick) > ServerNickLen) Or (IsValidNick(SeenNick) = False) Then
      LastSeen = ""
      MatchedOne = False
      Exit Function
    End If
    If LCase(SeenNick) = LCase(Nick) Then LastSeen = Nick & ": Yeah, I can see you!": Exit Function
    If LCase(SeenNick) = LCase(MyNick) Then LastSeen = Nick & ": Here I am!": Exit Function
  Else
    If MatchWM(SeenNick, MyHostmask) Then LastSeen = Nick & ": Here I am! I'm matching this hostmask! :)": Exit Function
  End If
  
  If Not IsHostmask Then
    If Chan <> "" Then
      'Find RegNicks
      For u = 1 To Channels(ChNum).UserCount
        If LCase(Channels(ChNum).User(u).RegNick) = LCase(SeenNick) Then
          If Channels(ChNum).User(u).Nick <> Nick Then
            'User I'm being asked for is on the same channel as the asking user
            If LCase(Channels(ChNum).User(u).Nick) <> LCase(Channels(ChNum).User(u).RegNick) Then
              LastSeen = Channels(ChNum).User(u).RegNick & " is " & Channels(ChNum).User(u).Nick & ", and " & Channels(ChNum).User(u).Nick & " is on the channel right now!"
              Exit Function
            End If
          Else
            'User I'm being asked for is the asking user :)
            Select Case Int(Rnd * 3) + 1
              Case 1: LastSeen = Nick & ": Identity crisis? YOU are " & Channels(ChNum).User(u).RegNick & "!! :)"
              Case 2: LastSeen = Nick & ": I think that *you* are " & Channels(ChNum).User(u).RegNick & "!"
              Case 3: LastSeen = Nick & ": Well, you should know the answer - YOU are " & Channels(ChNum).User(u).RegNick & "! ;-)"
            End Select
            Exit Function
          End If
        End If
      Next u
      'Find normal Nicks
      For u = 1 To Channels(ChNum).UserCount
        'User I'm being asked for is on the same channel
        If LCase(Channels(ChNum).User(u).Nick) = LCase(SeenNick) Then
          Select Case Int(Rnd * 4) + 1
            Case 1: LastSeen = Channels(ChNum).User(u).Nick & " is on the channel right now!"
            Case 2: LastSeen = Nick & ": Go check your eyes! " & Channels(ChNum).User(u).Nick & " is here right now! :)"
            Case 3: LastSeen = Nick & ": Open your eyes... " & Channels(ChNum).User(u).Nick & " is on the channel right now!"
            Case 4: LastSeen = Nick & ": You blind camel! " & Channels(ChNum).User(u).Nick & " is on the channel right now! :o)"
          End Select
          Exit Function
        End If
      Next u
    End If
    'User I'm being asked for is online, but on another channel
    For ChNum = 1 To ChanCount
      If LCase(Channels(ChNum).Name) <> LCase(Chan) Then
        'Look for RegNicks
        For u = 1 To Channels(ChNum).UserCount
          If (LCase(Channels(ChNum).User(u).RegNick) = LCase(SeenNick)) Then
            If Not (Channels(ChNum).Secret And (MatchFlags(GetUserChanFlags(RegNick, Channels(ChNum).Name), "-n") Or Chan <> "")) Then
              If LCase(Channels(ChNum).User(u).RegNick) <> LCase(Channels(ChNum).User(u).Nick) Then
                LastSeen = Channels(ChNum).User(u).RegNick & " is " & Channels(ChNum).User(u).Nick & ", and " & Channels(ChNum).User(u).Nick & " is on IRC channel " & Channels(ChNum).Name & " right now!"
              Else
                LastSeen = Channels(ChNum).User(u).RegNick & " is on IRC channel " & Channels(ChNum).Name & " right now!"
              End If
              Exit Function
            Else
              If LCase(Channels(ChNum).User(u).RegNick) <> LCase(Channels(ChNum).User(u).Nick) Then
                LastSeen = Channels(ChNum).User(u).RegNick & " is " & Channels(ChNum).User(u).Nick & ", and " & Channels(ChNum).User(u).Nick & " is on IRC right now!"
              Else
                LastSeen = Channels(ChNum).User(u).RegNick & " is on IRC right now!"
              End If
              Exit Function
            End If
          End If
        Next u
        'Look for normal Nicks
        For u = 1 To Channels(ChNum).UserCount
          If (LCase(Channels(ChNum).User(u).Nick) = LCase(SeenNick)) Then
            If Not (Channels(ChNum).Secret And (MatchFlags(GetUserChanFlags(RegNick, Channels(ChNum).Name), "-n") Or Chan <> "")) Then
              LastSeen = Channels(ChNum).User(u).Nick & " is on IRC channel " & Channels(ChNum).Name & " right now!"
            Else
              LastSeen = Channels(ChNum).User(u).Nick & " is on IRC right now!"
            End If
            Exit Function
          End If
        Next u
      End If
    Next ChNum
    'Party line
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If LCase(SocketItem(u).RegNick) = LCase(SeenNick) Then
          If (SocketItem(u).OnBot = BotNetNick) And (GetSockFlag(u, SF_LocalVisibleUser) = SF_YES) And (SocketItem(u).AwayMessage = "") Then
            LastSeen = SocketItem(u).RegNick & " is on my party line right now!"
            Exit Function
          End If
        End If
      End If
    Next u
    For u = 1 To SocketCount
      If IsValidSocket(u) Then
        If LCase(SocketItem(u).RegNick) = LCase(SeenNick) Then
          If (GetSockFlag(u, SF_Status) = SF_Status_BotNetParty) And (SocketItem(u).AwayMessage = "") Then
            LastSeen = SocketItem(u).RegNick & " is on " & SocketItem(u).OnBot & "'s party line right now!"
            Exit Function
          End If
        End If
      End If
    Next u
  Else
    If Chan <> "" Then
      'Find Hostmasks in local channel
      For u = 1 To Channels(ChNum).UserCount
        If MatchWM(SeenNick, Channels(ChNum).User(u).Hostmask) Then
          If (LCase(Channels(ChNum).User(u).Nick) <> LCase(Nick)) And (LCase(Channels(ChNum).User(u).RegNick) <> LCase(Nick)) Then
            LastSeen = Channels(ChNum).User(u).Nick & " (" & Mask(Channels(ChNum).User(u).Hostmask, 10) & ") is on the channel right now!"
            Exit Function
          Else
            'User I'm being asked for is the asking user :)
            Select Case Int(Rnd * 3) + 1
              Case 1: LastSeen = Nick & ": Identity crisis? YOU are matching this!! :)"
              Case 2: LastSeen = Nick & ": I think that *you* are matching this hostmask!"
              Case 3: LastSeen = Nick & ": Hey, YOU are the one you're searching for! ;-)"
            End Select
            Exit Function
          End If
        End If
      Next u
    End If
    'Find Hostmasks in other channels
    For ChNum = 1 To ChanCount
      If LCase(Channels(ChNum).Name) <> LCase(Chan) Then
        For u = 1 To Channels(ChNum).UserCount
          If MatchWM(SeenNick, Channels(ChNum).User(u).Hostmask) Then
            If (LCase(Channels(ChNum).User(u).Nick) <> LCase(Nick)) And (LCase(Channels(ChNum).User(u).RegNick) <> LCase(Nick)) Then
              If Not (Channels(ChNum).Secret And (MatchFlags(GetUserChanFlags(RegNick, Channels(ChNum).Name), "-n") Or Chan <> "")) Then
                LastSeen = Channels(ChNum).User(u).Nick & " (" & Mask(Channels(ChNum).User(u).Hostmask, 10) & ") is on IRC channel " & Channels(ChNum).Name & " right now!"
              Else
                LastSeen = Channels(ChNum).User(u).Nick & " (" & Mask(Channels(ChNum).User(u).Hostmask, 10) & ") is on IRC right now!"
              End If
              Exit Function
            Else
              'User I'm being asked for is the asking user :)
              Select Case Int(Rnd * 3) + 1
                Case 1: LastSeen = Nick & ": Identity crisis? YOU are matching this!! :)"
                Case 2: LastSeen = Nick & ": I think that *you* are matching this hostmask!"
                Case 3: LastSeen = Nick & ": Hey, YOU are the one you're searching for! ;-)"
              End Select
              Exit Function
            End If
          End If
        Next u
      End If
    Next ChNum
  End If
  
  'Give information about the requesting user to the procedures
  RegUser = RegNick: SeenChan = Chan: ExtRegUser = RegNick: ExtSeenChan = Chan
  'Read normal seen info
  Reply = SeenReply("Seen.txt", SeenNick, RegUser, SeenChan, Rest, IsHostmask, True, Average)
  'Read extended seen info (for unregistered users)
  UnReply = SeenReply("ExtSeen.txt", SeenNick, ExtRegUser, ExtSeenChan, ExtRest, True, True, ExtAverage)
  
  'Registered seen returned exact result and unregistered seen returned only a best match? Delete it!
  If (Average = False) And (Reply <> "") And (ExtAverage = True) Then UnReply = ""
  
  'Registered seen returned exact result and unregistered seen only an alias match? Delete it!
  If (Average = False) And (ExtAverage = False) And InStr(ExtRegUser, "alias") > 0 Then UnReply = ""
  
  If Not IsHostmask Then
    If (Average = True) Or (Reply = "") Then
      If InStr(ExtRegUser, "alias") > 0 Then
        TempStr = SeenReply("Seen.txt", Param(ExtRegUser, 1), RegUser, SeenChan, Rest, IsHostmask, False, Average)
        If MatchGrade(Param(ExtRegUser, 1), Param(ExtRegUser, 3)) < 35 Then TempStr = ""
        If TempStr <> "" Then
          If ExtAverage Then
            If Param(ExtRegUser, 1) = RegNick Then
              Select Case Int(Rnd * 3) + 1
                Case 1: LastSeen = Nick & ": I don't know anybody by that name. You were once using the best matching nick: " & Param(ExtRegUser, 3) & " :)"
                Case 2: LastSeen = Nick & ": The best matching person are YOU as '" & Param(ExtRegUser, 3) & "'! ;)"
                Case 3: LastSeen = Nick & ": No exact matches found, but you're the best match (with your nick '" & Param(ExtRegUser, 3) & "')! ;-)"
              End Select
            Else
              If Reply <> "" Then
                LastSeen = "Best match: " & Reply
              Else
                LastSeen = "Hmm. Do you mean " & Param(ExtRegUser, 3) & "? " & Param(ExtRegUser, 3) & " is " & Param(ExtRegUser, 1) & ", and " & TempStr
              End If
            End If
          Else
            If Param(ExtRegUser, 1) = RegNick Then
              'User I'm being asked for is the asking user :)
              Select Case Int(Rnd * 3) + 1
                Case 1: LastSeen = Nick & ": Identity crisis? YOU are " & Param(ExtRegUser, 3) & "!! :)"
                Case 2: LastSeen = Nick & ": I think that *you* are " & Param(ExtRegUser, 3) & "!"
                Case 3: LastSeen = Nick & ": Well, you should know the answer - YOU are " & Param(ExtRegUser, 3) & "! ;-)"
              End Select
            Else
              LastSeen = "I think that " & Param(ExtRegUser, 3) & " is " & Param(ExtRegUser, 1) & ". " & TempStr
            End If
          End If
          Exit Function
        End If
      End If
    End If
  End If
  
  If ((Average = True) Or (Reply = "")) And ((ExtAverage = True) Or (UnReply = "")) Then
    'Only best matches were found; react to keywords
    Select Case LCase(SeenNick)
      Case "": LastSeen = Nick & ": Seen whom?!": Exit Function
      Case "me", "myself": LastSeen = Nick & ": Yeah, I can see you!": Exit Function
      Case "you", "yourself": LastSeen = Nick & ": Every time I look in a mirror...": Exit Function
      Case "god": LastSeen = Nick & ": God? Here I am! ;)": Exit Function
      Case "your": LastSeen = Nick & ": Let's not get personal, okie? *^^*": Exit Function
      Case "my": LastSeen = Nick & ": I don't care.": Exit Function
    End Select
  End If
  If Reply = "" Then
    RegUser = GetRealNick(SeenNick)
    If UnReply <> "" Then
      If RegUser <> "" Then
        If InStr(ExtRegUser, "alias") > 0 Then
          LastSeen = "I've never seen my registered user " & RegUser & " around, but " & UnReply
        Else
          LastSeen = "I've never seen my registered user " & RegUser & " around, but the unregistered user " & UnReply
        End If
      Else
        LastSeen = IIf(ExtAverage, "Best match: ", "") + UnReply
      End If
      Exit Function
    End If
    'Nothing found at all; return 'I don't know who xyz is'
    MatchedOne = False
    If RegUser <> "" Then
      LastSeen = "I've never seen my user " & RegUser & " around."
    Else
      If IsHostmask Then
        LastSeen = Choose(Int(Rnd * 4) + 1, "Sorry, no matches for this hostmask.", "Sorry, no matches found.", "No matches.", "I never saw somebody matching this.")
      Else
        LastSeen = Choose(Int(Rnd * 4) + 1, "I don't know who " & SeenNick & " is.", "Sorry, I don't know " & SeenNick & ".", "I never saw somebody called " & SeenNick & ".", SeenNick & "? Sorry, I don't know this user.")
      End If
    End If
    Exit Function
  Else
    RegUser = GetRealNick(SeenNick)
    LastSeen = Reply
    'Add extended seen info (if necessary)
    If UnReply <> "" Then
      If (Average = True) And (ExtAverage = False) Then
        LastSeen = UnReply
      Else
        If A_after_B(ExtRest, Rest) Then
          If InStr(ExtRegUser, "alias") > 0 Then
            '"Hella alias Hippo" ... "Hippo was..."    bei Suche nach "Hippo" -> "The real" rechts hinzufügen
            If Param(ExtRegUser, 3) = RegUser Then LastSeen = UnReply & " The real " & Reply Else LastSeen = UnReply
            If ExtAverage Then LastSeen = "Best match: " & LastSeen
          Else
            If LCase(Param(UnReply, 1)) = LCase(Param(Reply, 1)) Then
              LastSeen = "Somebody called '" & Param(UnReply, 1) & "' " & GetRest(UnReply, 2) & " My registered user " & Reply
            Else
              LastSeen = "The unregistered user " & UnReply & " My registered user " & Reply
            End If
            If Average Or ExtAverage Then LastSeen = "Best matches: " & LastSeen
          End If
        Else
          If Average Then LastSeen = "Best match: " & LastSeen
        End If
      End If
    Else
      If Average Then LastSeen = "Best match: " & LastSeen
    End If
  End If
End Function

Public Function SeenReply(FileName As String, SeenNick As String, RegNick As String, Chan As String, ExtRest As String, ShowHostmask As Boolean, EnableAverage As Boolean, Average As Boolean) As String ' : AddStack "SeenRoutines_SeenReply(" & FileName & ", " & SeenNick & ", " & RegNick & ", " & Chan & ", " & ExtRest & ", " & ShowHostmask & ", " & EnableAverage & ", " & Average & ")"
Dim ChNum As Long, UsNum As Long, RegUser As String, Rest As String, u As Long, Alias As String
Dim SeenChan As String, BotUser As Long, Message As String, SecretMessage As String, Hostmask As String
Dim FullUserName As String, IsHostmask As Boolean
  IsHostmask = IsValidHostmask(SeenNick)
  RegUser = RegNick
  If (Len(SeenNick) <= ServerNickLen And IsValidNick(SeenNick)) Or (IsHostmask = True) Then
    ReadSeenEntry FileName, SeenNick, RegNick, Alias, Rest, SeenChan, Message, Hostmask, EnableAverage, Average
    ExtRest = Rest
    If RegNick = "" Then SeenReply = "": Exit Function
    If Alias <> "-" Then RegNick = Alias & " alias " & RegNick
    FullUserName = RegNick + IIf(ShowHostmask, " (" & Hostmask & ")", "")
    
    If (Channels(FindChan(SeenChan)).Secret = True) And (((Chan <> "") And (LCase(SeenChan) <> LCase(Chan))) Or ((Chan = "") And MatchFlags(GetUserChanFlags(RegUser, SeenChan), "-n"))) Then
      SeenReply = FullUserName & " was last on IRC " & TimeSpan(Rest): Exit Function
    Else
      Select Case Message
        Case "Connection reset by peer"
          SeenReply = FullUserName & " has quit " & SeenChan & " with 'Connection reset by peer' " & TimeSpan(Rest)
        Case "Ping timeout"
          SeenReply = FullUserName & " had a ping timeout in " & SeenChan & " " & TimeSpan(Rest)
        Case "*left*"
          SeenReply = FullUserName & " has left " & SeenChan & " " & TimeSpan(Rest)
        Case "*join*"
          SeenReply = FullUserName & " has joined " & SeenChan & " " & TimeSpan(Rest)
        Case "*partyline*"
          If SeenChan = "*mine*" Then
            SeenReply = FullUserName & " was on my party line " & TimeSpan(Rest)
          Else
            SeenReply = FullUserName & " was on " & SeenChan & "'s party line " & TimeSpan(Rest)
          End If
        Case Else
          If Param(Message, 1) = "*kicked*" Then
            SeenReply = FullUserName & " was kicked off " & SeenChan & " by " & Param(Message, 2) & " " & TimeSpan(Rest)
          Else
            If Chan = "" Then
              If Message <> "" Then
                SeenReply = FullUserName & " quit IRC channel " & SeenChan & " (saying '" & Strip(Message) & "') " & TimeSpan(Rest)
              Else
                SeenReply = FullUserName & " was last on IRC channel " & SeenChan & " " & TimeSpan(Rest)
              End If
            Else
              If LCase(Chan) = LCase(SeenChan) Then
                SeenReply = FullUserName & " was last on this channel " & TimeSpan(Rest)
              Else
                SeenReply = FullUserName & " was last on IRC channel " & SeenChan & " " & TimeSpan(Rest)
              End If
            End If
          End If
      End Select
    End If
  Else
    SeenReply = ""
  End If
End Function

Sub ReadSeenEntry(ByVal FileName As String, SearchFor As String, Nick As String, Alias As String, LastSeen As String, SeenChan As String, Message As String, Hostmask As String, EnableAverage As Boolean, Average As Boolean) ' : AddStack "SeenRoutines_ReadSeenEntry(" & FileName & ", " & SearchFor & ", " & Nick & ", " & Alias & ", " & LastSeen & ", " & SeenChan & ", " & Message & ", " & Hostmask & ", " & EnableAverage & ", " & Average & ")"
Dim u As Long, CurLine As String, FileNumber As Integer, Rest As String, HGrade As Long
Dim IsHostmask As Boolean, IsMatching As Boolean, Nicks As String, CurIdx As Long
Dim FileLines() As String, LineCount As Long
On Error GoTo RESEErr
  IsHostmask = IsValidHostmask(SearchFor)
  On Local Error Resume Next
  ReDim FileLines(100)
  
  'Add new lines first
  If FileName = "ExtSeen.txt" Then
    For u = ExtSeenCacheCount To 1 Step -1
      LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
      FileLines(LineCount) = ExtSeenCache(u)
      Nicks = Nicks & " " & Param(ExtSeenCache(u), 1)
    Next u
  Else
    For u = SeenCacheCount To 1 Step -1
      LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
      FileLines(LineCount) = SeenCache(u)
      Nicks = Nicks & " " & Param(SeenCache(u), 1)
    Next u
  End If
  Nicks = " " & LCase(Trim(Nicks)) & " "
  
  'Add the rest
  FileName = HomeDir & FileName
  FileNumber = FreeFile
  If Dir(FileName) <> "" Then
    WaitForAccess FileName
    AddAccessedFile FileName
    On Error GoTo RESEErr
    Open FileName For Input As #FileNumber
      Do While Not EOF(FileNumber)
        Line Input #FileNumber, CurLine
        If Trim(CurLine) <> "" Then
          If InStr(Nicks, " " & LCase(Param(CurLine, 1)) & " ") = 0 Then
            LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
            FileLines(LineCount) = CurLine
          End If
        End If
      Loop
    Close #FileNumber
    RemAccessedFile FileName
  End If
  
  'First try, get exact matches
  For CurIdx = 1 To LineCount
    CurLine = FileLines(CurIdx)
    If IsHostmask Then
      IsMatching = MatchWM(SearchFor, Param(CurLine, 1) & "!" & Param(CurLine, 6))
    Else
      IsMatching = (LCase(Param(CurLine, 1)) = LCase(SearchFor))
    End If
    If IsMatching Then
      Nick = Param(CurLine, 1)
      Alias = Param(CurLine, 2)
      LastSeen = AddSpaces(Param(CurLine, 3))
      SeenChan = Param(CurLine, 4)
      Message = AddSpaces(Param(CurLine, 5))
      Hostmask = Param(CurLine, 6)
      Exit Sub
    End If
  Next CurIdx
  If IsHostmask Then Nick = "": Exit Sub
  If Not EnableAverage Then Nick = "": Exit Sub
  
  'Second try, get the best match
  HGrade = 0
  Average = True
  For CurIdx = 1 To LineCount
    CurLine = FileLines(CurIdx)
    u = MatchGrade(SearchFor, Param(CurLine, 1))
    If u > HGrade Then
      HGrade = u
      Nick = Param(CurLine, 1)
      Alias = Param(CurLine, 2)
      LastSeen = AddSpaces(Param(CurLine, 3))
      SeenChan = Param(CurLine, 4)
      Message = AddSpaces(Param(CurLine, 5))
      Hostmask = Param(CurLine, 6)
    End If
  Next CurIdx
  If HGrade < 54 Then Nick = "": Average = False
  
Exit Sub
RESEErr:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  Close #FileNumber
  RemAccessedFile FileName
  SendNote "AnGeL ReadSeenEntry ERROR", "Hippo", "", "Der Fehler " & ErrNumber & " (" & ErrDescription & ") ist aufgetreten."
  SendNote "AnGeL ReadSeenEntry ERROR", "sensei", "", "Der Fehler " & ErrNumber & " (" & ErrDescription & ") ist aufgetreten."
End Sub

Sub WriteExtSeenEntry(Nick As String, Alias As String, LastSeen As Date, SeenChan As String, Message As String, Hostmask As String) ' : AddStack "SeenRoutines_WriteExtSeenEntry(" & Nick & ", " & Alias & ", " & LastSeen & ", " & SeenChan & ", " & Message & ", " & Hostmask & ")"
Dim u As Long, u2 As Long, RemovedOne As Boolean
  For u = ExtSeenCacheCount To 1 Step -1
    If LCase(Param(ExtSeenCache(u), 1)) = LCase(Nick) Then
      For u2 = u To ExtSeenCacheCount - 1
        ExtSeenCache(u2) = ExtSeenCache(u2 + 1)
      Next u2
      ExtSeenCacheCount = ExtSeenCacheCount - 1
    End If
  Next u
  ExtSeenCacheCount = ExtSeenCacheCount + 1: If ExtSeenCacheCount > UBound(ExtSeenCache()) Then ReDim Preserve ExtSeenCache(UBound(ExtSeenCache()) + 5)
  ExtSeenCache(ExtSeenCacheCount) = Nick & " " & IIf(Alias <> "", Alias, "-") & " " & Replace(CStr(CDbl(LastSeen)), ",", ".") & " " & IIf(SeenChan <> "", SeenChan, RemSpaces("<an unknown channel>")) & " " & IIf(Message <> "", RemSpaces(Message), RemSpaces("<no quit message>")) & " " & Hostmask
End Sub

Sub FlushExtSeenEntries() 'NoTrap' : AddStack "SeenRoutines_FlushExtSeenEntries()"
Dim FileName As String, FileLines() As String, LineCount As Long, u As Long, ErrLine As Long
Dim CurLine As String, FileNumber As Integer, Rest As String, Added As Boolean
Dim Nicks As String
  If ExtSeenCacheCount = 0 Then Exit Sub
  
  ReDim FileLines(100)
  FileName = HomeDir + "ExtSeen.txt"
  
  'Add new lines first
  For u = ExtSeenCacheCount To 1 Step -1
    LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 100)
    FileLines(LineCount) = ExtSeenCache(u)
    Nicks = Nicks & " " & Param(ExtSeenCache(u), 1)
  Next u
  Nicks = " " & LCase(Trim(Nicks)) & " "
  On Local Error Resume Next
  
  'Add all other seen file entries
  If Dir(FileName) <> "" Then
    WaitForAccess FileName
    AddAccessedFile FileName
    FileNumber = FreeFile: Open FileName For Input As #FileNumber
    If Err.Number <> 0 Then Close #FileNumber: RemAccessedFile FileName: Exit Sub 'Error; don't change anything
    On Error GoTo BigErr2
ErrLine = 3
    Do While Not EOF(FileNumber)
ErrLine = 4
      Line Input #FileNumber, CurLine
ErrLine = 5
      If Trim(CurLine) <> "" Then
        If InStr(Nicks, " " & LCase(Param(CurLine, 1)) & " ") = 0 Then
ErrLine = 6
          LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 100)
ErrLine = 7
          FileLines(LineCount) = CurLine
        End If
      End If
    Loop
ErrLine = 8
    Close #FileNumber
  Else
    WaitForAccess FileName
    AddAccessedFile FileName
  End If
  On Error GoTo BigErr2
ErrLine = 9
  FileNumber = FreeFile: Open FileName For Output As #FileNumber
ErrLine = 10
    For u = 1 To IIf(LineCount > MaxExtSeenCount, MaxExtSeenCount, LineCount)
ErrLine = 11
      Print #FileNumber, FileLines(u)
ErrLine = 12
    Next u
ErrLine = 13
  Close #FileNumber
  RemAccessedFile FileName
  ExtSeenCacheCount = 0
  ReDim ExtSeenCache(10)
Exit Sub
BigErr2:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  Close #FileNumber
  RemAccessedFile FileName
  SendNote "AnGeL FlushExtSeenEntries ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & ErrLine & "] ist aufgetreten."
  SendNote "AnGeL FlushExtSeenEntries ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & ErrLine & "] ist aufgetreten."
End Sub

Sub ClearSeenEntries()
Dim FileName As String, FileLines() As String, LineCount As Long, u As Long, ErrLine As Long
Dim CurLine As String, FileNumber As Integer, Rest As String, Added As Boolean
  ReDim FileLines(100)
  FileName = HomeDir & "Seen.txt"
  On Local Error Resume Next
  If Dir(FileName) <> "" Then
    WaitForAccess FileName
    AddAccessedFile FileName
    FileNumber = FreeFile
    Open FileName For Input As #FileNumber
    If Err.Number <> 0 Then
      Close #FileNumber
      Err.Clear
      RemAccessedFile FileName
      Exit Sub
    End If
    On Error GoTo ClearErr
ErrLine = 1
    Do While Not EOF(FileNumber)
ErrLine = 4
      Line Input #FileNumber, CurLine
ErrLine = 5
      If (GetUserNum(Param(CurLine, 1)) > 0) And Trim(CurLine) <> "" Then
ErrLine = 6
        LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
ErrLine = 7
        FileLines(LineCount) = CurLine
      End If
    Loop
ErrLine = 8
    Close #FileNumber
  Else
    Exit Sub
  End If
  On Error GoTo ClearErr
ErrLine = 9
  FileNumber = FreeFile: Open FileName For Output As #FileNumber
ErrLine = 10
    For u = 1 To IIf(LineCount > 2000, 2000, LineCount)
ErrLine = 11
      Print #FileNumber, FileLines(u)
ErrLine = 12
    Next u
ErrLine = 13
  Close #FileNumber
  RemAccessedFile FileName
Exit Sub
ClearErr:
  Close #FileNumber
  RemAccessedFile FileName
  Err.Clear
'  'Stop
End Sub

Sub WriteSeenEntry(Nick As String, Alias As String, LastSeen As Date, SeenChan As String, Message As String, Hostmask As String)
Dim FileName As String, FileLines() As String, LineCount As Long, u As Long, ErrLine As Long
Dim CurLine As String, FileNumber As Integer, Rest As String, Added As Boolean
  ReDim FileLines(100)
  FileName = HomeDir & "Seen.txt"
  On Local Error Resume Next
  If Dir(FileName) <> "" Then
    WaitForAccess FileName
    AddAccessedFile FileName
    FileNumber = FreeFile
    Open FileName For Input As #FileNumber
    If Err.Number <> 0 Then
      Err.Clear
      Close #FileNumber
      RemAccessedFile FileName
      Exit Sub
    End If
    On Error GoTo BigErr
ErrLine = 1
    LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
ErrLine = 2
    FileLines(LineCount) = Nick & " " & IIf(Alias <> "", Alias, "-") & " " & Replace(CStr(CDbl(LastSeen)), ",", ".") & " " & IIf(SeenChan <> "", SeenChan, RemSpaces("<an unknown channel>")) & " " & IIf(Message <> "", RemSpaces(Message), RemSpaces("<no quit message>")) & " " & Hostmask
ErrLine = 3
    Do While Not EOF(FileNumber)
ErrLine = 4
      Line Input #FileNumber, CurLine
ErrLine = 5
      If LCase(Param(CurLine, 1)) <> LCase(Nick) And Trim(CurLine) <> "" Then
ErrLine = 6
        LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
ErrLine = 7
        FileLines(LineCount) = CurLine
      End If
    Loop
ErrLine = 8
    Close #FileNumber
  Else
    WaitForAccess FileName
    AddAccessedFile FileName
    LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
ErrLine = 81
    FileLines(LineCount) = Nick & " " & IIf(Alias <> "", Alias, "-") & " " & CStr(CDbl(LastSeen)) & " " & IIf(SeenChan <> "", SeenChan, RemSpaces("<an unknown channel>")) & " " & IIf(Message <> "", RemSpaces(Message), RemSpaces("<no quit message>")) & " " & Hostmask
  End If
  On Error GoTo BigErr
ErrLine = 9
  FileNumber = FreeFile: Open FileName For Output As #FileNumber
ErrLine = 10
    For u = 1 To IIf(LineCount > 2000, 2000, LineCount)
ErrLine = 11
      Print #FileNumber, FileLines(u)
ErrLine = 12
    Next u
ErrLine = 13
  Close #FileNumber
  RemAccessedFile FileName
Exit Sub
BigErr:
  Close #FileNumber
  RemAccessedFile FileName
  Err.Clear
  'Stop
'  SendNote "AnGeL WriteSeenEntry ERROR", "Hippo", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & ErrLine) & "] ist aufgetreten."
'  SendNote "AnGeL WriteSeenEntry ERROR", "sensei", "", "Der Fehler " & Err.Number & " (" & Err.Description & ") - [" & ErrLine) & "] ist aufgetreten."
End Sub

