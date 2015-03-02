Attribute VB_Name = "Plugin_KI"
',-======================- ==-- -  -
'|   AnGeL - Plugins - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private LastSentences(10) As String, CurSentence As Long

Type KI_User
  Nick As String
  Thema As String
  Like As Integer
  LastSentences(10) As String
  CurSentence As Byte
  LastActionTime As Currency
End Type


Public KIGender As String
Public KIAge As String
Public KIFName As String
Public KILName As String
Public KICity As String
Public KICountry As String


Public KIs() As KI_User, KICount As Long


Sub KI_Load()
  ReDim Preserve KIs(5)
End Sub

Sub KI_Unload()
'
End Sub

Function MakeReply(Satz As String, NickName As String) As String ' : AddStack "KI_MakeReply(" & Satz & ", " & NickName & ")"
'Dim Subj As String, Praed As String, Obj As String, u As Long
'Dim NewOne As String, NowPraed As Boolean, SSatz As String
'Dim SavSatz As String, SatSatz As String, Subjadd As String
'Dim Tries As Byte, AddC As String, KNum As Long, TempStr As String
'Randomize Timer
'  KNum = 0
'  For u = 1 To KICount
'    If LCase(KIs(u).Nick) = LCase(NickName) Then KNum = u: Exit For
'  Next u
'  If KNum = 0 Then
'    KICount = KICount + 1: KNum = KICount
'    If KICount > UBound(KIs()) Then ReDim Preserve KIs(UBound(KIs()) + 5)
'    KIs(KNum).Nick = NickName
'    KIs(KNum).Like = 0
'    KIs(KNum).CurSentence = 0
'    KIs(KNum).Thema = ""
'    For u = 1 To 10
'      KIs(KNum).LastSentences(u) = ""
'    Next u
'  End If
'  KIs(KNum).LastActionTime = WinTickCount
'
'  'Remove unused KI threads after 4 minutes
'  For u = 1 To EventCount
'    If Events(u).DoThis = "RemKI " & NickName Then Events(u).DoThis = "": Exit For
'  Next u
'  TimedEvent "RemKI " & NickName, 240
'
'  For u = 1 To 10
'    If Satz = KIs(KNum).LastSentences(u) Then
'      Select Case KIs(KNum).Like
'        Case Is > 0
'          Select Case Int(Rnd * 7) + 1
'            Case 1: MakeReply = Choose(Int(Rnd * 5) + 1, "Hhm - ", "Erwischt! ", "Hey, ", "", "") & "du wiederholst dich! " & Choose(Int(Rnd * 2) + 1, ";", ":") & "-)" & Choose(Int(Rnd * 3) + 1, ")", "))", "")
'            Case 2: MakeReply = "Du mußt deine Sätze nicht mehrmals sagen :-)"
'            Case 3: MakeReply = Choose(Int(Rnd * 3) + 1, "Du, ich glaube, d", "D", "D") & "as hast du " & Choose(Int(Rnd * 2) + 1, "schon", "schonmal", "mir schon") & " gesagt! ;-)"
'            Case 4: MakeReply = ""
'            Case 5: MakeReply = "Mmh? Wieso wiederholst du dich?"
'            Case 6: MakeReply = "Warum sagst du das doppelt?"
'            Case 7: MakeReply = "Irgendwie wiederholst du dich! =)"
'          End Select
'          If Int(Rnd * 3) = 1 Then KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 0
'          Select Case Int(Rnd * 10) + 1
'            Case 1: MakeReply = Choose(Int(Rnd * 5) + 1, "Ähm... ", "Mmh. ", "Hey, ", "", "") & "du wiederholst dich!"
'            Case 2: MakeReply = Choose(Int(Rnd * 3) + 1, "Denkst du, ", "Glaubst du, ", "Hoffst du, ") + Choose(Int(Rnd * 3), "du kriegst ne andere Antwort, ", "eine andere Antwort von mir zu bekommen, ", "eine andere Antwort zu kriegen, ") & " indem du dich wiederholst?": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 3: MakeReply = Choose(Int(Rnd * 3) + 1, "Ich glaube, d", "Ich weiß - d", "D") & "as hast du " & Choose(Int(Rnd * 2) + 1, "schon einmal", "schonmal", "bereits") & " gesagt."
'            Case 4: MakeReply = Choose(Int(Rnd * 3) + 1, "Jep, ich weiß...", "I know.", "Ich weiß...")
'            Case 5: MakeReply = "Du wiederholst dich...": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 6: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 7: MakeReply = "Warum sagst du das doppelt?"
'            Case 8: MakeReply = Choose(Int(Rnd * 3) + 1, "Ich hab dich auch beim ersten Mal verstanden.", "Ich hab dich auch beim 1. Mal verstanden.", "Ich habe dich auch am Anfang verstanden."): KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 9: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 10: MakeReply = "Irgendwie wiederholst du dich =)": KIs(KNum).Like = KIs(KNum).Like - 1
'          End Select
'        Case -2 To -1
'          Select Case Int(Rnd * 10) + 1
'            Case 1: MakeReply = "Halt die Klappe.": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 2: MakeReply = "Halt bitte die Klappe!"
'            Case 3: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 4: MakeReply = "Fresse du " & Choose(Int(Rnd * 3) + 1, "Sack", "Wichser", "Spast") & "!": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 5: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 6: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 7: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 8: MakeReply = "ES REICHT!": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 9: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'            Case 10: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'          End Select
'        Case Is < -2
'          MakeReply = ""
'      End Select
'      KIs(KNum).Thema = "dbl"
'      If Int(Rnd * 5) + 1 < 3 Then MakeReply = ""
'      Exit Function
'    End If
'  Next u
'  KIs(KNum).CurSentence = KIs(KNum).CurSentence + 1: If KIs(KNum).CurSentence > 10 Then KIs(KNum).CurSentence = 1
'  KIs(KNum).LastSentences(KIs(KNum).CurSentence) = Satz
'
'  Satz = Trim(Satz)
'  SatSatz = Satz
'
'  Satz = RemStr(Satz, "=)")
'  Satz = RemStr(Satz, ":)")
'  Satz = RemStr(Satz, ":P")
'  Satz = RemStr(Satz, ";P")
'  Satz = RemStr(Satz, ":-P")
'  Satz = RemStr(Satz, ";-P")
'  Satz = RemStr(Satz, ":-)")
'  Satz = RemStr(Satz, "*^^*")
'  Satz = RemStr(Satz, "^_^;")
'  Satz = RemStr(Satz, "^_^")
'  Satz = RemStr(Satz, "^.^")
'  Satz = RemStr(Satz, ":-(")
'  Satz = RemStr(Satz, ":~(")
'  Satz = RemStr(Satz, ":>")
'  Satz = RemStr(Satz, ":<")
'  Satz = RemStr(Satz, ":->")
'  Satz = RemStr(Satz, ":-<")
'  Satz = RemStr(Satz, ":o)")
'  Satz = RemStr(Satz, ";o)")
'  Satz = RemStr(Satz, "=o)")
'  Satz = RemStr(Satz, ":o(")
'  Satz = RemStr(Satz, "=o(")
'  Satz = RemStr(Satz, ";)")
'  Satz = RemStr(Satz, ";-)")
'  While RemStr(Satz, "!!") <> Satz
'    Satz = RemStr(Satz, "!!"): AddC = "!"
'  Wend
'  If Right(Satz, 1) <> AddC Then Satz = Satz + AddC: AddC = ""
'  While RemStr(Satz, "??") <> Satz
'    Satz = RemStr(Satz, "??"): AddC = "?"
'  Wend
'  If Right(Satz, 1) <> AddC Then Satz = Satz + AddC: AddC = ""
'  For u = 1 To Len(Satz)
'    Select Case Mid(Satz, u, 1)
'      Case "=", ")", "(", "^", "_", "°", ":", ";", "!", "."
'        SSatz = SSatz
'      Case Else
'        SSatz = SSatz + Mid(Satz, u, 1)
'    End Select
'  Next u
'  Satz = Trim(SSatz)
'  SavSatz = SSatz
'  If (InStr(Satz, "daß") > 0) And (InStr(Satz, "daß") + 2 < Len(Satz)) Then
'    Satz = Right(Satz, Len(Satz) - InStr(Satz, "daß") - 3)
'    TempStr = Mid(Satz, Len(Param(Satz, 1)) + 2)
'    Satz = Trim(Param(Satz, 1) & " " & Param(Satz, ParamCount(Satz)) & " " & Left(TempStr, Len(TempStr) - Len(Param(TempStr, ParamCount(TempStr)))))
'  End If
'  SSatz = ""
'  For u = 1 To Len(Satz)
'    Select Case Mid(Satz, u, 1)
'      Case ".", "!", "?", ","
'        SSatz = SSatz
'      Case Else
'        SSatz = SSatz + Mid(Satz, u, 1)
'    End Select
'  Next u
'  SSatz = SSatz & " "
'  If Right(Satz, 1) = "." Or Right(Satz, 1) = "!" Then Satz = Left(Satz, Len(Satz) - 1)
'  For Tries = 1 To 2
'    NowPraed = False: u = 0: Subj = "": Praed = "": NewOne = "": Subjadd = ""
'    While u < ParamCount(Satz)
'      u = u + 1
'      If IsPossessive(Param(Satz, u)) And (Subj = "") Then
'        Subj = SwitchIfNeeded(Param(Satz, u)) & " " & Param(Satz, u + 1)
'        Praed = TurnAround(Param(Satz, u + 2))
'        u = u + 2
'      ElseIf (ConvertSubj(Param(Satz, u + 1), Tries) <> "") And (Subj = "") And (u > 1) Then
'        Subj = ConvertSubj(Param(Satz, u + 1), Tries)
'        Praed = TurnAround(Param(Satz, u))
'      ElseIf NowPraed Then
'        Praed = TurnAround(Param(Satz, u))
'        NowPraed = False
'      ElseIf (ConvertSubj(Param(Satz, u), Tries) <> "") And (Subj = "") Then
'        If u = 1 Then
'          Subj = ConvertSubj(Param(Satz, u), Tries)
'          NowPraed = True
'        End If
'      Else
'        Select Case LCase(Param(Satz, u))
'        Case "doch", "ja", "auch", "*g*", "so"
'          NewOne = NewOne
'        Case Else
'          If SwitchIfNeeded(Param(Satz, u)) <> Subj Then
'            If IsRel(Param(Satz, u)) Then
'              Subjadd = Subjadd & " " & SwitchIfNeeded(Param(Satz, u))
'            Else
'              NewOne = NewOne & " " & SwitchIfNeeded(Param(Satz, u))
'            End If
'          End If
'        End Select
'      End If
'    Wend
'    If Subj <> "" Then Exit For
'  Next Tries
'  SSatz = Trim(SSatz) & " "
'  If IsInLine(Satz, "hello|hallo|hai|hiya|hiho|hi|hoi|tach|moin|hepps|huhu|sei gegrüßt") Then
'    If (KIs(KNum).Thema = "hiagain") Then
'      Select Case KIs(KNum).Like
'        Case Is > 0
'          Select Case Int(Rnd * 5) + 1
'            Case 1: MakeReply = "*lach* zum x-ten Mal hi!"
'            Case 2: MakeReply = "Jaaaa, hi!!!"
'          End Select
'        Case 0
'          Select Case Int(Rnd * 5) + 1
'            Case 1: MakeReply = "verdammt nochmal, HI!"
'            Case 2: MakeReply = "wir haben schon oft genug hi gesagt!"
'            Case 3: MakeReply = "langsam reichts..."
'          End Select
'        Case -2 To -1
'          Select Case Int(Rnd * 9) + 1
'            Case 1: MakeReply = "GRRR!"
'            Case 2: MakeReply = "grr"
'            Case 3: MakeReply = "depp"
'            Case 4: MakeReply = "idiot"
'          End Select
'        Case Is < -2
'          MakeReply = ""
'      End Select
'      KIs(KNum).Like = KIs(KNum).Like - 1
'    ElseIf (KIs(KNum).Thema = "hi") Or (KIs(KNum).Thema = "howareyou") Then
'      Select Case Int(Rnd * 3) + 1
'        Case 1: MakeReply = "Nochmal hi :)": KIs(KNum).Thema = "hiagain"
'        Case 2: MakeReply = "nochmal hi ;)": KIs(KNum).Thema = "hiagain"
'        Case 3: MakeReply = "hi again": KIs(KNum).Thema = "hiagain"
'      End Select
'    Else
'      If Not IsInLine(Satz, "was?geht|wie?gehts|wie?geht?s|wie?geht?es|wie?fühlst|alles?fit|alles?klar") Then
'        Select Case KIs(KNum).Like
'          Case Is > 0
'            Select Case Int(Rnd * 14) + 1
'              Case 1: MakeReply = "Hallo " & NickName & "!": KIs(KNum).Thema = "hi"
'              Case 2: MakeReply = "Moin moin! | Wie gehts dir? :-)": KIs(KNum).Thema = "howareyou"
'              Case 3: MakeReply = "Hi!!" & Choose(Int(Rnd * 3) + 1, "*knuddel*", "*riesenknuddel*", "*freu!*"): KIs(KNum).Thema = "hi"
'              Case 4: MakeReply = "auch hi!" & Choose(Int(Rnd * 3) + 1, " wie gehts?", " Was macht das Leben? :-)", " was gibts neues?"): KIs(KNum).Thema = "howareyou"
'              Case 5: MakeReply = "Hi :))": KIs(KNum).Thema = "hi"
'              Case 6: MakeReply = "Hiya... wie gehts dir?": KIs(KNum).Thema = "howareyou"
'              Case 7: MakeReply = "Heyho! Kenne ich dich?! *scherz!* :-)": KIs(KNum).Thema = "hi"
'              Case 8: MakeReply = "hoi " & NickName & "!" & Choose(Int(Rnd * 4) + 1, " *knuddel*", " *drück*", " *freu*", ""): KIs(KNum).Thema = "hi"
'              Case 9: MakeReply = "hallo =))": KIs(KNum).Thema = "hi"
'              Case 10: MakeReply = "Hi! Wie geht's dir?": KIs(KNum).Thema = "howareyou"
'              Case 11: MakeReply = "Hepps.. wie gehts?": KIs(KNum).Thema = "howareyou"
'              Case 12: MakeReply = "hi! was gibts neues?": KIs(KNum).Thema = "howareyou"
'              Case 13: MakeReply = "hepps!": KIs(KNum).Thema = "hi"
'              Case 14: MakeReply = "Tach! :-)": KIs(KNum).Thema = "hi"
'            End Select
'          Case 0
'            Select Case Int(Rnd * 10) + 1
'              Case 1: MakeReply = "Hallo!": KIs(KNum).Thema = "hi"
'              Case 2: MakeReply = "Sei gegrüßt!": KIs(KNum).Thema = "hi"
'              Case 3: MakeReply = "auch hi! | wie gehts?": KIs(KNum).Thema = "howareyou"
'              Case 4: MakeReply = "Hi!": KIs(KNum).Thema = "hi"
'              Case 5: MakeReply = "Hi... wie gehts?": KIs(KNum).Thema = "howareyou"
'              Case 6: MakeReply = "Hi! Kenne ich dich?!": KIs(KNum).Thema = "hi"
'              Case 7: MakeReply = "hoi": KIs(KNum).Thema = "hi"
'              Case 8: MakeReply = "hallo =)": KIs(KNum).Thema = "hi"
'              Case 9: MakeReply = "Hi!": KIs(KNum).Thema = "hi"
'              Case 10: MakeReply = "hi!": KIs(KNum).Thema = "hi"
'            End Select
'          Case -2 To -1
'            Select Case Int(Rnd * 10) + 1
'              Case 1: MakeReply = "Hallo.": KIs(KNum).Thema = "hi"
'              Case 2: MakeReply = "hmpf." & Choose(Int(Rnd * 2), "..", "")
'              Case 3: MakeReply = "hi.": KIs(KNum).Thema = "hi"
'              Case 4: MakeReply = ""
'              Case 5: MakeReply = "Ja?"
'              Case 6: MakeReply = "Was ist denn schon wieder?"
'              Case 7: MakeReply = ""
'              Case 8: MakeReply = "Hi": KIs(KNum).Thema = "hi"
'              Case 9: MakeReply = ""
'              Case 10: MakeReply = ""
'            End Select
'          Case Is < -2
'            MakeReply = ""
'        End Select
'      Else
'        Select Case KIs(KNum).Like
'          Case Is > 0
'            Select Case Int(Rnd * 8) + 1
'              Case 1: MakeReply = "Hallo " & NickName & " " & Choose(Int(Rnd * 4) + 1, "- danke, ", "thx, ", "- ", "... ") & "mir gehts gut! :)": KIs(KNum).Thema = "hi"
'              Case 2: MakeReply = "Hi! Mir gehts gut, danke der Nachfrage! | Und wie stehts mit dir?": KIs(KNum).Thema = "howareyou"
'              Case 3: MakeReply = "Hi!!" & Choose(Int(Rnd * 3) + 1, "*knuddel*", "*riesenknuddel*", "*freu!*") & " - " & "Mir gehts klasse, und dir?": KIs(KNum).Thema = "hi"
'              Case 4: MakeReply = "auch hi - danke, gut!" & Choose(Int(Rnd * 3) + 1, " und wie gehts dir?", " Was macht das Leben? :-)", " was gibts neues?"): KIs(KNum).Thema = "howareyou"
'              Case 5: MakeReply = "Hi :)) Och, geht so! | Warum fragst du?": KIs(KNum).Thema = "hi"
'              Case 6: MakeReply = "Hiya... gut, und wie gehts dir?": KIs(KNum).Thema = "howareyou"
'              Case 7: MakeReply = "Heyho! Kenne ich dich?! *scherz!* :-) Gut gehts!": KIs(KNum).Thema = "hi"
'              Case 8: MakeReply = "Selber hi :) | Naja, ich lagge, wie üblich ;o)": KIs(KNum).Thema = "hi"
'            End Select
'          Case 0
'            Select Case Int(Rnd * 8) + 1
'              Case 1: MakeReply = "Hepps " & NickName & " " & Choose(Int(Rnd * 4) + 1, "- danke, ", "thx, ", "- ", "... ") & "es geht so :)": KIs(KNum).Thema = "hi"
'              Case 2: MakeReply = "Hi! Mir gehts gut, wie stehts mit dir?": KIs(KNum).Thema = "howareyou"
'              Case 3: MakeReply = "Hoi" & Choose(Int(Rnd * 3) + 1, "!", "...", ".") & " " & "Mir gehts gut, und dir?": KIs(KNum).Thema = "hi"
'              Case 4: MakeReply = "Ebenfalls hi - danke, gut!" & Choose(Int(Rnd * 3) + 1, " und wie gehts dir?", " Was macht das Leben? :)", " was geht ab?"): KIs(KNum).Thema = "howareyou"
'              Case 5: MakeReply = "Hi :)) Naja, geht so! Warum fragst du?": KIs(KNum).Thema = "hi"
'              Case 6: MakeReply = "Hi " & NickName & "... gut, und wie gehts dir?": KIs(KNum).Thema = "howareyou"
'              Case 7: MakeReply = "Hoi, kenne ich dich? Mir gehts gut soweit.": KIs(KNum).Thema = "hi"
'              Case 8: MakeReply = "Selber hi :) Naja, ich weiß nicht so recht... ich lagge heute ziemlich oft =)": KIs(KNum).Thema = "hi"
'            End Select
'          Case -2 To -1
'            Select Case Int(Rnd * 10) + 1
'              Case 1: MakeReply = "Hallo. Geht so.": KIs(KNum).Thema = "hi"
'              Case 2: MakeReply = "pah." & Choose(Int(Rnd * 2), "..", "")
'              Case 3: MakeReply = "hi.": KIs(KNum).Thema = "hi"
'              Case 4: MakeReply = ""
'              Case 5: MakeReply = "Ja, hi.. Bin schlecht gelaunt."
'              Case 6: MakeReply = "No comment..."
'              Case 7: MakeReply = ""
'              Case 8: MakeReply = "Hi - hält sich in Grenzen.": KIs(KNum).Thema = "hi"
'              Case 9: MakeReply = ""
'              Case 10: MakeReply = ""
'            End Select
'          Case Is < -2
'            MakeReply = ""
'        End Select
'      End If
'    End If
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "re|reh|rehi|ree|reeh|röh|back|wieder?da") Then
'    If (KIs(KNum).Thema = "hi") Or (KIs(KNum).Thema = "howareyou") Then
'      MakeReply = "Nochmal hi :-)"
'    Else
'      Select Case KIs(KNum).Like
'        Case Is > 0
'          Select Case Int(Rnd * 7) + 1
'            Case 1: MakeReply = "rehi " & NickName & "!": KIs(KNum).Thema = "hi"
'            Case 2: MakeReply = "reh. Wie gehts?": KIs(KNum).Thema = "howareyou"
'            Case 3: MakeReply = "ree! " & Choose(Int(Rnd * 3) + 1, "*knuddel*", "*knutsch*", "*freu!*"): KIs(KNum).Thema = "hi"
'            Case 4: MakeReply = "rehi!" & Choose(Int(Rnd * 3) + 1, " was gibts neues?", " Was war?", " was gibts neues?"): KIs(KNum).Thema = "howareyou"
'            Case 5: MakeReply = "Hi :))": KIs(KNum).Thema = "hi"
'            Case 6: MakeReply = "Hiya...": KIs(KNum).Thema = "hi"
'            Case 7: MakeReply = "hepps!": KIs(KNum).Thema = "hi"
'          End Select
'        Case 0
'          Select Case Int(Rnd * 6) + 1
'            Case 1: MakeReply = "Re!": KIs(KNum).Thema = "hi"
'            Case 2: MakeReply = "Rehi!": KIs(KNum).Thema = "hi"
'            Case 3: MakeReply = "rehi wie gehts?": KIs(KNum).Thema = "howareyou"
'            Case 4: MakeReply = "Hi!": KIs(KNum).Thema = "hi"
'            Case 5: MakeReply = "re": KIs(KNum).Thema = "hi"
'            Case 6: MakeReply = "re! | was war?": KIs(KNum).Thema = "howareyou"
'          End Select
'        Case -2 To -1
'          Select Case Int(Rnd * 10) + 1
'            Case 1: MakeReply = "Hallo.": KIs(KNum).Thema = "hi"
'            Case 2: MakeReply = "jo." & Choose(Int(Rnd * 2), "..", "")
'            Case 3: MakeReply = "tach.": KIs(KNum).Thema = "hi"
'            Case 4: MakeReply = ""
'            Case 5: MakeReply = "Ja?"
'            Case 6: MakeReply = "Jo, re."
'            Case 7: MakeReply = ""
'            Case 8: MakeReply = "Hi": KIs(KNum).Thema = "hi"
'            Case 9: MakeReply = ""
'            Case 10: MakeReply = ""
'          End Select
'        Case Is < -2
'          MakeReply = ""
'      End Select
'    End If
'    Exit Function
'  End If
'
'  'Clear questions - clear answers ;)
'  If IsInLine(Satz, "wer bist du|wie heißt du|wie heisst du|wie dein name|was dein name|deinen name?") Then
'    If KIFName = "" Then
'      TempStr = Choose(Int(Rnd * 13) + 1, "Mirko", "Franziska", "Daniela", "Jochen", "Sheila", "Julia", "Jenny", "Paul", "Anette", "Johanna", "Janine", "Liv", "Alex")
'    Else
'      TempStr = KIFName
'    End If
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Ich bin " & TempStr & "! | Und du?": KIs(KNum).Thema = "whatsyourname"
'      Case 2: MakeReply = "Hmm? Ich heiße " & TempStr & ". Und du?": KIs(KNum).Thema = "whatsyourname"
'      Case 3: MakeReply = "Hehe, neugierig? *G* Ich heiße " & TempStr & "! | Und wer bist du? :)": KIs(KNum).Thema = "whatsyourname"
'      Case 4: MakeReply = "Ich heiße " & TempStr & ", und du?": KIs(KNum).Thema = "whatsyourname"
'      Case 5: MakeReply = "Gestatten, " & TempStr & "! :-)) | Und wer bist du?": KIs(KNum).Thema = "whatsyourname"
'      Case 6: MakeReply = "Sag ich net :))"
'      Case 7: MakeReply = "Hmm. | Wieso so neugierig?  =)"
'      Case 8: MakeReply = "ich heiße " & TempStr & ". Jetzt will ich aber auch deinen Namen wissen :)": KIs(KNum).Thema = "whatsyourname"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "woher kommst*|woher bist|wo wohnst*|wo kommst*|aus welcher stadt du") Then
'    If KICity = "" Then
'      TempStr = Choose(Int(Rnd * 13) + 1, "Braunschweig", "Frankfurt", "München", "Berlin", "Stuttgart", "Saarbrücken", "Essen", "Erlangen", "Paderborn", "Karlsruhe", "Kaiserslautern", "Lahr", "Ludwigsfelde")
'    Else
'      TempStr = KICity
'    End If
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Ich komme aus " & TempStr & ", und du?": KIs(KNum).Thema = "wherefrom"
'      Case 2: MakeReply = "Aus " & TempStr & "..."
'      Case 3: MakeReply = "Ich bin aus " & TempStr & "!"
'      Case 4: MakeReply = "Sowas sag ich nicht :-)"
'      Case 5: MakeReply = "Aus " & TempStr & ", und du?": KIs(KNum).Thema = "wherefrom"
'      Case 6: MakeReply = "Ich wohne in " & TempStr & ". | Woher kommst du?": KIs(KNum).Thema = "wherefrom"
'      Case 7: MakeReply = "Ich komm aus " & TempStr & "! | Und wo wohnst du?": KIs(KNum).Thema = "wherefrom"
'      Case 8: MakeReply = "nene, das sag ich net =)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "wie alt|dein alter|age") Then
'    If KIAge = "" Then
'      TempStr = CStr(CInt(Int(Rnd * 10) + 15))
'    Else
'      TempStr = KIAge
'    End If
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = TempStr & ", und du?"
'      Case 2: MakeReply = "Ich bin " & TempStr & "."
'      Case 3: MakeReply = TempStr & "."
'      Case 4: MakeReply = "Ich bin 84 Jahre alt, junger Mann ;o)"
'      Case 5: MakeReply = TempStr & ". wie alt bist du?"
'      Case 6: MakeReply = "bin " & TempStr & "."
'      Case 7: MakeReply = "ich bin " & TempStr
'      Case 8: MakeReply = "sag ich nicht ;) ...na gut, " & TempStr & "."
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "m/f|männ* oder weib*|m oder f|m oder w|bist du w|bist du f|f oder m|w oder m|was bist du?") Then
'    If KIGender = "" Then
'      TempStr = Choose(Int(Rnd * 4) + 1, "weiblich", "f", "m", "w", "männlich")
'    Else
'      TempStr = IIf(KIGender = "f", Choose(Int(Rnd * 3) + 1, "weiblich", "f", "w", "femme fatale ;o)"), Choose(Int(Rnd * 4) + 1, "männlich", "m", "MANN", "geiler boy ;o)", "'standhaft' männlich' ;o)"))
'    End If
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = TempStr & "..."
'      Case 2: MakeReply = "ich bin " & TempStr & "!"
'      Case 3: MakeReply = TempStr & ", und du?"
'      Case 4: MakeReply = TempStr & ". ich wette, das wolltest du nicht hören ;)"
'      Case 5: MakeReply = "ich bin herrlich ;o)"
'      Case 6: MakeReply = TempStr
'      Case 7: MakeReply = "bin " & TempStr & "!"
'      Case 8: MakeReply = "warum willst du das wissen?"
'    End Select
'    Exit Function
'  End If
'  '---------
'
'  If IsInLine(Satz, "[hübscher|schöner|toller|netter|süßer] name|hast [hübschen|schönen|tollen|netten|süßen] [name|namen]|dein name* [hübsch|schön|toll|nett|süß]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Danke :))"
'      Case 2: MakeReply = "Freut mich daß er dir gefällt!"
'      Case 3: MakeReply = "thx =)"
'      Case 4: MakeReply = "danke ;)"
'      Case 5: MakeReply = "dankeschön! *knicksmach* :)"
'      Case 6: MakeReply = "daaaaanke! :-)"
'      Case 7: MakeReply = "ui, freut mich daß du ihn magst :)"
'      Case 8: MakeReply = "ja? danke!! :-)"
'    End Select
'    Exit Function
'  End If
'
'  'Thema: Wie heißt du?
'  If KIs(KNum).Thema = "whatsyourname" Then
'    If IsInLine(Satz, "nö|nee|nein|vergiss*es|nope|ne") Then
'      Select Case Int(Rnd * 8) + 1
'        Case 1: MakeReply = "pah..."
'        Case 2: MakeReply = "langweiler."
'        Case 3: MakeReply = "naja, dann halt net :/"
'        Case 4: MakeReply = "hm.. najo, dann nicht."
'        Case 5: MakeReply = "pff | ich versteh dich nicht."
'        Case 6: MakeReply = "menno :("
'        Case 7: MakeReply = "mist. | schon wieder keinen raubkopierer dingfest gemacht ;o)"
'        Case 8: MakeReply = "arsch ;P"
'      End Select
'      KIs(KNum).Thema = ""
'      Exit Function
'    End If
'    If IsInLine(Satz, "[sag*|verrat*] [nich*|net]|no comm*|kein kom*|nee|nö|nein") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = "Wieso nicht? | Bitte sag doch :)"
'        Case 2: MakeReply = "Och komm!"
'        Case 3: MakeReply = "Du kannst mir doch vertrauen *G*"
'        Case 4: MakeReply = "büdde sag =)"
'        Case 5: MakeReply = "gib dir nen ruck.. | sag schon :-)"
'        Case 6: MakeReply = "menno | jetzt sag schon ;)"
'        Case 7: MakeReply = "du kannst mir deinen namen ruhig sagen, | ich arbeite nich beim FBI ;)"
'        Case 8: MakeReply = "du kannsts mir ruhig sagen... | ich bin net bei der MS anti-piracy-abteilung *G*"
'        Case 9: MakeReply = "sach schon! ich schweige wie ein grab =)"
'      End Select
'      Exit Function
'    End If
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Hm, schöner Name :o)"
'      Case 2: MakeReply = "schöner name!"
'      Case 3: MakeReply = "nett, dich kennenzulernen!"
'      Case 4: MakeReply = "sehr erfreut ;)"
'      Case 5: MakeReply = "ah, hi :-)"
'      Case 6: MakeReply = "nett dich kennen zu lernen =)"
'      Case 7: MakeReply = "süßer name ;)"
'      Case 8: MakeReply = "hm, kennen wir uns?"
'    End Select
'    KIs(KNum).Thema = ""
'    Exit Function
'  End If
'
'  'Thema: Woher kommst du?
'  If KIs(KNum).Thema = "wherefrom" Then
'    If IsInLine(Satz, "[sag*|verrat*] [nich*|net]|no comm*|kein kom*|nee|nö|nein") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = "Wieso nicht? | Bitte sag doch :)"
'        Case 2: MakeReply = "Och bitte, sag!"
'        Case 3: MakeReply = "Du kannst mir vertrauen *gg*"
'        Case 4: MakeReply = "büdde sag =)"
'        Case 5: MakeReply = "gib dir nen ruck.. | sag schon :-)"
'        Case 6: MakeReply = "menno | jetzt sag schon ;)"
'        Case 7: MakeReply = "du kannst mir die stadt ruhig sagen! | ich bin harmlos =)"
'        Case 8: MakeReply = "du kannsts mir ruhig sagen... | ich bin net bei der MS anti-piracy-abteilung *G*"
'        Case 9: MakeReply = "sach schon! ich schweige wie ein grab =)"
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "aus|in|bei|nähe von") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = "Oh, schöne Stadt! =)"
'        Case 2: MakeReply = "Ist das weit von hier aus? :)"
'        Case 3: MakeReply = "Kenn ich nicht ;)"
'        Case 4: MakeReply = "Schönes Fleckchen :)"
'        Case 5: MakeReply = "Und, habt ihr gutes Wetter? *g*"
'        Case 6: MakeReply = "Kenn ich... das ist ne schöne Gegend."
'        Case 7: MakeReply = "Ich muß mal aufs Klo. Bin gleich wieder da."
'        Case 8: MakeReply = "Oh! | Das kenne ich sogar =)"
'        Case 9: MakeReply = "schöne stadt =)"
'      End Select
'      Exit Function
'    End If
'  End If
'
'  If IsInLine(Satz, "ich [bin|komm|komme|lebe|wohne] [aus|in|bei|nähe von]") Then
'    Select Case Int(Rnd * 9) + 1
'      Case 1: MakeReply = "Oh, schöne Stadt! =)"
'      Case 2: MakeReply = "Ist das weit von hier aus? :)"
'      Case 3: MakeReply = "Ich frage mich wie weit das von hier entfernt ist..."
'      Case 4: MakeReply = "Schönes Fleckchen :)"
'      Case 5: MakeReply = "Und, habt ihr gutes Wetter? *g*"
'      Case 6: MakeReply = "Kenn ich... das ist ne schöne Gegend."
'      Case 7: MakeReply = "Ich muß mal aufs Klo. Bin gleich wieder da."
'      Case 8: MakeReply = "Oh! | Das kenne ich sogar =)"
'      Case 9: MakeReply = "schöne stadt =)"
'    End Select
'    Exit Function
'  End If
'
'  'Thema: Wie gehts?
'  If (KIs(KNum).Thema = "howareyou") Or (KIs(KNum).Thema = "howareyou2") Then
'    If KIs(KNum).Thema = "howareyou" Then KIs(KNum).Thema = "howareyou2" Else KIs(KNum).Thema = ""
'    If IsInLine(Satz, "dir|du|was?geht|wie?gehts|wie?geht?s|wie?geht?es|wie?fühlst|alles?fit") Then
'      KIs(KNum).Thema = ""
'      Select Case KIs(KNum).Like
'        Case Is > 0
'          Select Case Int(Rnd * 14) + 1
'            Case 1, 2: MakeReply = Choose(Int(Rnd * 4) + 1, "Danke, ", "thx, ", "", "also ") & "mir gehts gut! :)"
'            Case 3, 4: MakeReply = Choose(Int(Rnd * 3) + 1, "Hehe, ", "och, ", "") & "gut gehts, danke der nachfrage!"
'            Case 5: MakeReply = Choose(Int(Rnd * 3) + 1, "also ", "hehe, ", "") & "mir gehts klasse!"
'            Case 6: MakeReply = "danke, gut!" & Choose(Int(Rnd * 3) + 1, "", " Was macht das Leben? =)", " was gibts neues?")
'            Case 7: MakeReply = "Och, geht so! Warum fragst du?"
'            Case 8, 9, 10: MakeReply = "Gut, danke!"
'            Case 11: MakeReply = "Gut gehts!"
'            Case 12: MakeReply = "Mir gehts klasse, danke!"
'            Case 13: MakeReply = "Naja, ich hab mir den Finger verknackst... :(": KIs(KNum).Thema = "malheur"
'            Case 14: MakeReply = "Ich habe mir dummerweise den Finger verknackst :~(": KIs(KNum).Thema = "malheur"
'          End Select
'        Case 0
'          Select Case Int(Rnd * 10) + 1
'            Case 1, 2, 3: MakeReply = Choose(Int(Rnd * 4) + 1, "Danke, ", "thx, ", "", "also ") & "es geht so :)"
'            Case 4: MakeReply = "Mir gehts gut... gibt nix zu meckern :)"
'            Case 5: MakeReply = "Ich fühle mich prächtig heute :))"
'            Case 6: MakeReply = "naja, bis auf daß es hier gerade anfängt zu regnen... gut!": KIs(KNum).Thema = "malheur"
'            Case 7: MakeReply = "Naja, geht so!"
'            Case 8: MakeReply = "Also mir gehts gut. :)"
'            Case 9: MakeReply = "Mir gehts gut soweit. | Kenn ich dich?"
'            Case 10: MakeReply = "Naja, ich weiß nicht so recht... ich lagge heute ziemlich oft =)": KIs(KNum).Thema = "malheur"
'          End Select
'        Case -2 To -1
'          Select Case Int(Rnd * 10) + 1
'            Case 1: MakeReply = "Naja, geht so."
'            Case 2: MakeReply = "es geht." & Choose(Int(Rnd * 2), "..", "")
'            Case 3: MakeReply = "lebbe geht weida."
'            Case 4: MakeReply = ""
'            Case 5: MakeReply = "Pff. Bin schlecht gelaunt."
'            Case 6: MakeReply = "No comment..."
'            Case 7: MakeReply = ""
'            Case 8: MakeReply = "Es hält sich in Grenzen."
'            Case 9: MakeReply = ""
'            Case 10: MakeReply = ""
'          End Select
'        Case Is < -2
'          MakeReply = ""
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "schlecht|mies|nicht gut|nicht ?oll|net ?oll|schrecklich|scheisse|scheiße|schlimm|unwohl") Then
'      Select Case Int(Rnd * 8) + 1
'        Case 1: MakeReply = "Oh! Warum das?"
'        Case 2: MakeReply = "Fühlst du dich öfters nicht so toll?"
'        Case 3: MakeReply = "Oh, das ist schade! Warum denn?"
'        Case 4: MakeReply = "Hhm... und warum gehts dir nicht so gut?"
'        Case 5: MakeReply = "Na komm, erzähl mir was dich bedrückt!"
'        Case 6: MakeReply = "Du kannst mir ruhig alles erzählen, was du auf dem Herzen hast!"
'        Case 7: MakeReply = "Hast du Liebeskummer oder sowas?"
'        Case 8: MakeReply = "Und was ist der Grund dafür, daß du dich nicht so toll fühlst?"
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "spitze|gut|super|toll|klasse|geht?so|es?geht|nicht?schlecht|wunderbar") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = Choose(Int(Rnd * 4) + 1, "na, ", "na ", "", "also ") & "das hört man doch gern!"
'        Case 2: MakeReply = Choose(Int(Rnd * 3) + 1, "Hehe, ", "super! ", "") & "freut mich!"
'        Case 3: MakeReply = Choose(Int(Rnd * 3) + 1, "hey, das ", "das ", "") & "freut mich :)"
'        Case 4: MakeReply = "sehr schön :))"
'        Case 5: MakeReply = "ah, freut mich das zu hören!"
'        Case 6: MakeReply = "hehe"
'        Case 7: MakeReply = "Gut gut... und, was gibts zu bereden?"
'        Case 8: MakeReply = "Ah gut, ich mag positive Leute =)"
'        Case 9: MakeReply = "Freut mich für dich!"
'      End Select
'      Exit Function
'    End If
'  End If
'
'  'Thema: Entschuldigung
'  If KIs(KNum).Thema = "sorry" Then
'    If IsInLine(Satz, "bitte|leid|ehrlich|echt|wirklich|sorry") Then
'      KIs(KNum).Thema = ""
'      Select Case Int(Rnd * 8) + 1
'        Case 1: MakeReply = "Ok, angenommen."
'        Case 2: MakeReply = "Na jut =)"
'        Case 3: MakeReply = "Hhm, okay ;) Akzeptiert."
'        Case 4: MakeReply = "Na gut, ist in Ordnung!"
'        Case 5: MakeReply = "Oki. Entschuldigung angenommen."
'        Case 6: MakeReply = "Najo, okay."
'        Case 7: MakeReply = "Schon gut =)"
'        Case 8: MakeReply = "Gut, akzeptiert."
'      End Select
'      KIs(KNum).Like = KIs(KNum).Like + 1
'      Exit Function
'    End If
'  End If
'
'  'Thema: Wirklich?
'  If KIs(KNum).Thema = "really?" Then
'    KIs(KNum).Thema = ""
'    If IsInLine(Satz, "nein|nee|noe|nö|nope|nicht|nich|net") Then
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "Hehe *G* Und wieso nicht?"
'        Case 2: MakeReply = "Plötzlicher Meinungsumschwung" & Choose(Int(Rnd * 2) + 1, ", oder wie soll ich das verstehen", "") & "? =)"
'        Case 3: MakeReply = "Naja gut, dann halt nich :-)"
'        Case 4: MakeReply = "Wieso denn jetzt auf einmal nicht?"
'        Case 5: MakeReply = "? aha | *smile*"
'        Case 6: MakeReply = "Was denn sonst?"
'        Case 7: MakeReply = "Nicht? Ja und was sonst?"
'        Case 8: MakeReply = "*gg*"
'        Case 9: MakeReply = "Hm. | Und wieso nicht?"
'        Case 10: MakeReply = "Wieso denn nicht?"
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "ja|jo|jup*|jau|jep|schon|ehrlich") Then
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "Hhm, klingt auch einleuchtend."
'        Case 2: MakeReply = "Aaahja =)"
'        Case 3: MakeReply = "Verstehe."
'        Case 4: MakeReply = "Aha."
'        Case 5: MakeReply = "Find ich interessant."
'        Case 6: MakeReply = "hehe *G*"
'        Case 7: MakeReply = "Oki."
'        Case 8: MakeReply = "Hehe =)"
'        Case 9: MakeReply = "Hm. Findest du das gut?"
'        Case 10: MakeReply = "Hehe :)"
'      End Select
'      Exit Function
'    End If
'  End If
'
'  'Thema: Hab ich nicht verstanden, kannst du mir das mal erklären?
'  If KIs(KNum).Thema = "helpme" Then
'    KIs(KNum).Thema = ""
'    If IsInLine(Satz, "nein|nee|nö|nope|nich|nicht|nix|nichts|vergis*|verges*|egal") Then
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "ahja. sehr aufschlussreich."
'        Case 2: MakeReply = "depp ;P"
'        Case 3: MakeReply = "na dann halt nicht.": KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 4: MakeReply = "dann nich.": KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 5: MakeReply = "hmz.": KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 6: MakeReply = "pah...": KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 7: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 8: MakeReply = ""
'        Case 9: MakeReply = "ah. vielen dank :/": KIs(KNum).Like = KIs(KNum).Like - 1
'        Case 10: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like - 1
'      End Select
'      Exit Function
'    End If
'  End If
'
'  'Nett dich kennenzulernen
'  If IsInLine(Satz, "nice meet you|[schön|angenehm|nett] dich [treffen|kennen*lernen]") Then
'    Select Case KIs(KNum).Like
'      Case Is > 0
'        Select Case Int(Rnd * 12) + 1
'          Case 1, 2: MakeReply = Choose(Int(Rnd * 4) + 1, "*knuddel* ", "*knuddelwuddel* ", "", "*liebhab* ") & "Ich freue mich auch dich zu kennen! :))"
'          Case 3, 4: MakeReply = ":)))) ich mag dich!"
'          Case 5: MakeReply = "*liebdrück* | gleichfalls! :)"
'          Case 6: MakeReply = "*drück*" & IIf(Int(Rnd * 2) = 0, " :-)", " :)")
'          Case 7: MakeReply = "Ich finds auch schön dich kennenzulernen!"
'          Case 8, 9, 10: MakeReply = "ich find dich total nett =)"
'          Case 11: MakeReply = "Gleichfalls =) | Ich freu mich immer wieder! *knuddel*"
'          Case 12: MakeReply = "Hab dich lieb!"
'        End Select
'      Case 0
'        Select Case Int(Rnd * 10) + 1
'          Case 1, 2, 3: MakeReply = "gleichfalls!"
'          Case 4: MakeReply = "nice to meet you :)"
'          Case 5: MakeReply = "sehr erfreut :-)"
'          Case 6: MakeReply = "=) | ist auch schön, dich kennenzulernen!"
'          Case 7: MakeReply = "danke! gleichfalls"
'          Case 8: MakeReply = "angenehm :)"
'          Case 9: MakeReply = "freut mich!"
'          Case 10: MakeReply = "das freut mich :)))"
'        End Select
'      Case -2 To -1
'        Select Case Int(Rnd * 10) + 1
'          Case 1: MakeReply = "no comment."
'          Case 2: MakeReply = "is scho recht" & Choose(Int(Rnd * 2), "..", "")
'          Case 3: MakeReply = "naja, es muß ja nicht alles auf gegenseitigkeit beruhen, oder?"
'          Case 4: MakeReply = ""
'          Case 5: MakeReply = "und du bist ein arschloch. deutlich genug? ;>"
'          Case 6: MakeReply = "kein kommentar"
'          Case 7: MakeReply = "ich bin natürlich ebenfalls erfreut, du sackgesicht" & Choose(Int(Rnd * 2), "!", "!!")
'          Case 8: MakeReply = "*tret*"
'          Case 9: MakeReply = ""
'          Case 10: MakeReply = ""
'        End Select
'      Case Is < -2
'        MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  'Magst du mich? *G*
'  If IsInLine(Satz, "[magst|wie findest] mich") Then
'    Select Case KIs(KNum).Like
'      Case Is > 0
'        Select Case Int(Rnd * 12) + 1
'          Case 1, 2: MakeReply = Choose(Int(Rnd * 4) + 1, "*knuddel* ", "*knuddelwuddel* ", "", "*liebhab* ") & "Ich mag dich sehr :))"
'          Case 3, 4: MakeReply = ":)))) ich find dich lieb =)"
'          Case 5: MakeReply = "*liebdrück* | antwort genug? :)"
'          Case 6: MakeReply = "*abschlabber*" & IIf(Int(Rnd * 2) = 0, " :o)", " :)")
'          Case 7: MakeReply = "*knuuuutsch* *lieb dich hab*! ;)"
'          Case 8, 9, 10: MakeReply = "ich find dich total nett =)"
'          Case 11: MakeReply = "ich mag dich sehr! | *knuddel*"
'          Case 12: MakeReply = "*knuddelknuddel* :))"
'        End Select
'      Case 0
'        Select Case Int(Rnd * 10) + 1
'          Case 1, 2, 3: MakeReply = ":) | wieso willst du das wissen? :)"
'          Case 4: MakeReply = "ich find dich nett... und du mich?"
'          Case 5: MakeReply = "ehrliche antwort oder nicht? =)"
'          Case 6: MakeReply = "*knuddel* ich mag dich!"
'          Case 7: MakeReply = "*knuddel*"
'          Case 8: MakeReply = "klar, warum nicht? :)"
'          Case 9: MakeReply = "ich hab nix gegen dich *G*"
'          Case 10: MakeReply = "naja, geht so... | ich weiß noch zu wenig über dich.. | wie heißt du? :)": KIs(KNum).Thema = "whatsyourname"
'        End Select
'      Case -2 To -1
'        Select Case Int(Rnd * 10) + 1
'          Case 1: MakeReply = "no comment."
'          Case 2: MakeReply = "das willst du nicht wissen" & Choose(Int(Rnd * 2), "..", "")
'          Case 3: MakeReply = "DICH mag ich auf jeden fall nicht!"
'          Case 4: MakeReply = ""
'          Case 5: MakeReply = "du bist ein arschloch. genug gesagt? ;>"
'          Case 6: MakeReply = "kein kommentar"
'          Case 7: MakeReply = "ich würde dir natürlich am liebsten die füße küssen, du sack" & Choose(Int(Rnd * 2), "!", "!!")
'          Case 8: MakeReply = "*tret*"
'          Case 9: MakeReply = ""
'          Case 10: MakeReply = ""
'        End Select
'      Case Is < -2
'        MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "ich [[mag?dich|liebe?dich] | [finde dich|finde du|find dich|find du] [spitze|genial|super|klasse|nett|toll|sympa*isch]] | [knuddel|kuschel|knutsch|drück|küss|schmatz]") Then
'    If IsInLine(Satz, "knuddel|drück") Then
'      Select Case KIs(KNum).Like
'        Case Is > 0
'          Select Case Int(Rnd * 12) + 1
'            Case 1, 2: MakeReply = Choose(Int(Rnd * 4) + 1, "*knuddel* ", "*knuddelwuddel* ", "", "*liebhab* ") & "Ich mag dich sehr :))"
'            Case 3, 4: MakeReply = ":)))) *Freu*"
'            Case 5: MakeReply = "*liebdrück*"
'            Case 6: MakeReply = "*abschlabber*" & IIf(Int(Rnd * 2) = 0, " :o)", " :)")
'            Case 7: MakeReply = "*knuuuutsch*"
'            Case 8, 9, 10: MakeReply = "*schmatz*"
'            Case 11: MakeReply = "*reknuddel*"
'            Case 12: MakeReply = "*knuddelknuddel* :))"
'          End Select
'          KIs(KNum).Like = KIs(KNum).Like + 1
'        Case 0
'          Select Case Int(Rnd * 10) + 1
'            Case 1, 2, 3: MakeReply = ":)"
'            Case 4: MakeReply = "*reknuddel*"
'            Case 5: MakeReply = ":)) *reknuddel*"
'            Case 6: MakeReply = "*reknuddel!*"
'            Case 7: MakeReply = "*knuddel*"
'            Case 8: MakeReply = "*wuschel*"
'            Case 9: MakeReply = "Immer dieses Geknuddele :)"
'            Case 10: MakeReply = "Hehe =)"
'          End Select
'          KIs(KNum).Like = KIs(KNum).Like + 1
'        Case -2 To -1
'          Select Case Int(Rnd * 10) + 1
'            Case 1: MakeReply = "*Arme abhack*"
'            Case 2: MakeReply = "scheiß geknuddele" & Choose(Int(Rnd * 2), "..", "")
'            Case 3: MakeReply = "Von DIR lass ich mich nicht knuddeln!"
'            Case 4: MakeReply = ""
'            Case 5: MakeReply = "Pah."
'            Case 6: MakeReply = "No comment."
'            Case 7: MakeReply = "*batsch*"
'            Case 8: MakeReply = "*TRET*"
'            Case 9: MakeReply = ""
'            Case 10: MakeReply = ""
'          End Select
'        Case Is < -2
'          MakeReply = ""
'      End Select
'    Else
'      Select Case KIs(KNum).Like
'        Case Is > 0
'          Select Case Int(Rnd * 14) + 1
'            Case 1, 2: MakeReply = Choose(Int(Rnd * 4) + 1, "*knuddel* ", "*knuddelwuddel* ", "", "*liebhab* ") & "Ich mag dich sehr :))"
'            Case 3, 4: MakeReply = ":)))) *Freu*"
'            Case 5: MakeReply = Choose(Int(Rnd * 3) + 1, "hehe =) ", "hehe, ", "") & "ich freu mich daß du mich magst!!"
'            Case 6: MakeReply = "ich mag dich auch =))"
'            Case 7: MakeReply = "ui! ich mag dich auch, thx :)"
'            Case 8, 9, 10: MakeReply = "hihi :)"
'            Case 11: MakeReply = "*küss*"
'            Case 12: MakeReply = "Hmm, DU bist nett :))"
'            Case 13: MakeReply = ""
'            Case 14: MakeReply = "ARGH! Hab mir eben den Finger verknackst :~(": KIs(KNum).Thema = "malheur"
'          End Select
'        Case 0
'          Select Case Int(Rnd * 10) + 1
'            Case 1, 2, 3: MakeReply = Choose(Int(Rnd * 4) + 1, "Danke, ", "thx, ", "", "thanks! ") & "über dich kann man auch nich meckern :)"
'            Case 4: MakeReply = ":)) Das ist nett!"
'            Case 5: MakeReply = "hihi, thx =)"
'            Case 6: MakeReply = "Daaaanke! :) Ich mag dich auch."
'            Case 7: MakeReply = "Hm, danke! Das hab ich mal gebraucht :)"
'            Case 8: MakeReply = "Also jetzt gates mir gut. :o)"
'            Case 9: MakeReply = "*knuddel*"
'            Case 10: MakeReply = "thx :-) *knuddel*"
'          End Select
'        Case -2 To -1
'          Select Case Int(Rnd * 10) + 1
'            Case 1: MakeReply = "Ach, auf einmal?"
'            Case 2: MakeReply = "Ha, jetzt kommst du angekrochen" & Choose(Int(Rnd * 2), "...", "! Typisch.")
'            Case 3: MakeReply = "Woher der Sinneswandel?"
'            Case 4: MakeReply = ""
'            Case 5: MakeReply = "Pff. Bin schlecht gelaunt."
'            Case 6: MakeReply = "Lass mich in Ruhe!"
'            Case 7: MakeReply = ""
'            Case 8: MakeReply = "Is recht."
'            Case 9: MakeReply = ""
'            Case 10: MakeReply = ""
'          End Select
'        Case Is < -2
'          MakeReply = ""
'      End Select
'    End If
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "[wieso|weshalb|warum|weswegen] [mich|ich] [kiggst*|kickst*|kickest|banst*|bannst*|bannest|kicked|*kickt|gebannt|gebant|gebanned|geband|verbannt|verbanned|kickbanned|kickbannt|gekickbannt|gekickbanned]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Weil du scheisse gebaut hast, ganz einfach."
'      Case 2: MakeReply = "Weil ich dich nicht mag."
'      Case 3: MakeReply = "Ich kann dich halt net leiden."
'      Case 4: MakeReply = "Weil du ein Depp bist ;)"
'      Case 5: MakeReply = "Darum."
'      Case 6: MakeReply = "Nur so. Aus Spaß."
'      Case 7: MakeReply = "Es verschafft mir Befriedigung! ;P"
'      Case 8: MakeReply = "Ich find das geil."
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "thx|danke*|thanks|thank?you|dank?dir|vielen?dank") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1, 2: MakeReply = "np"
'      Case 3, 4: MakeReply = "np!"
'      Case 5: MakeReply = "kein problem"
'      Case 6: MakeReply = "no prob" & IIf(Int(Rnd * 2) = 0, " =)", " :)")
'      Case 7: MakeReply = "np..."
'      Case 8: MakeReply = "kein problem!"
'    End Select
'    If Int(Rnd * 3) + 1 = 1 Then KIs(KNum).Like = KIs(KNum).Like + 1
'    Exit Function
'  End If
'
'  If ((LCase(Subj) = "ich") Or (Subj = "")) And IsInLine(Satz, "wixer|wichser|beschissener|drecks*|arsch*|hure|idiot|lamer|dumm*|spast|*sack|scheiß*|schwanz|fucker|sucker|fick?dich|leck?mich|hal*maul|hal*fresse|hal*klappe") Then
'    Select Case KIs(KNum).Like
'      Case Is > 0
'        Select Case Int(Rnd * 10) + 1
'          Case 1: MakeReply = "hhm? Was ist denn los, daß du mich auf einmal beleidigst?"
'          Case 2: MakeReply = Choose(Int(Rnd * 3) + 1, "Na, ", "??? ", "mmh? ") & "schlecht drauf?"
'          Case 3: MakeReply = "=)"
'          Case 4: MakeReply = "Bist du sauer? " & Choose(Int(Rnd * 2) + 1, "Was hab ich denn gemacht?", "Was ist denn passiert?")
'          Case 5: MakeReply = "Hey! ;o)"
'          Case 6: MakeReply = "selber *G*"
'          Case 7: MakeReply = "Wieso beschimpfst du mich?"
'          Case 8: MakeReply = Choose(Int(Rnd * 5) + 1, "Komisch... ", "Seltsam... ", "hää? ", "Kapier ich nich. ", "") + Choose(Int(Rnd * 2) + 1, "erst", "eben noch") & " total nett und jetzt so" & Choose(Int(Rnd * 2) + 1, "... :/", "?!?")
'          Case 9: MakeReply = "Du Nudel =)"
'          Case 10: MakeReply = "Lamer! :)"
'        End Select
'      Case 0
'        Select Case Int(Rnd * 10) + 1
'          Case 1: MakeReply = "Halt's Maul!"
'          Case 2: MakeReply = "Drecksack!"
'          Case 3: MakeReply = "Wichser!"
'          Case 4: MakeReply = "Spast!"
'          Case 5: MakeReply = "Sackgesicht!!"
'          Case 6: MakeReply = "Ach, leck mich!"
'          Case 7: MakeReply = "Du kannst mich mal!"
'          Case 8: MakeReply = "Halt die Fresse du egomanischer Gefühlskoloss! :o)"
'          Case 9: MakeReply = "pff."
'          Case 10: MakeReply = ""
'        End Select
'      Case -2 To -1
'        Select Case Int(Rnd * 9) + 1
'          Case 1: MakeReply = "hmpf."
'          Case 2: MakeReply = "Halt die Klappe du Arschloch!"
'          Case 3: MakeReply = ""
'          Case 4: MakeReply = "Leck mich am Arsch."
'          Case 5: MakeReply = "Ach, red " & Choose(Int(Rnd * 2) + 1, "", "gefälligst ") & "mit deiner Tastatur" & Choose(Int(Rnd * 3) + 1, " und laß mich in Ruhe!", "!", " oder halt ganz die Klappe.")
'          Case 6: MakeReply = ""
'          Case 7: MakeReply = "Halts Maul!"
'          Case 8: MakeReply = "Klappe!"
'          Case 9: MakeReply = "Laß deinen Scheiß an jemand anderem aus, okay?"
'        End Select
'      Case Is < -2
'        MakeReply = ""
'    End Select
'    KIs(KNum).Like = KIs(KNum).Like - 1
'    Exit Function
'  End If
'  If IsInLine(Satz, "wieso|weshalb|warum|why") Then
'    Select Case Int(Rnd * 11) + 1
'      Case 1: MakeReply = "Warum nicht?"
'      Case 2: MakeReply = "Darum!"
'      Case 3: MakeReply = "Warum? Darum."
'      Case 4: MakeReply = "Du stellst vielleicht Fragen =) Bin ich Moses?"
'      Case 5: MakeReply = "Woher soll ich das denn wissen?"
'      Case 6: MakeReply = "Warum, warum, warum ist die Banane krumm... *g*"
'      Case 7: MakeReply = "Das weiß ich ehrlich gesagt auch nicht ;)"
'      Case 8: MakeReply = "Weißt du das etwa nicht?"
'      Case 9: MakeReply = "Das solltest du aber selbst wissen..."
'      Case 10: MakeReply = "Rate doch mal! Nicht fragen, denken =)"
'      Case 11: MakeReply = "Fragst du das öfters?"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "darum|deshalb") Then
'    Select Case Int(Rnd * 11) + 1
'      Case 1: MakeReply = ";P"
'      Case 2: MakeReply = "Tolle Erklärung ;)"
'      Case 3: MakeReply = "hehe"
'      Case 4: MakeReply = "bll ;P~~"
'      Case 5: MakeReply = "sehr ergiebige Antwort..."
'      Case 6: MakeReply = "hmz"
'      Case 7: MakeReply = "pff..."
'      Case 8: MakeReply = "hmpf"
'      Case 9: MakeReply = "pah =)"
'      Case 10: MakeReply = ""
'      Case 11: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "tut mir leid|sorry|*schuldigung|*schuldige|nich* so gemeint|gomen") Then
'    Select Case KIs(KNum).Like
'      Case Is > 0
'        Select Case Int(Rnd * 8) + 1
'          Case 1: MakeReply = "Kein Problem! =))"
'          Case 2: MakeReply = "Das macht doch nix!"
'          Case 3: MakeReply = "Schwamm drüber! Ich mag dich doch! :))"
'          Case 4: MakeReply = "Macht doch nix!"
'          Case 5: MakeReply = "Schon vergeben und vergessen..."
'          Case 6: MakeReply = "np! =)"
'          Case 7: MakeReply = "Hehe *G* kein Problem!"
'          Case 8: MakeReply = "Das ist schon in Ordnung =))"
'        End Select
'      Case 0
'        KIs(KNum).Thema = "sorry"
'        Select Case Int(Rnd * 8) + 1
'          Case 1: MakeReply = "Ist schon okay..."
'          Case 2: MakeReply = "Du brauchst dich nicht zu entschuldigen!"
'          Case 3: MakeReply = "Das ist kein Problem, du brauchst dich nicht zu entschuldigen..."
'          Case 4: MakeReply = "Macht nix!"
'          Case 5: MakeReply = "Ist schon vergeben und vergessen..."
'          Case 6: MakeReply = "Okay! Entschuldigung akzeptiert.": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 7: MakeReply = "np"
'          Case 8: MakeReply = "Macht nix, das ist in Ordnung."
'        End Select
'      Case -2 To -1
'        Select Case Int(Rnd * 8) + 1
'          Case 1: MakeReply = "Naja... meinetwegen. Entschuldigung akzeptiert.": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 2: MakeReply = "Hmmz. Oki. Ist vergeben! =)": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 3: MakeReply = "Fand ich nich so toll. Aber ok *G* =)": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 4: MakeReply = "Gut, Entschuldigung angenommen! :)": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 5: MakeReply = "Ist schon vergessen...": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 6: MakeReply = "Oki. Entschuldigung akzeptiert.": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 7: MakeReply = "Naja...": KIs(KNum).Thema = "sorry"
'          Case 8: MakeReply = "Na gut, wenn du meinst.": KIs(KNum).Like = KIs(KNum).Like + 1
'        End Select
'      Case Is < -2
'        KIs(KNum).Thema = "sorry"
'        Select Case Int(Rnd * 8) + 1
'          Case 1: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 2: MakeReply = "pah. glaub ich dir nich."
'          Case 3: MakeReply = ""
'          Case 4: MakeReply = "Ach, geh mir nich auf " & Choose(Int(Rnd * 3) + 1, "die Senkel.", "den Geist.", "die Nerven...")
'          Case 5: MakeReply = "Na gut, akzeptiert.": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 6: MakeReply = "Ok. Entschuldigung angenommen.": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 7: MakeReply = "": KIs(KNum).Like = KIs(KNum).Like + 1
'          Case 8: MakeReply = ""
'        End Select
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "bot|robot") Then
'    Select Case Int(Rnd * 10) + 1
'      Case 1: MakeReply = "Du denkst wohl, daß ich ein Programm bin, hm?"
'      Case 2: MakeReply = "Warum sollte ich ein Bot sein?"
'      Case 3: MakeReply = "Wie kommst du darauf?"
'      Case 4: MakeReply = "Ich denke nicht *G*"
'      Case 5: MakeReply = "Ein BOT? Was ist das?"
'      Case 6: MakeReply = "Ich würde gerne wissen warum ich ein bot sein sollte..."
'      Case 7: MakeReply = "Wieso? Magst du bots?"
'      Case 8: MakeReply = "Hm? | Wieso ein bot?"
'      Case 9: MakeReply = "Ich bin doch kein Bot! :) | Wie kommst du denn darauf?"
'      Case 10: MakeReply = "Bot? Warum? | Redest du nicht gerne mit Computerprogrammen?"
'      Case 11: MakeReply = "Ich? Ein Bot? | Haha =)"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "schade|schad|doof") Then
'    Select Case Int(Rnd * 6) + 1
'      Case 1: MakeReply = "naja, man kann nicht alles haben..."
'      Case 2: MakeReply = "naja egal"
'      Case 3: MakeReply = "findest du?": KIs(KNum).Thema = "really?"
'      Case 4: MakeReply = "hm. wieso findest du das schade?"
'      Case 5: MakeReply = "hm.."
'      Case 6: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "was?geht|wie?gayt*|wie?gehts|wie?geht?s|wie?geht?es|wie?fühlst|alles?fit") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Also mir gehts blendend =)"
'      Case 2: MakeReply = "Mir gehts gut, danke der Nachfrage! | Und wie stehts mit dir?": KIs(KNum).Thema = "howareyou"
'      Case 3: MakeReply = "Geht so, danke! Und dir?": KIs(KNum).Thema = "howareyou"
'      Case 4: MakeReply = "Och, geht so! | Und dir?": KIs(KNum).Thema = "howareyou"
'      Case 5: MakeReply = "Klasse! Und wie gehts dir?": KIs(KNum).Thema = "howareyou"
'      Case 6: MakeReply = "Ich weiß nicht so recht... ich lagge heute ziemlich oft =)"
'      Case 7: MakeReply = "Naja, ich kann nicht klagen - und wie geht's dir?": KIs(KNum).Thema = "howareyou"
'      Case 8: MakeReply = "Nicht schlecht! Und wie fühlst du dich so?": KIs(KNum).Thema = "howareyou"
'      Case 9: MakeReply = "Mir gehts super :) | Bei dir auch alles klar?": KIs(KNum).Thema = "howareyou"
'    End Select
'    Exit Function
'  End If
'  If (InStr(Satz, "?") > 0) And IsInLine(Satz, "alles?klar|alles?korrekt|alles?senkrecht") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Jap, alles bestens :) | Und bei dir?": KIs(KNum).Thema = "howareyou"
'      Case 2: MakeReply = "Mir gehts gut, thx! | Und dir?": KIs(KNum).Thema = "howareyou"
'      Case 3: MakeReply = "Alles klar hier =) Und wie geht's dir?": KIs(KNum).Thema = "howareyou"
'      Case 4: MakeReply = "Ich fühl mich super. Und du?": KIs(KNum).Thema = "howareyou"
'      Case 5: MakeReply = "Mir gates gut ;) Und wie gehts dir?": KIs(KNum).Thema = "howareyou"
'      Case 6: MakeReply = "Scheiß Netsplits laufend... ansonsten gehts mir gut =)"
'      Case 7: MakeReply = "Kann nich meckern - und wie geht's dir?": KIs(KNum).Thema = "howareyou"
'      Case 8: MakeReply = "Naja, mittelmäßig... und wie fühlst du dich so?": KIs(KNum).Thema = "howareyou"
'      Case 9: MakeReply = "Mir gehts super :) Bei dir auch alles klar?": KIs(KNum).Thema = "howareyou"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "stimmt|richtig|right|genau|bingo") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = ":)"
'      Case 2: MakeReply = "hehe, wußte ich's doch."
'      Case 3: MakeReply = "ich wußte es..."
'      Case 4: MakeReply = "hehe"
'      Case 5: MakeReply = ":))"
'      Case 6: MakeReply = "hab ich mir gedacht :o)"
'      Case 7: MakeReply = ""
'      Case 8: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "schlecht|mies|nicht?gut|nicht?so?gut|schrecklich|schlimm|unwohl") And (Right(Satz, 1) <> "?") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Oh! Warum das?"
'      Case 2: MakeReply = "Hhhm... gar nicht gut..."
'      Case 3: MakeReply = "Oh, das ist schade! Warum denn?"
'      Case 4: MakeReply = ""
'      Case 5: MakeReply = "Na komm, erzähl mir was dich bedrückt!": KIs(KNum).Thema = "helpme"
'      Case 6: MakeReply = "Du kannst mir ruhig alles erzählen, was du auf dem Herzen hast!": KIs(KNum).Thema = "helpme"
'      Case 7: MakeReply = "Hast du Liebeskummer oder sowas?"
'      Case 8: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "[mir gehts|mir geht?s|ich fühl*|es geht mir|] [spitze|gut|super|toll|klasse|geht?so|es?geht|nicht?schlecht|wunderbar]") Then
'    Select Case Int(Rnd * 9) + 1
'      Case 1: MakeReply = Choose(Int(Rnd * 4) + 1, "na, ", "na ", "", "also ") & "das hört man doch gern!"
'      Case 2: MakeReply = Choose(Int(Rnd * 3) + 1, "Hehe, ", "super! ", "") & "freut mich!"
'      Case 3: MakeReply = Choose(Int(Rnd * 3) + 1, "hey, das ", "das ", "") & "freut mich :)"
'      Case 4: MakeReply = "sehr schön :))"
'      Case 5: MakeReply = "ah, freut mich das zu hören!"
'      Case 6: MakeReply = "hehe"
'      Case 7: MakeReply = "Gut gut... und, was gibts zu bereden?"
'      Case 8: MakeReply = "Ah gut, ich mag positive Leute =)"
'      Case 9: MakeReply = "Freut mich für dich!"
'    End Select
'    Exit Function
'  End If
'
'  If (((Trim(Replace(Replace(Satz, "?", ""), "!", "")) = "") Or IsInLine(Satz, "hä*|hm*|mh*|mmh")) And (Len(SSatz) < 10)) And (InStr(Satz, "?") > 0) Then
'    Select Case Int(Rnd * 6) + 1
'      Case 1: MakeReply = "Was ist los, hast du was nich kapiert?"
'      Case 2: MakeReply = "?"
'      Case 3: MakeReply = "Was ist" & Choose(Int(Rnd * 3) + 1, " los", " passiert", "") & "?!"
'      Case 4: MakeReply = "??"
'      Case 5: MakeReply = Choose(Int(Rnd * 3) + 1, "hä?", "hmm?", "mh?")
'      Case 6: MakeReply = ""
'      Case 7: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "auch|ebenfalls|genauso") And (Right(Satz, 1) <> "?") Then
'    Select Case Int(Rnd * 6) + 1
'      Case 1: MakeReply = "Hey, da haben wir ja was gemeinsam =)"
'      Case 2: MakeReply = "*hehe* Ich freue mich immer wenn ich etwas mit jemandem gemeinsam habe..."
'      Case 3: MakeReply = "Faszinierend! Glaubst du, daß diese Gemeinsamkeit ein Zufall ist?"
'      Case 4: MakeReply = "Ich denke das ist ein gutes Zeichen, wenn wir was gemeinsam haben!"
'      Case 5: MakeReply = "Willkommen, Partner =)"
'      Case 6: MakeReply = "Tja, da scheinen wir wohl übereinzustimmen, hm? | *g*"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "weil") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Denkst du das ist Begründung genug?": KIs(KNum).Thema = "really?"
'      Case 2: MakeReply = "Glaubst du das wirklich?": KIs(KNum).Thema = "really?"
'      Case 3: MakeReply = "Kann es sein, daß du da falsch liegst?": KIs(KNum).Thema = "really?"
'      Case 4: MakeReply = "Hast du dir diese Begründung lange überlegt oder war das eher spontan?"
'      Case 5: MakeReply = "Ah, ich verstehe!"
'      Case 6: MakeReply = "axo"
'      Case 7: MakeReply = "verstehe..."
'      Case 8: MakeReply = "Ich kann diese Erklärung nicht ganz nachvollziehen..."
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "[gibst|krieg|bekomme|machst du] [ich|mich|mir] [op|@|operator|chanop]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Okay. | Oder nee, doch nicht."
'      Case 2: MakeReply = "Warum sollte ich? *lach*"
'      Case 3: MakeReply = "Ich seh's überhaupt nicht ein."
'      Case 4: MakeReply = "Hör auf zu betteln. | Dadurch kriegst du erst recht keinen."
'      Case 5: MakeReply = "Vergisses. Du kriegst von mir nix! ;-)"
'      Case 6: MakeReply = "Hahaha. 'Oppt mich, ich bin toll!' =))"
'      Case 7: MakeReply = "LOL. Klar, du mich auch :P"
'      Case 8: MakeReply = "Du kannst mich mal am Op lecken! ;-)"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "[sag|[wie|was|welches] [ist|is|iss]] [ident*command|ident*befehl|ident*kommando]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Frag bitte einen Owner. Ich sage dazu nix :)"
'      Case 2: MakeReply = "Sag ich nicht. | Frag einen Owner."
'      Case 3: MakeReply = "Was für ein Ident-Befehl?!"
'      Case 4: MakeReply = "Glaubst du ich binde dir das auf die Nase? :-)"
'      Case 5: MakeReply = "Nee, das sag ich dir nicht ;)"
'      Case 6: MakeReply = "Es liegt mir auf der Tastatur... | Nee, ich komm nicht drauf. Sorry. ;P"
'      Case 7: MakeReply = "Hihi. Von mir erfährst du das nicht, vergisses! =)"
'      Case 8: MakeReply = "*schweig*"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "was [machst*|tust*]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "nix besonderes... rumidlen, mit dir chatten... und du?"
'      Case 2: MakeReply = "Ich mache nichts großartiges im Moment."
'      Case 3: MakeReply = "Mir gleich in die Hose zum Beispiel!! | bin mal weg!"
'      Case 4: MakeReply = "Das willst du nicht wissen ;o)"
'      Case 5: MakeReply = "Ich sitze hier rum und langweile mich."
'      Case 6: MakeReply = "Das frage ich mich auch. Nix aufregendes jedenfalls."
'      Case 7: MakeReply = "Abraven ;) ich liebe mp3's"
'      Case 8: MakeReply = "mich langweilen..."
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "was [[magst|liebst|findest [gut|toll]] du|[magste|liebste|findeste [gut|toll]]]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "nix besonderes... rumidlen, mit dir chatten... und du?"
'      Case 2: MakeReply = "Ich mache nichts großartiges im Moment."
'      Case 3: MakeReply = "Mir gleich in die Hose zum Beispiel!! | bin mal weg!"
'      Case 4: MakeReply = "Das willst du nicht wissen ;o)"
'      Case 5: MakeReply = "Ich sitze hier rum und langweile mich."
'      Case 6: MakeReply = "Das frage ich mich auch. Nix aufregendes jedenfalls."
'      Case 7: MakeReply = "Abraven ;) ich liebe mp3's"
'      Case 8: MakeReply = "mich langweilen..."
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "[was [haste|hast du|sind|deine|für]|welche|bei dir|was deine|welches deine] [hobby*|hobbi*|hobi*|hoby*]") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Also ich hab als Hobbies chatten, programmieren und radfahren... nix besonderes also ;) | Was hast du so für Hobbies?"
'      Case 2: MakeReply = "hhm. Meine hobbys sind anime, fernsehen und chatten."
'      Case 3: MakeReply = "Freunde treffen, Parties, alles was fun ist halt... Wie sind deine hobbies?"
'      Case 4: MakeReply = "Nix besonderes... chatten, computerspielen und musik hören :) | Und was hast du für Hobbies?"
'      Case 5: MakeReply = "Ich lese gerne, hör gern Musik und spiele für mein Leben gern Nibbles. | Und du?"
'      Case 6: MakeReply = "Oh, da fällt mir nicht viel ein. Vielleicht programmieren, aber das kann ich nicht so doll."
'      Case 7: MakeReply = "Discogehen und abraven macht mir echt spaß! Was magst du so?"
'      Case 8: MakeReply = "Ich hab keine Hobbies :~( | Bin einfach unfähig..."
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "was") Then
'    If (Subj = "ich") Or (InStr(LCase(Subj), "mein") > 0) Then
'      If MakeReply = "" Then
'        Select Case Int(Rnd * 8) + 1
'          Case 1: MakeReply = "Sorry, keine ahnung!"
'          Case 2: MakeReply = "Das weiß ich nicht..."
'          Case 3: MakeReply = "Das solltest du selbst wissen!!"
'          Case 4: MakeReply = "Sag ich nicht ;o)"
'          Case 5: MakeReply = "Puh, keine Ahnung..."
'          Case 6: MakeReply = "weiß ich nicht"
'          Case 7: MakeReply = "Das hab ich mich auch schonmal gefragt..."
'          Case 8: MakeReply = "k.A."
'        End Select
'      End If
'      Exit Function
'    Else
'      If Right(Satz, 1) = "?" Then
'        Select Case Int(Rnd * 8) + 1
'          Case 1: MakeReply = "Also das kann ich dir echt nicht sagen!"
'          Case 2: MakeReply = "Eigentlich eine gute Frage..."
'          Case 3: MakeReply = "Also das solltest du selbst wissen!"
'          Case 4: MakeReply = "Bevor ich darauf antworte will ich mit meinem Anwalt sprechen ;)"
'          Case 5: MakeReply = "Das weiß ich nicht!"
'          Case 6: MakeReply = "Frag mich was leichteres!"
'          Case 7: MakeReply = "Ich frage mich das manchmal auch."
'          Case 8: MakeReply = "Keine Ahnung =)"
'        End Select
'        Exit Function
'      End If
'    End If
'  End If
'  If IsInLine(Satz, "eigentlich") And Right(Satz, 1) <> "?" And Int(Rnd * 3) = 1 Then
'    Select Case Int(Rnd * 4) + 1
'      Case 1: MakeReply = "'Eigentlich'? Also nicht immer, oder wie soll ich das verstehen?"
'      Case 2: MakeReply = "Du bist dir nicht sicher, oder?"
'      Case 3: MakeReply = "Warum 'eigentlich'? | Ist das nicht immer so?"
'      Case 4: MakeReply = "ah ja :)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "cu|bai bai|bye|c-ya|cya|c u|ciao*|c ya|machs gut|n8|goodni*|tschau") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "cu! machs gut!"
'      Case 2: MakeReply = "see juu =)"
'      Case 3: MakeReply = "bai bai baby =)"
'      Case 4: MakeReply = "bye!"
'      Case 5: MakeReply = "CU"
'      Case 6: MakeReply = "cu!!!"
'      Case 7: MakeReply = "bye und good n8! =)"
'      Case 8: MakeReply = "ciao!"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "okay|oki|ok|okie*") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "gut, gut =)"
'      Case 2: MakeReply = "jep, okay"
'      Case 3: MakeReply = "ok."
'      Case 4: MakeReply = "okay"
'      Case 5: MakeReply = "ist in ordnung ;)"
'      Case 6: MakeReply = "oki"
'      Case 7: MakeReply = "okie-dokie."
'      Case 8: MakeReply = "gut so :)"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "echt|wirklich|nich* wahr|tats*chlich|ehrlich") And Right(Satz, 1) = "?" Then
'    Select Case Int(Rnd * 13) + 1
'      Case 1: MakeReply = "jau"
'      Case 2: MakeReply = "jap"
'      Case 3: MakeReply = "ja"
'      Case 4: MakeReply = "jo :)"
'      Case 5: MakeReply = "*nick*"
'      Case 6: MakeReply = "wirklich :)"
'      Case 7: MakeReply = "hehe, ja"
'      Case 8: MakeReply = "ja!"
'      Case 9: MakeReply = "hm, nee"
'      Case 10: MakeReply = "eigentlich nicht..."
'      Case 11: MakeReply = "nee"
'      Case 12: MakeReply = "jein *G*"
'      Case 13: MakeReply = "hehe, nö ;)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "cool|goil|geil|krass|kraß") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Immer diese Jugendliche Sprache... ""cool"", ""geil"", naja =)"
'      Case 2: MakeReply = "kraß =)"
'      Case 3: MakeReply = ";o)"
'      Case 4: MakeReply = "=)"
'      Case 5: MakeReply = "*g*"
'      Case 6: MakeReply = "hehe..."
'      Case 7: MakeReply = "tja *hehe*..."
'      Case 8: MakeReply = "jupjup =)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "konkret|ultra*|alda|dragan") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Ey alda, konkret!"
'      Case 2: MakeReply = "Ultrakorrekt, Dragan!"
'      Case 3: MakeReply = "=)"
'      Case 4: MakeReply = "Konkret!"
'      Case 5: MakeReply = "Dragaaan... Aldaaa :)"
'      Case 6: MakeReply = "Ultrakrass Dragan | :o)"
'      Case 7: MakeReply = "Krass ;)"
'      Case 8: MakeReply = "hehe =)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "mist*|scheiße|scheisse|fuck|verdammt") Then
'    Select Case Int(Rnd * 9) + 1
'      Case 1: MakeReply = "Was ist denn los?"
'      Case 2: MakeReply = "Ist irgendwas passiert?"
'      Case 3: MakeReply = "Geht was nich?"
'      Case 4: MakeReply = "Probleme?"
'      Case 5: MakeReply = "Hmm?"
'      Case 6: MakeReply = "was ist los?"
'      Case 7: MakeReply = "alles okay?!"
'      Case 8: MakeReply = ""
'      Case 9: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "[lebst|bist] du noch|noch da") Then
'    Select Case Int(Rnd * 9) + 1
'      Case 1: MakeReply = "ähm | ja, wieso?"
'      Case 2: MakeReply = "bin da! was gibts?"
'      Case 3: MakeReply = "jo, hi"
'      Case 4: MakeReply = "wieso sollte ich weg sein? *G*"
'      Case 5: MakeReply = "immer doch!"
'      Case 6: MakeReply = "wie immer... ja"
'      Case 7: MakeReply = "ja, bin anwesend"
'      Case 8: MakeReply = "*kopfnick*"
'      Case 9: MakeReply = "jep! | wieso?"
'    End Select
'    Exit Function
'  End If
'
'  'Du tippst so schnell
'  If IsInLine(Satz, "[bist|du tip*|du schreib*] [schnell*|flott]") Then
'    Select Case Int(Rnd * 9) + 1
'      Case 1: MakeReply = "das ist doch net schnell ;)"
'      Case 2: MakeReply = "was?! | ich dachte immer ich wäre so lahm *G*"
'      Case 3: MakeReply = "naja, zehn-finger-suchsystem halt ;-)"
'      Case 4: MakeReply = "ich werde mich mal bemühen, langsamer zu schreiben ;P"
'      Case 5: MakeReply = "ich schreib doch net schnell..."
'      Case 6: MakeReply = "naja, ich finds langsam =)"
'      Case 7: MakeReply = "schnell? echt?"
'      Case 8: MakeReply = "bin ich wirklich schnell?"
'      Case 9: MakeReply = "das ist standard *G*"
'    End Select
'    Exit Function
'  End If
'
'  'Detect widerworte ;)
'  If IsInLine(Satz, "doch|dochdoch|dooch") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Echt? Hätt ich nicht gedacht."
'      Case 2: MakeReply = "Wirklich? | Hätte ich nicht geglaubt."
'      Case 3: MakeReply = "Ja? | Na gut =)"
'      Case 4: MakeReply = "Hm. Wenn du meinst."
'      Case 5: MakeReply = "Ehrlich? | Hätt ich nicht für möglich gehalten."
'      Case 6: MakeReply = "Doch? Hm. Na gut ;-)"
'      Case 7: MakeReply = "Ja? Faszinierend."
'      Case 8: MakeReply = "Doch? *staun*"
'    End Select
'    Exit Function
'  End If
'
'  'Detect general questions
'  If IsInLine(Satz, "wie") And Right(Satz, 1) = "?" Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Also das kann ich dir echt nicht sagen!"
'      Case 2: MakeReply = "Eigentlich eine gute Frage..."
'      Case 3: MakeReply = "Also das solltest du selbst wissen!"
'      Case 4: MakeReply = "Bevor ich darauf antworte will ich mit meinem Anwalt sprechen ;)"
'      Case 5: MakeReply = "Das weiß ich nicht!"
'      Case 6: MakeReply = "Frag mich was leichteres!"
'      Case 7: MakeReply = "Ich frage mich das manchmal auch."
'      Case 8: MakeReply = "Keine Ahnung =)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "wo") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "Auf dem Mond."
'      Case 2: MakeReply = "In München"
'      Case 3: MakeReply = "Unter der Dusche..."
'      Case 4: MakeReply = "Im Bett =)"
'      Case 5: MakeReply = "Weiß ich nicht :)"
'      Case 6: MakeReply = "Das willst du nicht wissen ;)"
'      Case 7: MakeReply = "Wo... hhm... gute Frage!"
'      Case 8: MakeReply = "Irgendwo wo's Flaschenöffner gibt ^.^"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "wann") Then
'    Select Case Int(Rnd * 7) + 1
'      Case 1: MakeReply = "Morgen wahrscheinlich..."
'      Case 2: MakeReply = "Gestern!"
'      Case 3: MakeReply = "Ich hab keine Uhr an..."
'      Case 4: MakeReply = "In 15 Minuten!"
'      Case 5: MakeReply = "Och, das ist nicht lang her..."
'      Case 6: MakeReply = "DAS wüßte ich auch gern!"
'      Case 7: MakeReply = "Wann? Man verliert so das Zeitgefühl, wenn man so lang online ist wie ich *G*"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "keine?ahnung | [weiss|weiß] [nicht|nich|net]") Then
'    Select Case Int(Rnd * 12) + 1
'      Case 1:  MakeReply = "Oh, das ist nicht gut... man sollte immer etwas wissen."
'      Case 2:  MakeReply = "Wirklich absolut keine Ahnung?": KIs(KNum).Thema = "really?"
'      Case 3:  MakeReply = "Also das muß man doch wissen... tsk tsk tsk"
'      Case 4:  MakeReply = "Ich glaube irgendwie du verheimlichst mir etwas... ;)"
'      Case 5:  MakeReply = "Weißt du es wirklich nicht?": KIs(KNum).Thema = "really?"
'      Case 6:  MakeReply = "Verdrängst du etwas oder weißt du wirklich nix?": KIs(KNum).Thema = "really?"
'      Case 7:  MakeReply = "Warum nicht? Hast du es vergessen?"
'      Case 8:  MakeReply = "Denk doch noch einmal scharf nach!"
'      Case 9:  MakeReply = "Du mußt das doch wissen!"
'      Case 10: MakeReply = "Hast du's vergessen?"
'      Case 11: MakeReply = "Das könnten erste Anzeichen von Alzheimer sein... ;)"
'      Case 12: MakeReply = "Sieht ganz nach einer geistigen Blockade aus! ;P"
'    End Select
'    Exit Function
'  End If
'
'  If IsInLine(Satz, "nein|nee|nö|nope|nicht") Then
'    If KIs(KNum).Thema = "yousure?" Then
'      KIs(KNum).Thema = ""
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "interessant..."
'        Case 2: MakeReply = "aahja"
'        Case 3: MakeReply = "Verstehe."
'        Case 4: MakeReply = "Aha."
'        Case 5: MakeReply = "i c"
'        Case 6: MakeReply = "verstehe"
'        Case 7: MakeReply = "ic"
'        Case 8: MakeReply = "hehe"
'        Case 9: MakeReply = "aha"
'        Case 10: MakeReply = ""
'      End Select
'      Exit Function
'    End If
'    KIs(KNum).Thema = "yousure?"
'    Select Case Int(Rnd * 10) + 1
'      Case 1: MakeReply = "Warum nicht?"
'      Case 2: MakeReply = "Bist du dir da sicher?": KIs(KNum).Thema = "really?"
'      Case 3: MakeReply = "Du bist ein bischen negativ, findest du nicht?"
'      Case 4: MakeReply = "Nicht?": KIs(KNum).Thema = "really?"
'      Case 5: MakeReply = "Wirklich nicht?": KIs(KNum).Thema = "really?"
'      Case 6: MakeReply = "achso"
'      Case 7: MakeReply = "verstehe..."
'      Case 8: MakeReply = ""
'      Case 9: MakeReply = "Macht dir das Angst?"
'      Case 10: If Int(Rnd * 3) = 1 Then MakeReply = "hm, wie alt bist du eigentlich?" Else MakeReply = ""
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "ja|jo|jup*|jau|jep") Then
'    If KIs(KNum).Thema = "yousure?" Then
'      KIs(KNum).Thema = ""
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "Klingt einleuchtend."
'        Case 2: MakeReply = "ic"
'        Case 3: MakeReply = "Verstehe."
'        Case 4: MakeReply = "Aha."
'        Case 5: MakeReply = "interessant..."
'        Case 6: MakeReply = "hehe *G*"
'        Case 7: MakeReply = ""
'        Case 8: MakeReply = "Hehe =)"
'        Case 9: MakeReply = "Ah, ok..."
'        Case 10: MakeReply = "Hehe :)"
'      End Select
'      Exit Function
'    End If
'    KIs(KNum).Thema = "yousure?"
'    Select Case Int(Rnd * 10) + 1
'      Case 1: MakeReply = "Und warum?"
'      Case 2: MakeReply = "Bist du dir sicher?": KIs(KNum).Thema = "really?"
'      Case 3: MakeReply = "Klingt gut :)"
'      Case 4: MakeReply = "Ja?"
'      Case 5: MakeReply = "Wirklich?": KIs(KNum).Thema = "really?"
'      Case 6: MakeReply = "Mir ist voll langweilig... | ich glaub ich leg mich schlafen."
'      Case 7: MakeReply = "Wieso ja?"
'      Case 8: MakeReply = ""
'      Case 9: MakeReply = "Findest du das gut?"
'      Case 10: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  'Short words
'  If IsInLine(Satz, "ey") And (Len(Satz) < 12) Then
'    Select Case Int(Rnd * 7) + 1
'      Case 1: MakeReply = "ey was ey? ;P"
'      Case 2: MakeReply = "Na, was ist los?"
'      Case 3: MakeReply = "Ja, ey? =)"
'      Case 4: MakeReply = "Ey boah ey was gibts ey?"
'      Case 5: MakeReply = "Ey ja bitte? ;-)"
'      Case 6: MakeReply = "Was ey?"
'      Case 7: MakeReply = "Was is?"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "hey") And (Len(Satz) < 12) Then
'    Select Case Int(Rnd * 7) + 1
'      Case 1: MakeReply = "Was hey?"
'      Case 2: MakeReply = "Hey was?"
'      Case 3: MakeReply = "Ja?"
'      Case 4: MakeReply = "Hey, was ist los? =)"
'      Case 5: MakeReply = "Ja bitte? :-)"
'      Case 6: MakeReply = "Was ist los?"
'      Case 7: MakeReply = "Hmm? | Was gibts?"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "aha") And (Len(Satz) < 12) Then
'    Select Case Int(Rnd * 7) + 1
'      Case 1, 2, 3, 4: MakeReply = ""
'      Case 5: MakeReply = "was 'aha'? | war das nicht ausreichend erklärt? :)"
'      Case 6: MakeReply = "wieso ""aha""? =) | glaubst du's nicht? :)"
'      Case 7: MakeReply = "war das nicht ausreichend erklärt? ;)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(Satz, "soso|ah*ja|achja|ahso|achso|ic|i c") And (Len(Satz) < 12) Then
'    Select Case Int(Rnd * 7) + 1
'      Case 1, 2, 3, 4: MakeReply = ""
'      Case 5: MakeReply = ":)"
'      Case 6: MakeReply = "=)"
'      Case 7: MakeReply = ":-)"
'    End Select
'    Exit Function
'  End If
'
'  'No matches found - were smileys used?
'  If IsInLine(SatSatz, "hehe|hihi|haha|g|gg|grins|:-)|;-)|:o)|;o)|:>|;>|:)|;)|=)|^_^|^.^") Then
'    Select Case Int(Rnd * 9) + 1
'      Case 1: MakeReply = "*grins*"
'      Case 2: MakeReply = "*hehe*"
'      Case 3: MakeReply = ";o)"
'      Case 4: MakeReply = "*G*"
'      Case 5: MakeReply = "*g*"
'      Case 6: MakeReply = "hehe..."
'      Case 7: MakeReply = ";)"
'      Case 8: MakeReply = "lol"
'      Case 9: MakeReply = ":)"
'    End Select
'    Exit Function
'  End If
'  If IsInLine(SatSatz, "lol|rofl*|rotfl*|lach") Then
'    Select Case Int(Rnd * 8) + 1
'      Case 1: MakeReply = "*eg* ;)"
'      Case 2: MakeReply = "=)"
'      Case 3: MakeReply = "*gg*"
'      Case 4: MakeReply = ""
'      Case 5: MakeReply = ";o)"
'      Case 6: MakeReply = ""
'      Case 7: MakeReply = "hehehe"
'      Case 8: MakeReply = ""
'    End Select
'    Exit Function
'  End If
'
'  'Absolutely no keywords matched - use default replies
'  Select Case Right(SavSatz, 1)
'  Case "?"
'    If IsInLine(Satz, "oder") Then
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "Ich kann mich nich entscheiden *G*"
'        Case 2: MakeReply = "Ich bin mir nicht sicher..."
'        Case 3: MakeReply = "Ersteres."
'        Case 4: MakeReply = "letzteres ;)"
'        Case 5: MakeReply = "Letzteres."
'        Case 6: MakeReply = "Komische Frage :)"
'        Case 7: MakeReply = "? hm, k.A."
'        Case 8: MakeReply = "ersteres. | oder nee, doch das andere ;)"
'        Case 9: MakeReply = "weiß ich nicht!"
'        Case 10: MakeReply = "das erste."
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "[bist|hast] du") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = "Wieso fragst du?"
'        Case 2: MakeReply = "Ich? Wieso?"
'        Case 3: MakeReply = "Ja :)"
'        Case 4: MakeReply = "Hm - was wäre wenn?"
'        Case 5: MakeReply = "Wie kommst du denn darauf?!"
'        Case 6: MakeReply = "nee..."
'        Case 7: MakeReply = "jo"
'        Case 8: MakeReply = "ich??"
'        Case 9: MakeReply = "wieso?"
'      End Select
'      Exit Function
'    End If
'    If IsInLine(SavSatz, "[magst|willst|denkst|findest] du") Then
'      Select Case Int(Rnd * 13) + 1
'        Case 1: MakeReply = "Ja klar, wieso auch nicht?"
'        Case 2: MakeReply = "Jap!"
'        Case 3: MakeReply = "jo :)"
'        Case 4: MakeReply = "eher weniger"
'        Case 5: MakeReply = "nee"
'        Case 6: MakeReply = "nö."
'        Case 7: MakeReply = "ja"
'        Case 8: MakeReply = "Was du für Sachen fragst ;)"
'        Case 9: MakeReply = "nein"
'        Case 10: MakeReply = "hm. | keine ahnung."
'        Case 11: MakeReply = "doch, schon..."
'        Case 12: MakeReply = "ja :)"
'        Case 13: MakeReply = "jo"
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "du") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = "Warum ich?"
'        Case 2: MakeReply = "Ich? Warum das denn?"
'        Case 3: MakeReply = "Würde dich das stören?"
'        Case 4: MakeReply = "Was wäre wenn?"
'        Case 5: MakeReply = "Hhm... ich will erstmal meinen Anwalt sprechen, bevor ich darauf antworte =)"
'        Case 6: MakeReply = "Fändest du das gut?"
'        Case 7: MakeReply = "Kann schon sein..."
'        Case 8: MakeReply = "Möglich wäre es."
'        Case 9: MakeReply = "Frag mich was leichteres!"
'      End Select
'      Exit Function
'    End If
'    If IsInLine(Satz, "da") Then
'      Select Case Int(Rnd * 9) + 1
'        Case 1: MakeReply = "Jep!"
'        Case 2: MakeReply = "Jau..."
'        Case 3: MakeReply = "jo"
'        Case 4: MakeReply = "ja"
'        Case 5: MakeReply = "Ja."
'        Case 6: MakeReply = "Jau!"
'        Case 7: MakeReply = "Immer doch!"
'        Case 8: MakeReply = "Klar, wieso?"
'      End Select
'      Exit Function
'    End If
'    If ParamCount(SSatz) = 1 Then
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "hmm?"
'        Case 2: MakeReply = "Was gibts?"
'        Case 3: MakeReply = "Ja? :)"
'        Case 4: MakeReply = ""
'        Case 5: MakeReply = "jo"
'        Case 6: MakeReply = "mh?"
'        Case 7: MakeReply = ""
'        Case 8: MakeReply = ""
'        Case 9: MakeReply = "bin mal wech"
'        Case 10: MakeReply = ""
'      End Select
'    Else
'      If Len(Satz) > 5 Then
'        Select Case Int(Rnd * 13) + 1
'          Case 1: MakeReply = "Hhm... das darfst du mich nicht fragen!"
'          Case 2: MakeReply = "Eigentlich eine gute Frage..."
'          Case 3: MakeReply = "Würde dich das stören?"
'          Case 4: MakeReply = "Was wäre wenn?"
'          Case 5: MakeReply = "Also das solltest du selbst wissen!"
'          Case 6: MakeReply = "Fändest du das gut?"
'          Case 7: MakeReply = "Ich weiß nicht was ich darauf antworten sollte..."
'          Case 8: MakeReply = "Möglich wäre es."
'          Case 9: MakeReply = "Das weiß ich nicht!"
'          Case 10: MakeReply = "Frag mich was leichteres!"
'          Case 11: MakeReply = "Ich frage mich das manchmal auch."
'          Case 12: MakeReply = "Würde dich das froh machen?"
'          Case 13: MakeReply = "*pff* das weiß ich auch nicht =)"
'        End Select
'      Else
'        MakeReply = ""
'      End If
'    End If
'  Case Else
'    If IsInLine(Satz, "ich?will|ich?möchte|ich?hätte?gern|ich?haette?gern") And (Subj = "" Or NewOne = "" Or Praed = "") Then
'      Select Case Int(Rnd * 7) + 1
'        Case 1: MakeReply = "Hast du oft diesen Wunsch?"
'        Case 2: MakeReply = "Glaubst du daß ich dir da helfen kann?"
'        Case 3: MakeReply = "Willst du das wirklich?"
'        Case 4: MakeReply = "Bist du dir sicher, daß du das willst?"
'        Case 5: MakeReply = "Warum willst du das denn?"
'        Case 6: MakeReply = "Gibt es einen Grund für diesen Wunsch von dir?"
'        Case 7: MakeReply = "Was würde es an deinem Leben verändern, wenn dieser Wunsch in Erfüllung gehen würde?"
'      End Select
'      Exit Function
'    End If
'    If Subj = "" Or NewOne = "" Or Praed = "" Then
'      If ParamCount(SSatz) = 1 Then
'        Select Case Int(Rnd * 10) + 1
'          Case 1: MakeReply = "hm?"
'          Case 2: MakeReply = "Häh?"
'          Case 3: MakeReply = ""
'          Case 4: MakeReply = "Was gibts?"
'          Case 5: MakeReply = "mmh?"
'          Case 6: MakeReply = "Was willst du mir damit sagen?"
'          Case 7: MakeReply = "ahja!"
'          Case 8: MakeReply = "eh?"
'          Case 9: MakeReply = ""
'          Case 10: MakeReply = "wus?"
'        End Select
'        Exit Function
'      End If
'      If Len(Satz) > 4 Then
'        Select Case Int(Rnd * 10) + 1
'          Case 1: MakeReply = "Kannst du das mal erläutern?": KIs(KNum).Thema = "helpme"
'          Case 2: MakeReply = "Interessant!"
'          Case 3: MakeReply = "Wie?": KIs(KNum).Thema = "helpme"
'          Case 4: MakeReply = "Hä?": KIs(KNum).Thema = "helpme"
'          Case 5: MakeReply = ""
'          Case 6: MakeReply = "Wie bitte?!": KIs(KNum).Thema = "helpme"
'          Case 7: MakeReply = "Was gibts?": KIs(KNum).Thema = "helpme"
'          Case 8: MakeReply = ""
'          Case 9: MakeReply = ""
'          Case 10: MakeReply = "?!?": KIs(KNum).Thema = "helpme"
'        End Select
'      Else
'        MakeReply = ""
'      End If
'      Exit Function
'    End If
'    If Subj = "du" And Praed = "willst" Then
'      Select Case Int(Rnd * 6) + 1
'        Case 1: MakeReply = "Hast du oft diesen Wunsch?"
'        Case 2: MakeReply = "Glaubst du daß ich dir da helfen kann?"
'        Case 3: MakeReply = "Willst du wirklich" & NewOne & "?"
'        Case 4: MakeReply = "Warum willst du" & NewOne & "?"
'        Case 5: MakeReply = "Gibt es einen Grund dafür, daß " & Subj + NewOne & " " & Praed & "?"
'        Case 6: MakeReply = "Ist es ein starker Wunsch von dir, daß " & Subj + NewOne & " " & Praed & "?"
'      End Select
'      Exit Function
'    End If
'    If Subj = "ich" And Praed = "wiederhole" Then
'      Select Case Int(Rnd * 6) + 1
'        Case 1: MakeReply = "Das kann gar nicht sein - ich habe zu jedem Satz mindestens 2 verschiedene Antworten ;)"
'        Case 2: MakeReply = "*hehe* du paßt ja richtig scharf auf =)"
'        Case 3: MakeReply = "Das mache ich extra, war nur ein Test. | Du hast bestanden! *G*"
'        Case 4: MakeReply = "Ich will damit nur den Anschein erwecken, ein Computerprogramm zu sein *G*"
'        Case 5: MakeReply = "Ach ja? Hast du Beweise dafür? | Ein Log vielleicht? Wenn ja, dann lösch es ;o)"
'        Case 6: MakeReply = "Ich wiederhole mich grundsätzlich nie. | nie. | nie :o)"
'      End Select
'      Exit Function
'    End If
'    If (Len(Subj) < 15) And (Len(NewOne) < 20) And (Len(Praed) < 15) Then
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "Warum denkst du, daß " & Subj + NewOne & " " & Praed & "?"
'        Case 2: MakeReply = "Bist du dir sicher, daß " & Subj + NewOne & " " & Praed & "?": KIs(KNum).Thema = "really?"
'        Case 3: MakeReply = "Findest du wirklich, daß " & Subj + NewOne & " " & Praed & "?": KIs(KNum).Thema = "really?"
'        Case 4: MakeReply = "Warum meinst du, daß " & Subj + NewOne & " " & Praed & "?": KIs(KNum).Thema = "really?"
'        Case 5: MakeReply = "Gibt es einen Grund dafür, daß du denkst, daß " & Subj + NewOne & " " & Praed & "?"
'        Case 6: MakeReply = "Weißt du genau, daß " & Subj + NewOne & " " & Praed & "?": KIs(KNum).Thema = "really?"
'        Case 7: MakeReply = "Woher möchtest du wissen, daß " & Subj + NewOne & " " & Praed & "?"
'        Case 8: MakeReply = "Denkst du wirklich, daß " & Subj + NewOne & " " & Praed & "?": KIs(KNum).Thema = "really?"
'        Case 9: MakeReply = "Hmm... Macht dir das etwas aus?"
'        Case 10: MakeReply = "Und hast du Probleme damit?"
'      End Select
'    Else
'      Select Case Int(Rnd * 10) + 1
'        Case 1: MakeReply = "hehe"
'        Case 2: MakeReply = ""
'        Case 3: MakeReply = "hm"
'        Case 4: MakeReply = ""
'        Case 5: MakeReply = ":)"
'        Case 6: MakeReply = "*G*"
'        Case 7: MakeReply = "hmm.."
'        Case 8: MakeReply = "mh..."
'        Case 9: MakeReply = "?"
'        Case 10: MakeReply = ""
'      End Select
'    End If
'  End Select
'
'  SSatz = ""
'  For u = 1 To Len(Satz)
'    Select Case Mid(Satz, u, 1)
'      Case ".", "!", "?", ","
'        SSatz = SSatz
'      Case Else
'        SSatz = SSatz + Mid(Satz, u, 1)
'    End Select
'  Next u
'  SSatz = SSatz & " "
'  If Praed = "dachtst" Then MakeReply = "Aaa-ha =)"
End Function

Sub RemKI(ByVal Nick As String) ' : AddStack "KI_RemKI(" & Nick & ")"
Dim u As Long, u2 As Long
  For u = 1 To KICount
    If LCase(KIs(u).Nick) = LCase(Nick) Then
      For u2 = u To KICount - 1
        KIs(u2) = KIs(u2 + 1)
      Next u2
      KICount = KICount - 1
      ReDim Preserve KIs(((KICount \ 5) + 1) * 5)
      Exit Sub
    End If
  Next u
End Sub

Function SwitchIfNeeded(Word As String) As String ' : AddStack "KI_SwitchIfNeeded(" & Word & ")"
  Select Case LCase(Word)
    Case "du": SwitchIfNeeded = "ich"
    Case "ich": SwitchIfNeeded = "du"
    Case "mir": SwitchIfNeeded = "dir"
    Case "mich": SwitchIfNeeded = "dich"
    Case "dir": SwitchIfNeeded = "mir"
    Case "dich": SwitchIfNeeded = "mich"
    Case "uns": SwitchIfNeeded = "euch"
    Case "euch": SwitchIfNeeded = "uns"
    Case "deine": SwitchIfNeeded = "meine"
    Case "deinen": SwitchIfNeeded = "meinen"
    Case "meinen": SwitchIfNeeded = "deinen"
    Case "dein": SwitchIfNeeded = "mein"
    Case "meine": SwitchIfNeeded = "deine"
    Case "mein": SwitchIfNeeded = "dein"
    Case "unsere": SwitchIfNeeded = "euere"
    Case "eure", "euere": SwitchIfNeeded = "unsere"
    Case Else: SwitchIfNeeded = Word
  End Select
End Function

Function IsRel(Word As String) As Boolean ' : AddStack "KI_IsRel(" & Word & ")"
  Select Case LCase(Word)
    Case "dir", "mir": IsRel = True
    Case "euch", "uns": IsRel = True
    Case "ihm", "ihnen": IsRel = True
    Case "ihr": IsRel = True
  End Select
End Function

Function IsPossessive(Word As String) As Boolean ' : AddStack "KI_IsPossessive(" & Word & ")"
  Select Case LCase(Word)
    Case "mein", "meine": IsPossessive = True
    Case "dein", "deine": IsPossessive = True
    Case "unser", "unsere": IsPossessive = True
    Case "euer", "eure", "euere": IsPossessive = True
    Case "ihr", "ihre": IsPossessive = True
  End Select
End Function

Function ConvertSubj(Subj As String, Tries As Byte) As String ' : AddStack "KI_ConvertSubj(" & Subj & ", " & Tries & ")"
  If Tries = 2 And LCase(Subj) = "das" Then ConvertSubj = "das": Exit Function
  Select Case LCase(Subj)
    Case "ich": ConvertSubj = "du"
    Case "du": ConvertSubj = "ich"
    Case "er", "sie", "es", "da", "sie": ConvertSubj = LCase(Subj)
    Case "wir": ConvertSubj = "ihr"
    Case "ihr": ConvertSubj = "wir"
    Case "mir": ConvertSubj = "dir"
    Case "dir": ConvertSubj = "mir"
  End Select
End Function

Function TurnAround(Verb As String) As String ' : AddStack "KI_TurnAround(" & Verb & ")"
  Select Case LCase(Verb)
    Case "ist", "war": TurnAround = LCase(Verb)
    Case "hab", "habe": TurnAround = "hast"
    Case "nehme", "nehm": TurnAround = "nimmst"
    Case "laufe", "lauf": TurnAround = "läufst"
    Case "gebe", "geb": TurnAround = "gibst"
    Case "musst", "mußt": TurnAround = "muss"
    Case "hast": TurnAround = "habe"
    Case "habt": TurnAround = "haben"
    Case "haben": TurnAround = "habt"
    Case "sehe", "seh": TurnAround = "siehst"
    Case "siehst": TurnAround = "sehe"
    Case "sende": TurnAround = "sendest"
    Case "senden": TurnAround = "sendet"
    Case "bist": TurnAround = "bin"
    Case "bin": TurnAround = "bist"
    Case "weißt": TurnAround = "weiß"
    Case "weiß": TurnAround = "weißt"
    Case "sind": TurnAround = "seid"
    Case "seid", "seit": TurnAround = "sind"
    Case "mag": TurnAround = "magst"
    Case "halt", "halte": TurnAround = "hältst"
    Case Else
      Select Case LCase(Right(Verb, 3))
        Case "sst": TurnAround = LCase(Left(Verb, Len(Verb) - 1)) & "e"
        Case "sse": TurnAround = LCase(Left(Verb, Len(Verb) - 1)) & "t"
        Case Else
          Select Case LCase(Right(Verb, 2))
            Case "ßt": TurnAround = LCase(Left(Verb, Len(Verb) - 2)) & "sse"
            Case "ße": TurnAround = LCase(Left(Verb, Len(Verb) - 2)) & "sst"
            Case "in": TurnAround = LCase(Left(Verb, Len(Verb) - 1)) & "st"
            Case "st"
              If InStr("hl", Mid(Verb, Len(Verb) - 2, 1)) > 0 Then
                TurnAround = LCase(Left(Verb, Len(Verb) - 2)) & "e"
              Else
                TurnAround = LCase(Left(Verb, Len(Verb) - 2))
              End If
            Case "ze": TurnAround = LCase(Left(Verb, Len(Verb) - 1)) & "t"
            Case "ts": TurnAround = LCase(Left(Verb, Len(Verb) - 1))
            Case "nn": TurnAround = Verb & "st"
            Case "nd": TurnAround = Verb & "est"
            Case "de": TurnAround = Verb & "st"
            Case "ss": TurnAround = Verb & "e"
            Case "en": TurnAround = Verb
            Case "eh": TurnAround = Verb & "st"
            Case "te": TurnAround = Verb & "st"
            Case "au": TurnAround = Verb & "st"
            Case "ll": TurnAround = Verb & "st"
            Case "ar": TurnAround = Verb & "st"
            Case Else
              Select Case LCase(Right(Verb, 1))
                Case "e"
                  If Mid(Verb, Len(Verb) - 2, 1) = "b" Then
                    TurnAround = LCase(Left(Verb, Len(Verb) - 2)) & "st"
                  ElseIf InStr("aeiou", Mid(Verb, Len(Verb) - 1, 1)) = 0 Then
                    TurnAround = LCase(Left(Verb, Len(Verb) - 1)) & "st"
                  End If
                Case "b": TurnAround = Verb & "st"
                Case "ß": TurnAround = LCase(Left(Verb, Len(Verb) - 1)) & "ßt"
                Case Else
                  TurnAround = Verb
              End Select
          End Select
      End Select
  End Select
End Function

Function RemStr(ByVal OldLine As String, Remove As String) As String ' : AddStack "KI_RemStr(" & OldLine & ", " & Remove & ")"
Dim SPos As Long, NewLine As String
  NewLine = OldLine
  Do
    SPos = InStr(NewLine, Remove)
    If SPos > 0 Then NewLine = Left(NewLine, SPos - 1) + Mid(NewLine, SPos + Len(Remove)) Else Exit Do
  Loop
  RemStr = NewLine
End Function

Function IsInLine(ByVal Line As String, ByVal What As String) As Boolean ' : AddStack "KI_IsInLine(" & Line & ", " & What & ")"
Dim u As Long, i As Long, ThePart As String, TheWord As String, ResVal As Boolean
Dim NewLine As String, Char As String, Begins(10) As Long, BeginCount As Long
  Line = Replace(Line, "ä", "ae")
  Line = Replace(Line, "ö", "oe")
  Line = Replace(Line, "ü", "ue")
  Line = Replace(Line, "ß", "ss")
  What = Replace(What, "ä", "ae")
  What = Replace(What, "ö", "oe")
  What = Replace(What, "ü", "ue")
  What = Replace(What, "ß", "ss")
  Line = " " & LCase(MakeININick(Line)) & " "
  What = "[" & LCase(What) & "]"
  For u = 1 To Len(What)
    Char = Mid(What, u, 1)
    If InStr("[] |", Char) > 0 Then
      If TheWord <> "" Then
        ThePart = "*[!a-z0-9äöüß]" & MakeININick(TheWord) & "[!a-z0-9äöüß]*"
        What = Left(What, u - Len(TheWord) - 1) + Spaces2(Len(TheWord), IIf(Line Like ThePart, "+", "-")) + Mid(What, u)
        TheWord = ""
      End If
    End If
    Select Case Char
      Case "["
        If BeginCount < 10 Then
          BeginCount = BeginCount + 1
          Begins(BeginCount) = u
        End If
      Case "]"
        If BeginCount > 0 Then
          ThePart = Mid(What, Begins(BeginCount) + 1, u - Begins(BeginCount) - 1)
          What = Left(What, Begins(BeginCount) - 1) + Spaces2(Len(ThePart) + 2, IIf(IsTrue(ThePart), "+", "-")) + Mid(What, u + 1)
          BeginCount = BeginCount - 1
        End If
      Case " ", "|"
      Case Else
        TheWord = TheWord + Char
    End Select
  Next u
  IsInLine = (Left(What, 1) = "+")
End Function

Function IsTrue(ByVal What As String) As Boolean ' : AddStack "KI_IsTrue(" & What & ")"
Dim u As Long, i As Long, ThePart As String, ResVal As Boolean
  For i = 1 To ParamXCount(What, "|")
    ThePart = Trim(ParamX(What, "|", i))
    ResVal = True
    For u = 1 To ParamCount(ThePart)
      If Param(ThePart, u) = "-" Then ResVal = False: Exit For
    Next u
    If ResVal = True Then IsTrue = True: Exit Function
  Next i
  IsTrue = False
End Function

