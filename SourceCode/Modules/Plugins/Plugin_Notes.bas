Attribute VB_Name = "Plugin_Notes"
Option Explicit

Private NotesImport As Byte

Private Type tNoteType
  Index As String * 10
  Date As Date
  Flag As String
  From As String
  Text As String
  Nicks As String
End Type

Public Note() As tNoteType

Public NoteCount As Long


Sub Notes_Load()
  ReDim Preserve Note(5)
End Sub


Sub Notes_Unload()
'
End Sub


Sub CheckNotes(RegNick As String) ' : AddStack "mdlWinsock_CheckNotes(" & RegNick & ")"
Dim ChNum As Long, u As Long, u2 As Long, RegUser As String, NoteCount As Long
  For ChNum = 1 To ChanCount
    For u2 = 1 To Channels(ChNum).UserCount
      RegUser = Channels(ChNum).User(u2).RegNick
      If RegUser <> "" Then
        If Not IsIgnored(Channels(ChNum).User(u2).Hostmask) Then
          NoteCount = NotesCount(RegUser)
          If ((RegNick <> "") And (LCase(RegNick) = LCase(RegUser))) Or (RegNick = "") Then
            If NoteCount = 1 Then
              If BotUsers(GetUserNum(RegUser)).Password <> "" Then
                AddIgnore Mask(Channels(ChNum).User(u2).Hostmask, 2), 10, 1
                SendLine "notice " & Channels(ChNum).User(u2).Nick & " :Hi! I've got 1 note waiting for you.", 3
                SendLine "notice " & Channels(ChNum).User(u2).Nick & " :To get it, type: /msg " & MyNick & " notes <pass>", 3
              Else
                AddIgnore Mask(Channels(ChNum).User(u2).Hostmask, 2), 10, 1
                SpreadFlagMessage 0, "+m", "14[" & Time & "] *** I told " & IIf(Channels(ChNum).User(u2).Nick <> RegUser, Channels(ChNum).User(u2).Nick & " (" & RegUser & ")", Channels(ChNum).User(u2).Nick) & " to set a password: Waiting notes!"
                SendLine "notice " & Channels(ChNum).User(u2).Nick & " :Hi! I've got notes for you, but you don't have a password set.", 3
                SendLine "notice " & Channels(ChNum).User(u2).Nick & " :To set one, type: /msg " & MyNick & " pass <your password>", 3
              End If
            End If
          End If
        End If
      End If
    Next u2
  Next ChNum
End Sub

Function SendNote(FromUser As String, ToUser As String, NoteFlag As String, Text As String, Optional NDate As String = "") As Boolean
  Dim NoteDate As Date, NoteID As Long, Rest As String, RealNick As String
  If NDate <> "" Then NoteDate = GetDate(NDate) Else NoteDate = Now
SendNote:
  If FromUser = "" Or ToUser = "" Or Text = "" Then Exit Function
  If GetUserData(GetUserNum(ToUser), "fwdaddr", "") = "" Then
    NoteID = FindNote(FromUser, NoteFlag, Text, NoteDate)
    If NoteID <> 0 Then
      If InStr(1, " " & Note(NoteID).Nicks & " ", " " & ToUser & " ", vbBinaryCompare) <> 0 Then
        SendNote = False
      Else
        Note(NoteID).Nicks = Note(NoteID).Nicks & " " & ToUser
        SendNote = True
      End If
    Else
      NoteCount = NoteCount + 1
      NoteID = NoteCount
      If NoteCount > UBound(Note()) Then ReDim Preserve Note(NoteCount + 5)
      Note(NoteID).Index = CreateNoteIndex
      Note(NoteID).Date = NoteDate
      Note(NoteID).Flag = NoteFlag
      Note(NoteID).Nicks = ToUser
      Note(NoteID).Text = Text
      Note(NoteID).From = FromUser
      SendNote = True
    End If
  Else
    Rest = GetPartBot(GetUserData(GetUserNum(ToUser), "fwdaddr", ""))
    RealNick = GetPartNick(GetUserData(GetUserNum(ToUser), "fwdaddr", ""))
    If GetBotPos(Rest) = 0 Then SetUserData GetUserNum(ToUser), "fwdaddr", "": GoTo SendNote
    SendToBot Rest, "p *:" & FromUser & "@" & BotNetNick & " " & GetUserData(GetUserNum(ToUser), "fwdaddr", "") & " " & Text
    SendNote = True
  End If
  If SendNote = True And NotesImport = 0 Then WriteNotes
End Function

Sub NotesChangeNick(OldNick As String, NewNick As String)
  Dim Index As Long
  If NoteCount > 0 Then
    For Index = 1 To NoteCount
      Note(Index).Nicks = Replace(Note(Index).Nicks, OldNick, NewNick)
      Note(Index).From = Replace(Note(Index).From, OldNick, NewNick)
    Next Index
  End If
  WriteNotes
End Sub

Sub NotesErase(Nick As String)
  Dim Index As Long
  Dim Index2 As Long
  Dim Dummy As Long
  If NoteCount > 0 Then
    For Index = 1 To NoteCount
      If Index > NoteCount Then Exit For
      Note(Index).Nicks = Replace(Note(Index).Nicks, Nick, "")
      Note(Index).Nicks = Replace(Note(Index).Nicks, "  ", " ")
      If Left(Note(Index).Nicks, 1) = " " Then Note(Index).Nicks = Mid(Note(Index).Nicks, 2)
      If Right(Note(Index).Nicks, 1) = " " Then Note(Index).Nicks = Left(Note(Index).Nicks, Len(Note(Index).Nicks))
    Next Index
  End If
  For Index = 1 To NoteCount
    If Note(Index).Nicks = "" Then
      For Index2 = Index To NoteCount - 1
        If Note(Index2).Nicks <> "" Then
          Note(Index2) = Note(Index2 + 1)
          Dummy = Dummy + 1
          Exit For
        End If
      Next Index2
    Else
      Dummy = Dummy + 1
    End If
  Next
  NoteCount = Dummy
  WriteNotes
End Sub

Sub NotesKill()
  NoteCount = 0
  ReDim Preserve Note(5)
  WriteNotes
End Sub

Sub ImportOldNotes()
  Dim FileNum As Integer, Line As String, Index As Integer
  FileNum = FreeFile
  Open HomeDir & "Notes.ini" For Input As #FileNum
  If Err.Number = 0 Then
    Do While Not EOF(FileNum)
      Line Input #FileNum, Line
      If Left(Line, 1) = "[" And Right(Line, 1) = "]" Then
        If Val(GetPPString(Mid(Line, 2, Len(Line) - 2), "NoteCount", "0", HomeDir & "Notes.ini")) > 0 Then
          For Index = 1 To Val(GetPPString(Mid(Line, 2, Len(Line) - 2), "NoteCount", "0", HomeDir & "Notes.ini"))
            SendNote GetPPString(Mid(Line, 2, Len(Line) - 2), "Note" & CStr(Index) & "From", "<unknown>", HomeDir & "Notes.ini"), Mid(Line, 2, Len(Line) - 2), GetPPString(Mid(Line, 2, Len(Line) - 2), "Note" & CStr(Index) & "Flag", "", HomeDir & "Notes.ini"), GetPPString(Mid(Line, 2, Len(Line) - 2), "Note" & CStr(Index) & "Text", "<empty note>", HomeDir & "Notes.ini"), Replace(GetPPString(Mid(Line, 2, Len(Line) - 2), "Note" & CStr(Index) & "Date", DateAdd("h", 1, DateSerial(1998, 1, 1)), HomeDir & "Notes.ini"), ",", ".")
          Next Index
        End If
      End If
    Loop
  End If
  Close #FileNum
  Err.Clear
  Kill HomeDir & "Notes.ini"
End Sub

Function NotesCount(Nick As String) As Long
  Dim Index As Long
  Dim Dummy As Long
  Dummy = 0
  If NoteCount > 0 Then
    For Index = 1 To NoteCount
      If InStr(1, " " & Note(Index).Nicks & " ", " " & Nick & " ", vbBinaryCompare) Then Dummy = Dummy + 1
    Next Index
  End If
  NotesCount = Dummy
End Function

Function NotesFlag(Nick As String, Index As Long) As String
  Dim NIndex As String
  NIndex = NoteNickIndex(Nick, Index)
  NotesFlag = NoteIndexFlag(NIndex)
End Function

Function NotesText(Nick As String, Index As Long) As String
  Dim NIndex As String
  NIndex = NoteNickIndex(Nick, Index)
  NotesText = NoteIndexText(NIndex)
End Function

Function NotesFrom(Nick As String, Index As Long) As String
  Dim NIndex As String
  NIndex = NoteNickIndex(Nick, Index)
  NotesFrom = NoteIndexFrom(NIndex)
End Function

Function NotesDate(Nick As String, Index As Long) As String
  Dim NIndex As String
  NIndex = NoteNickIndex(Nick, Index)
  NotesDate = NoteIndexDate(NIndex)
End Function

Function CreateNoteIndex() As String
  Dim NotesID As String, u As Long
Start:
  NotesID = ""
  For u = 1 To 10
    NotesID = NotesID & Int(Rnd * 10)
  Next u
  If NoteIndexText(NotesID) <> "" Then GoTo Start
  CreateNoteIndex = NotesID
End Function

Function NoteNickIndex(Nick As String, Index As Long)
  Dim u As Long, u2 As Long
  u2 = 0
  If NoteCount = 0 Then Exit Function
  For u = 1 To NoteCount
    If InStr(" " & Note(u).Nicks & " ", " " & Nick & " ") > 0 Then
      u2 = u2 + 1
      If u2 = Index Then
        NoteNickIndex = Note(u).Index
        Exit Function
      End If
    End If
  Next u
End Function

Function NoteIndexText(Index As String) As String
  Dim u As Long
  NoteIndexText = ""
  If NoteCount = 0 Then Exit Function
  For u = 1 To NoteCount
    If Note(u).Index = Index Then
      NoteIndexText = Note(u).Text
      Exit Function
    End If
  Next u
End Function

Function NoteIndexDate(Index As String) As String
  Dim u As Long
  NoteIndexDate = ""
  If NoteCount = 0 Then Exit Function
  For u = 1 To NoteCount
    If Note(u).Index = Index Then
      NoteIndexDate = Note(u).Date
      Exit Function
    End If
  Next u
End Function

Function NoteIndexFlag(Index As String) As String
  Dim u As Long
  NoteIndexFlag = ""
  If NoteCount = 0 Then Exit Function
  For u = 1 To NoteCount
    If Note(u).Index = Index Then
      NoteIndexFlag = Note(u).Flag
      Exit Function
    End If
  Next u
End Function

Function NoteIndexFrom(Index As String) As String
  Dim u As Long
  NoteIndexFrom = ""
  If NoteCount = 0 Then Exit Function
  For u = 1 To NoteCount
    If Note(u).Index = Index Then
      NoteIndexFrom = Note(u).From
      Exit Function
    End If
  Next u
End Function

Function FindNote(FromUser As String, NoteFlag As String, Text As String, NoteDate As Date) As Long
  Dim u As Long
  If NoteCount = 0 Then Exit Function
  For u = 1 To NoteCount
    If Note(u).From = FromUser And Note(u).Flag = NoteFlag And Note(u).Text = Text And Format(Note(u).Date, "dd.mm.yy, hh:nn:ss") = Format(NoteDate, "dd.mm.yy, hh:nn:ss") Then
      FindNote = u
      Exit Function
    End If
  Next u
End Function

Public Sub ReadNotes()
  NotesImport = 0
  If Dir(HomeDir & "notes.ini") <> "" Then
    NotesImport = 1
    ImportOldNotes
  End If
  
  Dim FileNum As Integer, Line As String
  Dim NoteID As Long, CurIndex As String, CurDate As Date, CurFlag As String, CurNicks As String, CurText As String, CurSender As String
  NoteID = 0
  CurIndex = ""
  CurDate = DateSerial(1998, 1, 1)
  CurFlag = ""
  CurNicks = ""
  CurText = ""
  CurSender = ""
  
  If Dir(HomeDir & "notes.txt") = "" Then Exit Sub
  
  On Local Error Resume Next
  
  FileNum = FreeFile
  Open HomeDir & "Notes.txt" For Input As #FileNum
  If Err.Number = 0 Then
    On Error GoTo Erroar
    Do While Not EOF(FileNum)
      Line Input #FileNum, Line
      Select Case Param(Line, 1)
        Case "---"
          If NoteID <> 0 Then
            If CurIndex = "" Or CurDate = DateSerial(1998, 1, 1) Or CurNicks = "" Or CurText = "" Or CurSender = "" Then
              NoteCount = NoteCount - 1
              NoteID = ""
              If NoteCount > 0 And NoteCount + 5 <= UBound(Note) Then ReDim Preserve Note(NoteCount)
            Else
              Note(NoteID).Index = CurIndex
              Note(NoteID).Date = CurDate
              Note(NoteID).Flag = CurFlag
              Note(NoteID).Nicks = CurNicks
              Note(NoteID).Text = CurText
              Note(NoteID).From = CurSender
            End If
          End If
          CurIndex = Param(Line, 2)
          CurSender = Param(Line, 3)
          CurFlag = ""
          CurNicks = ""
          CurText = ""
          CurDate = DateSerial(1998, 1, 1)
              
          NoteCount = NoteCount + 1
          NoteID = NoteCount
          If NoteCount > UBound(Note()) Then ReDim Preserve Note(NoteCount + 5)
        Case "f"
          CurFlag = Param(Line, 2)
        Case "d"
          CurDate = GetDate(Param(Line, 2))
        Case "n"
          CurNicks = GetRest(Line, 2)
        Case "t"
          CurText = GetRest(Line, 2)
      End Select
    Loop
  End If
  Close #FileNum
  Err.Clear
  
  If NoteID <> 0 Then
    If CurIndex = "" Or CurDate = DateSerial(1998, 1, 1) Or CurNicks = "" Or CurText = "" Or CurSender = "" Then
      NoteCount = NoteCount - 1
      NoteID = 0
      If NoteCount > 0 And NoteCount + 5 <= UBound(Note()) Then ReDim Preserve Note(NoteCount)
    Else
      Note(NoteID).Index = CurIndex
      Note(NoteID).Date = CDate(CDbl(CurDate))
      Note(NoteID).Flag = CurFlag
      Note(NoteID).Nicks = CurNicks
      Note(NoteID).Text = CurText
      Note(NoteID).From = CurSender
    End If
  End If
  
  On Error GoTo Erroar

  If NotesImport = 1 Then WriteNotes
  Exit Sub
Erroar:
  MsgBox "Ein Fehler ist aufgetreten!" & " - VB meint dazu: " & Err.Number & " - '" & Err.Description & "'. Zeile: " & Line
  End
End Sub

Sub WriteNotes()
  Dim FileName As String, FileNumber As Integer, u As Long, ErrLine As Long
  
  FileName = HomeDir & "Notes.txt"
  If NoteCount > 0 Then
    WaitForAccess FileName
    AddAccessedFile FileName
    On Error GoTo WULErr
    FileNumber = FreeFile
    Open FileName For Output As #FileNumber
    Print #FileNumber, "' This is an AnGeL Bot notefile. Please don't change anything"
    Print #FileNumber, "' in here unless you know what you're doing."
    Print #FileNumber, ""
    For u = 1 To NoteCount
      Print #FileNumber, "--- " & Note(u).Index & " " & Note(u).From
      Print #FileNumber, "d " & Replace(CStr(CDbl(Note(u).Date)), ",", ".")
      Print #FileNumber, "t " & Note(u).Text
      Print #FileNumber, "n " & Note(u).Nicks
      If Note(u).Flag <> "" Then Print #FileNumber, "f " & Note(u).Flag
    Next u
    Close #FileNumber
    RemAccessedFile FileName
  Else
    If Dir("notes.txt") = "" Then Exit Sub
    Kill FileName
  End If
Exit Sub
WULErr:
  Close #FileNumber
  RemAccessedFile FileName
  Err.Clear
End Sub

