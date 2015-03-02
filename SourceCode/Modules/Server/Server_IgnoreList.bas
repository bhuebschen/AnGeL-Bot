Attribute VB_Name = "Server_IgnoreList"
',-======================- ==-- -  -
'|   AnGeL - Server - IgnoreList
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public Type Ignore
  Hostmask As String
  CreatedAt As Date
  CreatedBy As String
  Comment As String
End Type

Private Type IgnoreData
  Hostmask As String
  IgnoreLevel As Long
  SpecialIgnore As Boolean
End Type


Public IgnoreCount As Long
Public Ignores() As Ignore

Public IgnoredUserCount As Byte
Public IgnoredUsers() As IgnoreData


Sub IgnoreList_Load()
  ReDim Preserve IgnoredUsers(5)
  ReDim Preserve Ignores(5)
End Sub


Sub IgnoreList_Unload()
'
End Sub


Public Function GetIgnoreLevel(Hostmask As String) As Long
  Dim u As Byte
  GetIgnoreLevel = 0
  If IgnoredUserCount > 0 Then
    For u = 1 To IgnoredUserCount
      If MatchWM(IgnoredUsers(u).Hostmask, Hostmask) Then
        GetIgnoreLevel = IgnoredUsers(u).IgnoreLevel
        Exit Function
      End If
    Next u
  End If
End Function


Public Sub SetIgnoreLevel(Hostmask As String, Level As Long)
  Dim u As Byte
  If IgnoredUserCount > 0 Then
    For u = 1 To IgnoredUserCount
      If MatchWM(IgnoredUsers(u).Hostmask, Hostmask) Then
        IgnoredUsers(u).IgnoreLevel = Level
        Exit Sub
      End If
    Next u
  End If
End Sub


Public Function IsIgnored(Hostmask As String) As Boolean
  Dim u As Byte
  IsIgnored = False
  If IgnoredUserCount > 0 Then
    For u = 1 To IgnoredUserCount
      If MatchWM(IgnoredUsers(u).Hostmask, Hostmask) Then
        IsIgnored = True
        Exit Function
      End If
    Next u
  End If
End Function


Public Function IsSpecialIgnored(Hostmask As String) As Boolean
  Dim u As Byte
  IsSpecialIgnored = False
  If IgnoredUserCount > 0 Then
    For u = 1 To IgnoredUserCount
      If MatchWM(IgnoredUsers(u).Hostmask, Hostmask) And IgnoredUsers(u).SpecialIgnore Then
        IsSpecialIgnored = True
        Exit Function
      End If
    Next u
  End If
End Function


Public Sub AddIgnore(Hostmask As String, HowLong As Currency, Level As Long)
  IgnoredUserCount = IgnoredUserCount + 1
  If IgnoredUserCount > UBound(IgnoredUsers()) Then ReDim Preserve IgnoredUsers(UBound(IgnoredUsers()) + 5)
  IgnoredUsers(IgnoredUserCount).Hostmask = Hostmask
  IgnoredUsers(IgnoredUserCount).IgnoreLevel = Level
  TimedEvent "UnIgnore " & Hostmask, HowLong
End Sub


Public Sub AddSpecialIgnore(Hostmask As String, HowLong As Currency)
  IgnoredUserCount = IgnoredUserCount + 1
  If IgnoredUserCount > UBound(IgnoredUsers()) Then ReDim Preserve IgnoredUsers(UBound(IgnoredUsers()) + 5)
  IgnoredUsers(IgnoredUserCount).Hostmask = Hostmask
  IgnoredUsers(IgnoredUserCount).SpecialIgnore = True
  TimedEvent "UnIgnore " & Hostmask, HowLong
End Sub


Public Sub RemIgnore(Hostmask As String)
  Dim u As Byte, posi As Byte
  For u = 1 To IgnoredUserCount
    If LCase(Hostmask) = LCase(IgnoredUsers(u).Hostmask) Then
      posi = u
      Exit For
    End If
  Next u
  If posi = 0 Then Exit Sub
  For u = posi To IgnoredUserCount - 1
    IgnoredUsers(u) = IgnoredUsers(u + 1)
  Next u
  IgnoredUserCount = IgnoredUserCount - 1
  u = ((IgnoredUserCount \ 5) + 1) * 5
  If u < UBound(IgnoredUsers()) Then ReDim Preserve IgnoredUsers(u)
End Sub


Sub ReadIgnores()
  Dim FileNum As Integer, Line As String, CreatedBy As String
  If Dir(HomeDir & "Ignores.ini") = "" Then Exit Sub
  On Local Error Resume Next
  IgnoreCount = 0
  FileNum = FreeFile
  Open HomeDir & "Ignores.ini" For Input As #FileNum
  If Err.Number = 0 Then
    While Not EOF(FileNum)
      Line Input #FileNum, Line
      Line = Trim(Line)
      If Left(Line, 1) = "[" And Right(Line, 1) = "]" Then
        Line = MakeNormalNick(Mid(Line, 2, Len(Line) - 2))
        CreatedBy = GetPPString(Line, "CreatedBy", "", HomeDir & "Ignores.ini")
        If CreatedBy <> "" Then
          IgnoreCount = IgnoreCount + 1
          If IgnoreCount > UBound(Ignores()) Then ReDim Preserve Ignores(UBound(Ignores()) + 50)
          Ignores(IgnoreCount).Hostmask = Line
          Ignores(IgnoreCount).CreatedBy = CreatedBy
          Ignores(IgnoreCount).CreatedAt = GetPPString(Line, "CreatedAt", DateAdd("h", 1, DateSerial(1998, 1, 1)), HomeDir & "Ignores.ini")
          Ignores(IgnoreCount).Comment = GetPPString(Line, "Comment", "", HomeDir & "Ignores.ini")
        End If
      End If
    Wend
    Close #FileNum
  Else
    Err.Clear
  End If
End Sub


Function IgnoreCheck(Hostmask As String, IgnoreType As Long, Seconds As Currency) As Boolean
  If Not IsIgnored(Hostmask) Then
    AddIgnore Mask(Hostmask, IgnoreType), Seconds, 1
    IgnoreCheck = True
  Else
    IgnoreCheck = False
  End If
End Function


Function IsIgnoredHost(Hostmask As String) As Boolean
  Dim u As Long
  IsIgnoredHost = False
  If IgnoreCount > 0 Then
    For u = 1 To IgnoreCount
      If MatchWM2(Ignores(u).Hostmask, Hostmask) Or MatchWM2(Hostmask, Ignores(u).Hostmask) Then
        IsIgnoredHost = True
        Exit Function
      End If
    Next u
  End If
End Function


Function IsIgnoredTHost(Hostmask As String) As Boolean
  Dim u As Long
  IsIgnoredTHost = False
  If IgnoreCount > 0 Then
    For u = 1 To IgnoreCount
      If Mask(Ignores(u).Hostmask, 14) = "*!*@" Then
        If MatchWM2(Ignores(u).Hostmask, Hostmask) Or MatchWM2(Hostmask, Ignores(u).Hostmask) Then
          IsIgnoredTHost = True
          Exit Function
        End If
      End If
    Next u
  End If
End Function

