Attribute VB_Name = "Plugin_WhatIs"
',-======================- ==-- -  -
'|   AnGeL - Plugins - Whatis
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public LastWhatisOutput As String


Sub Whatis_Load()
'
End Sub

Sub Whatis_Unload()
'
End Sub


Function WriteWhatis(Entry As String, Value As String) As String
  Dim FileName As String, FileLines() As String, LineCount As Long, u As Long, ErrLine As Long
  Dim CurLine As String, FileNumber As Integer, Rest As String
  ReDim FileLines(100)
  FileName = HomeDir & "Whatis.txt"
  On Local Error Resume Next
  If Dir(FileName) <> "" Then
    On Local Error Resume Next
ErrLine = 1
    WaitForAccess FileName
    AddAccessedFile FileName
    FileNumber = FreeFile: Open FileName For Input As #FileNumber
    If Err.Number <> 0 Then Close #FileNumber: RemAccessedFile FileName: Exit Function 'Error; don't change anything
    On Error GoTo WriteWhatErr
ErrLine = 2
    Do While Not EOF(FileNumber)
ErrLine = 3
      Line Input #FileNumber, CurLine
ErrLine = 4
      If LCase(AddSpaces(Param(CurLine, 1))) <> LCase(Entry) And Trim(CurLine) <> "" Then
ErrLine = 5
        LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
ErrLine = 13
        FileLines(LineCount) = CurLine
ErrLine = 14
      Else
        If LCase(AddSpaces(Param(CurLine, 1))) = LCase(Entry) Then WriteWhatis = AddSpaces(Param(CurLine, 1))
      End If
    Loop
ErrLine = 15
    Close #FileNumber
  Else
    WaitForAccess FileName
    AddAccessedFile FileName
  End If
  On Error GoTo WriteWhatErr
ErrLine = 16
  If Value <> "" Then
    LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
ErrLine = 17
    FileLines(LineCount) = RemSpaces(Entry) & " " & RemSpaces(Value)
  End If
ErrLine = 18
  FileNumber = FreeFile: Open FileName For Output As #FileNumber
ErrLine = 19
    For u = 1 To LineCount
ErrLine = 20
      Print #FileNumber, FileLines(u)
ErrLine = 21
    Next u
ErrLine = 22
  Close #FileNumber
  RemAccessedFile FileName
Exit Function
WriteWhatErr:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  Close #FileNumber
  RemAccessedFile FileName
  PutLog "||| ] WriteWhatis ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
  PutLog "||| ] Der Fehler " & ErrNumber & " (" & ErrDescription & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & ErrLine
  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
End Function

Function ReadWhatis(Entry As String, ExactMatchRequired As Boolean, Nick As String) As String
  Dim FileName As String, FileLines() As String, LineCount As Long, u As Long, ErrLine As Long
  Dim CurLine As String, FileNumber As Integer, Rest As String, u2 As Long, Match As Boolean
  Dim MatchCounter As Long, MatchList As String, Temp As String, Temp2 As String
  If Trim(Entry) = "" Then Exit Function
  ReDim FileLines(100)
  FileName = HomeDir & "Whatis.txt"
  On Local Error Resume Next
  If Dir(FileName) <> "" Then
    On Local Error Resume Next
ErrLine = 1
    FileNumber = FreeFile
    Open FileName For Input As #FileNumber
    If Err.Number = 0 Then
      On Error GoTo ReadWhatErr
  ErrLine = 2
      Do While Not EOF(FileNumber)
  ErrLine = 3
        Line Input #FileNumber, CurLine
  ErrLine = 4
        'Exact match
        If LCase(AddSpaces(Param(CurLine, 1))) = LCase(Entry) Then
          ReadWhatis = AddSpaces(Param(CurLine, 1)) & ": " & AddSpaces(Param(CurLine, 2))
          Close #FileNumber
          Exit Function
        End If
        LineCount = LineCount + 1: If LineCount > UBound(FileLines()) Then ReDim Preserve FileLines(UBound(FileLines()) + 5)
  ErrLine = 13
        FileLines(LineCount) = CurLine
  ErrLine = 14
      Loop
  ErrLine = 15
      Close #FileNumber
    Else
      Close #FileNumber
      Err.Clear
    End If
  End If
  On Error GoTo ReadWhatErr
ErrLine = 16
  If ExactMatchRequired Then ReadWhatis = "": Exit Function
  'Non-exact match
  For u = 1 To LineCount
    Match = True
    Rest = LCase(AddSpaces(Param(FileLines(u), 1)))
    For u2 = 1 To ParamCount(Entry)
      If InStr(Rest, LCase(Param(Entry, u2))) = 0 Then Match = False: Exit For
    Next u2
    If Match Then
      MatchCounter = MatchCounter + 1: If MatchCounter > 10 Then ReadWhatis = "Too many matches. Please search for something more specific.": Exit Function
      If MatchList = "" Then Temp = AddSpaces(Param(FileLines(u), 1)): Temp2 = AddSpaces(Param(FileLines(u), 2)): MatchList = Nick & ", choose an entry: """ & Rest & """" Else MatchList = MatchList & ", """ & Rest & """"
    End If
  Next u
  If MatchCounter > 0 Then
    If MatchCounter = 1 Then ReadWhatis = Temp & ": " & Temp2 Else ReadWhatis = MatchList
  End If
Exit Function
ReadWhatErr:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  Close #FileNumber
  PutLog "||| ] ReadWhatis ERROR!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
  PutLog "||| ] Der Fehler " & CStr(ErrNumber) & " (" & ErrDescription & ") ist beim Bearbeiten folgender Zeile aufgetreten: " & ErrLine
  PutLog "||| ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
End Function

