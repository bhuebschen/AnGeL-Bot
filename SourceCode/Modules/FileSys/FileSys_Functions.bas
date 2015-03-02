Attribute VB_Name = "FileSys_Functions"
',-======================- ==-- -  -
'|   AnGeL - FileSys - Functions
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public JustAccessing As String




Sub DeletePPString(ByVal Section As String, ByVal Entry As String, ByVal FileName As String)
  Dim Result As Long
  
  Section = MakeININick(Section)
  
  If Entry <> "" Then
    Result = kernel32_WritePrivateProfileStringA(Section, Entry, ByVal 0&, FileName)
  Else
    Result = kernel32_WritePrivateProfileStringA(Section, ByVal 0&, ByVal 0&, FileName)
  End If
  If Result <> 0 Then Exit Sub
  
  Dim FileAttrib As Byte
  FileAttrib = GetAttr(FileName)
  If (FileAttrib And vbSystem) > 0 Or (FileAttrib And vbReadOnly) > 0 Then
    Call SetAttr(FileName, vbArchive)
    While (FileAttrib And vbSystem) > 0 Or (FileAttrib And vbReadOnly) > 0
      Result = Result + 1
      If Result = 500 Then Exit Sub
      DoEvents
      FileAttrib = GetAttr(FileName)
    Wend
    If Entry <> "" Then
      Result = kernel32_WritePrivateProfileStringA(Section, Entry, ByVal 0&, FileName)
    Else
      Result = kernel32_WritePrivateProfileStringA(Section, ByVal 0&, ByVal 0&, FileName)
    End If
  End If
End Sub


Sub WritePPString(ByVal Section As String, ByVal Entry As String, ByVal Value As String, ByVal FileName As String)
  Dim Result As Long
  
  Section = MakeININick(Section)
  Value = MakeINIValue(Value)
  Result = kernel32_WritePrivateProfileStringA(Section, Entry, Value, FileName)
  If Result <> 0 Then Exit Sub
  
  Dim FileAttrib As Byte
  FileAttrib = GetAttr(FileName)
  If (FileAttrib And vbSystem) > 0 Or (FileAttrib And vbReadOnly) > 0 Then
    Call SetAttr(FileName, vbArchive)
    While (FileAttrib And vbSystem) > 0 Or (FileAttrib And vbReadOnly) > 0
      Result = Result + 1
      If Result = 500 Then Exit Sub
      DoEvents
      FileAttrib = GetAttr(FileName)
    Wend
    Result = kernel32_WritePrivateProfileStringA(Section, Entry, Value, FileName)
  End If
End Sub


Sub WritePPSection(ByVal Section As String, ByVal Value As String, ByVal FileName As String)
  Dim Result As Boolean
  
  Section = MakeININick(Section)
  Result = kernel32_WritePrivateProfileSectionA(Section, Value, FileName)
  If Result = True Then Exit Sub
  
  Dim FileAttrib As Byte
  FileAttrib = GetAttr(FileName)
  If (FileAttrib And vbSystem) > 0 Or (FileAttrib And vbReadOnly) > 0 Then
    Call SetAttr(FileName, vbArchive)
    While (FileAttrib And vbSystem) > 0 Or (FileAttrib And vbReadOnly) > 0
      Result = Result + 1
      If Result = 500 Then Exit Sub
      DoEvents
      FileAttrib = GetAttr(FileName)
    Wend
    Result = kernel32_WritePrivateProfileSectionA(Section, Value, FileName)
  End If
End Sub


Function GetPPString(ByVal Section As String, ByVal Entry As String, ByVal Default As String, ByVal FileName As String) As String
  Dim Result As Integer, Value As String
  
  If Section = "" Then
    GetPPString = ""
  Else
    Section = MakeININick(Section)
    GetPPString = Space(5000&)
    Result = kernel32_GetPrivateProfileStringA(Section, Entry, "NOT_DEFINED", GetPPString, 5000&, FileName)
    Value = Left(GetPPString, Result)
    Value = MakeNormalValue(Value)
    If Value = "NOT_DEFINED" Then Value = Default
    GetPPString = Value
  End If
  If LCase(Section) = "others" Then Debug.Print Section, Entry, FileName
End Function


Function GetPPSection(ByVal Section As String, ByVal FileName As String) As String
  Dim Result As Integer, Value As String
  If Section = "" Then
    GetPPSection = ""
  Else
    Section = MakeININick(Section)
    GetPPSection = Space(8000&)
    Result = kernel32_GetPrivateProfileSectionA(Section, GetPPSection, 8000&, FileName)
    GetPPSection = Left(GetPPSection, Result)
  End If
End Function


Function GetNewDir(OldDir As String, ByVal Target As String, BaseDir As String) As String
  Dim Index As Long, Char As String, Part As String, NewDir As String
  Target = Replace(Target, "/", "\")
  NewDir = OldDir
  For Index = 1 To Len(Target)
    Char = Mid(Target, Index, 1)
    Select Case Char
      Case "\"
        If Index = 1 Then
          NewDir = BaseDir
        Else
          If Part <> "" Then
            Select Case Part
              Case ".": NewDir = NewDir
              Case "..": NewDir = OneDirBack(NewDir)
              Case Else
                If Trim(Replace(Part, ".", "")) <> "" Then
                  If Right(NewDir, 1) = "\" Then NewDir = NewDir & Part Else NewDir = NewDir & "\" & Part
                End If
            End Select
          End If
        End If
        Part = ""
      Case Else
        Part = Part & Char
    End Select
  Next Index
  If Part <> "" Then
    Select Case Part
      Case ".": NewDir = NewDir
      Case "..": NewDir = OneDirBack(NewDir)
      Case Else
        If Trim(Replace(Part, ".", "")) <> "" Then
          If Right(NewDir, 1) = "\" Then NewDir = NewDir & Part Else NewDir = NewDir & "\" & Part
        End If
    End Select
  End If
  GetNewDir = NewDir
End Function


Function OneDirBack(OldDir As String) As String
  Dim NewDir As String
  If Right(OldDir, 1) = "\" Then OldDir = Left(OldDir, Len(OldDir) - 1)
  If InStr(1, OldDir, "\", vbBinaryCompare) <> 0 Then NewDir = Mid(OldDir, 1, InStrRev(OldDir, "\", -1, vbBinaryCompare) - 1)
  If NewDir = "" Then NewDir = "\" Else If Right(NewDir, 1) = ":" Then NewDir = NewDir & "\"
  OneDirBack = NewDir
End Function


Function MakeININick(sNick As String) As String
  MakeININick = Replace(Replace(sNick, "[", "²"), "]", "³")
End Function


Function MakeNormalNick(sNick As String) As String
  MakeNormalNick = Replace(Replace(sNick, "²", "["), "³", "]")
End Function


Function MakeNormalValue(Value As String) As String
  MakeNormalValue = Replace(Replace(Replace(Replace(Value, "", ""), "", ""), "Ž", ""), "", "")
End Function


Function MakeINIValue(Value As String) As String
  MakeINIValue = Replace(Replace(Replace(Replace(Value, "", ""), "", ""), "", "Ž"), "", "")
End Function


Function FileSendInProgress(RegNick As String) As Boolean
  Dim u As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If LCase(SocketItem(u).RegNick) = LCase(RegNick) Then
        If (GetSockFlag(u, SF_Status) = SF_Status_SendFile) Or (GetSockFlag(u, SF_Status) = SF_Status_SendFileWaiting) Then
          FileSendInProgress = True
          Exit Function
        End If
      End If
    End If
  Next u
  FileSendInProgress = False
End Function


Function AddQueuedFile(TheSock As Long, FileName As String) As Boolean
  Dim u As Long, uCount As Long
  For u = 1 To QueuedFiles
    If LCase(FileSendQueue(u).RegNick) = LCase(SocketItem(TheSock).RegNick) Then
      If LCase(FileSendQueue(u).FileName) = LCase(FileName) Then AddQueuedFile = False: Exit Function
      uCount = uCount + 1
    End If
  Next u
  If uCount >= 20 Then AddQueuedFile = False: Exit Function
  QueuedFiles = QueuedFiles + 1: ReDim Preserve FileSendQueue(UBound(FileSendQueue) + 5)
  FileSendQueue(QueuedFiles).FileName = FileName
  FileSendQueue(QueuedFiles).RegNick = SocketItem(TheSock).RegNick
  FileSendQueue(QueuedFiles).IRCNick = SocketItem(TheSock).IRCNick
  AddQueuedFile = True
End Function


Function SendNextQueuedFile(ByVal RegNick As String) As Boolean
  Dim u As Long, FoundPos As Long, uCount As Long, DCCResult As Boolean
  Do
    DCCResult = True: FoundPos = 0
    For u = 1 To QueuedFiles
      If LCase(FileSendQueue(u).RegNick) = LCase(RegNick) Then
        DCCResult = InitiateDCCSend(FileSendQueue(u).IRCNick, FileSendQueue(u).RegNick, FileSendQueue(u).FileName)
        FoundPos = u
        Exit For
      End If
    Next u
    If FoundPos > 0 Then
      For u = FoundPos To QueuedFiles - 1
        FileSendQueue(u) = FileSendQueue(u + 1)
      Next u
      QueuedFiles = QueuedFiles - 1: ReDim Preserve FileSendQueue(QueuedFiles)
    End If
    If DCCResult = True Then Exit Do
  Loop
End Function


Function CountQueuedFiles(TheSock As Long) As Long
  Dim u As Long, uCount As Long
  For u = 1 To QueuedFiles
    If LCase(FileSendQueue(u).RegNick) = LCase(SocketItem(TheSock).RegNick) Then
      uCount = uCount + 1
    End If
  Next u
  CountQueuedFiles = uCount
End Function


Function AddLineToFile(Line As String, File As String)
  Dim FileNum As Integer, FileData As String
  On Local Error Resume Next
  FileNum = FreeFile
  Open File For Input As FileNum
  If Err.Number = 0 Then
    FileData = Input(LOF(FileNum), FileNum)
    Close FileNum
    FileData = FileData & vbCrLf & Line
    While InStr(1, FileData, vbCrLf & vbCrLf) > 0
      FileData = Replace(FileData, vbCrLf & vbCrLf, vbCrLf)
    Wend
    If Left(FileData, 2) = vbCrLf Then FileData = Mid(FileData, 3)
    SetAttr File, vbArchive
    FileNum = FreeFile
    Open File For Output As FileNum
    Print #FileNum, FileData
    Close FileNum
  Else
    If Err.Number = 53 Then
      FileNum = FreeFile
      Open File For Output As FileNum
      Print #FileNum, Line
      Close FileNum
    ElseIf Err.Number = 76 Then
      MkDir OneDirBack(File)
      FileNum = FreeFile
      Open File For Output As FileNum
      Print #FileNum, Line
      Close FileNum
    End If
    Err.Clear
  End If
End Function


Function LineInFile(Line As String, File As String) As Boolean
  Dim FileNum As Integer, FileData As String
  On Local Error Resume Next
  FileNum = FreeFile
  Open File For Input As FileNum
  If Err.Number > 0 Then
    LineInFile = False
    Err.Clear
  Else
    FileData = Input(LOF(FileNum), FileNum)
    Close FileNum
    If InStr(1, vbCrLf & LCase(FileData) & vbCrLf, vbCrLf & LCase(Line) & vbCrLf, vbBinaryCompare) <> 0 Then LineInFile = True Else LineInFile = False
  End If
End Function


Function RemLineFromFile(Line As String, File As String)
  Dim FileNum As Integer, FileData As String, FileIndex As Long
  On Local Error Resume Next
  FileNum = FreeFile
  Open File For Input As FileNum
  If Err.Number > 0 Then
    Err.Clear
  Else
    FileData = vbCrLf & Input(LOF(FileNum), FileNum) & vbCrLf
    Close FileNum
    FileData = Replace(FileData, vbCrLf & Line & vbCrLf, vbCrLf)
    While InStr(1, FileData, vbCrLf & vbCrLf) > 0
      FileData = Replace(FileData, vbCrLf & vbCrLf, vbCrLf)
    Wend
    If Left(FileData, 2) = vbCrLf Then FileData = Mid(FileData, 3)
    If Right(FileData, 2) = vbCrLf Then FileData = Mid(FileData, 1, Len(FileData) - 2)
    SetAttr File, vbArchive: Err.Clear
    FileNum = FreeFile
    Open File For Output As FileNum
    Print #FileNum, FileData
    Close FileNum
  End If
End Function


Function DirExist(ByVal DirName As String) As Boolean
  If (InStr(DirName, "*") > 0) Or (InStr(DirName, "?") > 0) Then DirExist = False: Exit Function
  If Not Dir(DirName, 30 Or vbDirectory) = "" Then DirExist = True
End Function


Function GetFileName(ByVal PathFile As String) As String
  Dim LastChar As Integer
  If PathFile = "" Then GetFileName = "": Exit Function
  PathFile = Replace(PathFile, "|", "_")
  LastChar = InStrRev(PathFile, "\")
  If LastChar > 1 Then
    GetFileName = Right(PathFile, Len(PathFile) - LastChar)
  Else
    GetFileName = PathFile
  End If
End Function


Public Sub AddAccessedFile(FileName As String) ' : AddStack "Windows_AddAccessedFile(" & FileName & ")"
  If InStr(" " & JustAccessing & " ", " " & LCase(FileName) & " ") = 0 Then
    JustAccessing = JustAccessing & " " & LCase(FileName)
  End If
End Sub

Public Function IsAccessedFile(FileName As String) As Boolean ' : AddStack "Windows_IsAccessedFile(" & FileName & ")"
  IsAccessedFile = (InStr(" " & JustAccessing & " ", " " & LCase(FileName) & " ") > 0)
End Function

Public Sub RemAccessedFile(FileName As String) ' : AddStack "Windows_RemAccessedFile(" & FileName & ")"
  JustAccessing = Trim(Replace(" " & JustAccessing & " ", " " & LCase(FileName) & " ", " "))
End Sub

Public Sub WaitForAccess(FileName As String) ' : AddStack "Windows_WaitForAccess(" & FileName & ")"
Dim t As Currency
  If IsAccessedFile(FileName) = True Then
    t = WinTickCount
    Do
      DoEvents
      If IsAccessedFile(FileName) = False Then Exit Do
    Loop While (WinTickCount - t) < 4000
  End If
End Sub

'Converts a file size to a string
Function SizeToString(ByVal Zahl As Currency) As String
  Dim KbZahl As Long, MbZahl As Currency, GbZahl As Currency
  Dim FunctionResult As String
   KbZahl = 1024&
   MbZahl = 1024& * 1024&
   GbZahl = MbZahl * 1024&

      If Zahl / GbZahl >= 1 Then
         FunctionResult = Format(Zahl / GbZahl, "####.#")
         If Right(FunctionResult, 1) = "." Then FunctionResult = Left(FunctionResult, Len(FunctionResult) - 1)
         If Right(FunctionResult, 1) = "," Then FunctionResult = Left(FunctionResult, Len(FunctionResult) - 1)
         FunctionResult = FunctionResult & "G"
      ElseIf Zahl / MbZahl >= 16 Then
         FunctionResult = Format(Zahl / MbZahl, "####") & "M"
      ElseIf Zahl / MbZahl >= 1 Then
         FunctionResult = Format(Zahl / MbZahl, "####.#")
         If Right(FunctionResult, 1) = "." Then FunctionResult = Left(FunctionResult, Len(FunctionResult) - 1)
         If Right(FunctionResult, 1) = "," Then FunctionResult = Left(FunctionResult, Len(FunctionResult) - 1)
         FunctionResult = FunctionResult & "M"
      ElseIf Zahl / KbZahl >= 1 Then
         FunctionResult = Format(Zahl / KbZahl, "####") & "k"
      Else
         FunctionResult = Format(Zahl, "####") & "b"
      End If

   If FunctionResult = "b" Then FunctionResult = "0b"
   SizeToString = FunctionResult
End Function

'Checks the checksum of a given file
Function ValidChecksum(FileName As String) As Boolean ' : AddStack "Routines_ValidChecksum(" & FileName & ")"
Dim FileNum As Integer, TotalLen As Long, TotalProcessed As Long
Dim CurBufSize As Long, u As Integer, ShortBuf As String
Dim strBuf As String, byteVal As Integer, ChkSum As Long
  On Local Error Resume Next
  FileNum = FreeFile
  Open FileName For Binary As #FileNum
    If Err.Number > 0 Then ValidChecksum = False: Close #FileNum: Exit Function
    TotalLen = LOF(FileNum)
    ChkSum = TotalLen
    CurBufSize = 32000: If TotalProcessed + CurBufSize > TotalLen Then CurBufSize = TotalLen - TotalProcessed
    While CurBufSize > 0
      strBuf = Space(CurBufSize): Get FileNum, , strBuf
      u = 1
      While u < CurBufSize
        If ((TotalProcessed + u) < 79) Then
          byteVal = Asc(Mid(strBuf, u, 1))
          ChkSum = ChkSum + (byteVal * ((u Mod 32) + 1))
        ElseIf ((TotalProcessed + u) > 116) Then
          u = u + 1
          byteVal = Asc(Mid(strBuf, u, 1))
          ChkSum = ChkSum + (byteVal * ((u Mod 32) + 1))
        End If
        If ChkSum > 20000000 Then ChkSum = ChkSum - 20000000
        u = u + 1
      Wend
      TotalProcessed = TotalProcessed + CurBufSize
      CurBufSize = 32000: If TotalProcessed + CurBufSize > TotalLen Then CurBufSize = TotalLen - TotalProcessed
    Wend
    'Calculate checksum
    strBuf = Left(EncryptIt(Hex(ChkSum)), 38)
    strBuf = strBuf + Space(38 - Len(strBuf))
    'Get saved checksum from file
    ShortBuf = Space(38): Get FileNum, 79, ShortBuf
    ValidChecksum = (Trim(ShortBuf) = Trim(strBuf))
  Close #FileNum
End Function

