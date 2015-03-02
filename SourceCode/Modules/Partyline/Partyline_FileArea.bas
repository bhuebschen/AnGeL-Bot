Attribute VB_Name = "Partyline_FileArea"
',-======================- ==-- -  -
'|   AnGeL - Partyline - FileArea
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Type FAFiles
  FileName As String
  FileSize As String
  CreatedBy As String
  Directory As Boolean
End Type


Sub FileArea_Load()
'
End Sub


Sub FileArea_Unload()
'
End Sub


Sub SpreadLevelFileAreaMessage(vsock As Long, Line As String)
  Dim u As Long
  For u = 1 To SocketCount
    If IsValidSocket(u) Then If vsock <> u And ((GetSockFlag(u, SF_Status) = SF_Status_Party) Or (GetSockFlag(u, SF_Status) = SF_Status_FileArea)) And MatchFlags(SocketItem(u).Flags, "+m") Then TU u, Line
  Next u
End Sub

'File area
Public Sub FileArea(vsock As Long, Line As String) ' : AddStack "FileAreaHandler_FileArea(" & vsock & ", " & Line & ")"
Dim u As Long, Rest As String, FileName As String, FileCounter As Long
Dim Errors As Boolean, ToDir As String, ToFADir As String, NotDid As Boolean
Dim SwapStr As String, ScNum As Long
  If Left(Line, 1) = "." Then Line = Right(Line, Len(Line) - 1)
  HaltDefault = False
  For ScNum = 1 To ScriptCount
    If Scripts(ScNum).Hooks.fa_command Then
      RunScriptX ScNum, "fa_command", vsock, SocketItem(vsock).RegNick, Line
    End If
  Next ScNum
  If HaltDefault = True Then Exit Sub
  Select Case LCase(Param(Line, 1))
    Case "quit", "exit", "q", "x", "0"
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Hooks.fa_userleft Then
          RunScriptX ScNum, "fa_userleft", vsock, SocketItem(vsock).RegNick
        End If
      Next ScNum
      NotDid = True
      TU vsock, "10*** Bringing you back to the party line..."
      SetSockFlag vsock, SF_Status, SF_Status_Party
      SetAway vsock, ""
    Case "help"
      NotDid = True
      TU vsock, "*** FILE AREA COMMANDS for " & BotNetNick & ", AnGeL " & BotVersion + IIf(ServerNetwork <> "", "+" & ServerNetwork, "") & ":"
      TU vsock, "2  dir, ls (mask)     14Shows a list of all files in the current"
      TU vsock, "2                     14directory. You can specify a mask like"
      TU vsock, "2                     14'*.zip' to get zip files only."
      TU vsock, "2  cd <new dir>       14Changes your current file area directory"
      TU vsock, "2                     14to <new dir>."
      TU vsock, "2  get <file(s)>      14Sends the file <file> to you. You can"
      TU vsock, "2                     14specify a mask like '*.txt' to make me"
      TU vsock, "2                     14send you all text files."
      TU vsock, "2  quit, exit, q, x   14Leaves the file area and brings you back"
      TU vsock, "2                     14to the party line."
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        TU vsock, "For file area janitors:"
        TU vsock, "2  del, rm <file(s)>"
        TU vsock, "14     Removes/deletes files from the file area."
        TU vsock, "2  move, mv <file(s)> <new directory>"
        TU vsock, "14     Moves files from the current directory to another. *"
        TU vsock, "2  copy, cp <file(s)> <new directory>"
        TU vsock, "14     Copies files from the current directory to another. *"
        TU vsock, "2  ren <filename> <new filename>"
        TU vsock, "14     Renames a file from <filename> to <new filename>. *"
        TU vsock, "2  mkdir, md <new directory>"
        TU vsock, "14     Creates a new directory in the file area."
        TU vsock, "2  rmdir, rd <directory>"
        TU vsock, "14     Removes an empty directory from the file area."
        TU vsock, EmptyLine
        TU vsock, "14* Please don't use spaces in the first parameter! You can"
        TU vsock, "14  use the '?' character to replace spaces."
      End If
      TU vsock, EmptyLine
    Case "dir", "ls"
      NotDid = True
      If Param(Line, 2) <> "" Then
        Rest = GetNewDir(SocketItem(vsock).FileAreaDir, GetRest(Line, 2), "\")
        ListFiles vsock, Rest
      Else
        ListFiles vsock, "*.*"
      End If
    Case "cd", "cd..", "cd\"
      If LCase(Line) = "cd.." Then Line = "cd .."
      If LCase(Line) = "cd\" Then Line = "cd \"
      If Param(Line, 2) = "" Then TU vsock, "5*** Usage: cd <new directory>": Exit Sub
      Rest = GetNewDir(SocketItem(vsock).FileAreaDir, GetRest(Line, 2), "\")
      If LCase(Rest) = "\scripts" Or LCase(Rest) = "\scripts" Then
        If MatchFlags(SocketItem(vsock).Flags, "-s") Then
          TU vsock, "5*** Sorry, you don't have access to my scripts (+s needed).": Exit Sub
        End If
      End If
      If Right(Rest, 1) = "\" Then Rest = FileAreaHome + Rest Else Rest = FileAreaHome + Rest & "\"
      If DirExist(Rest) = False Then
        TU vsock, "5*** Directory not found.": NotDid = True
      Else
        SocketItem(vsock).FileAreaDir = GetNewDir(SocketItem(vsock).FileAreaDir, Right(Line, Len(Line) - Len(Param(Line, 1)) - 1), "\")
        TU vsock, "2*** Current directory: " & SocketItem(vsock).FileAreaDir
      End If
    Case "abort"
      If Param(Line, 2) = "" Then TU vsock, "5*** Usage: abort <filename>": Exit Sub
      For FileCounter = 0 To QueuedFiles - 1
        If LCase(FileSendQueue(FileCounter).FileName) = LCase(Param(Line, 2)) Then
          For u = FileCounter To QueuedFiles - 1
            FileSendQueue(u) = FileSendQueue(u + 1)
          Next u
          QueuedFiles = QueuedFiles - 1
        End If
      Next FileCounter
      ReDim Preserve FileSendQueue(QueuedFiles)
      If QueuedFiles < FileCounter Then TU vsock, "3*** Removed file." Else TU vsock, "5*** No matching files found."
    Case "get"
      If Param(Line, 2) = "" Then TU vsock, "5*** Usage: get <filename or wildcard mask> (ircnick)": Exit Sub
      If Param(Line, 3) <> "" Then SocketItem(vsock).IRCNick = Param(Line, 3) Else
      If SocketItem(vsock).IRCNick = "" Then TU vsock, "5*** Usage: get <filename or wildcard mask> <ircnick>": Exit Sub
      Rest = GetFileName(Param(Line, 2))
      If InStr(Rest, "..") > 0 Then TU vsock, "5*** Usage: get <filename or wildcard mask>": Exit Sub
      On Local Error Resume Next
      If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = Left(FileAreaHome, Len(FileAreaHome) - 1) & SocketItem(vsock).FileAreaDir & Rest Else Rest = Left(FileAreaHome, Len(FileAreaHome) - 1) & SocketItem(vsock).FileAreaDir & "\" & Rest
      FileName = Dir(Rest)
      If FileName = "" Then TU vsock, "5*** No matching files found.": Exit Sub
      Errors = False
      Do
        If FileName <> "" Then
          If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + FileName Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & FileName
          If AddQueuedFile(vsock, Rest) = False Then Errors = True Else FileCounter = FileCounter + 1
        Else
          Exit Do
        End If
        FileName = Dir
      Loop
      If Errors = True Then
        If CountQueuedFiles(vsock) = 30 Then
          TU vsock, "5*** Too many files in your send queue (30)."
        Else
          If FileCounter = 0 Then
            TU vsock, "5*** The file(s) you specified are already queued up.": NotDid = True
          Else
            TU vsock, "5*** Some of the files you specified are already queued up."
          End If
        End If
      End If
      If FileCounter > 0 Then
        If Not FileSendInProgress(SocketItem(vsock).RegNick) Then
          SendNextQueuedFile SocketItem(vsock).RegNick
          TU vsock, "3*** Sending " & CStr(FileCounter) & " matching file" & IIf(FileCounter = 1, "", "s") & " to " & SocketItem(vsock).IRCNick & "..."
        Else
          TU vsock, "3*** Added " & CStr(FileCounter) & " matching file" & IIf(FileCounter = 1, "", "s") & " to your send queue."
        End If
      End If
      Err.Clear
      On Error GoTo 0
    Case "clearqueue"
      For FileCounter = 0 To QueuedFiles - 1
        If FileSendQueue(FileCounter).RegNick = SocketItem(vsock).RegNick Then
          For u = FileCounter To QueuedFiles - 1
            FileSendQueue(u) = FileSendQueue(u + 1)
          Next u
          QueuedFiles = QueuedFiles - 1
        End If
      Next FileCounter
      ReDim Preserve FileSendQueue(QueuedFiles)
      TU vsock, "3*** send queue cleared"
    Case "viewqueue"
      u = 0
      For FileCounter = 0 To QueuedFiles - 1
        If FileSendQueue(FileCounter).RegNick = SocketItem(vsock).RegNick Then
          u = 1
          TU vsock, "14-3 " & GetFileName(FileSendQueue(FileCounter).FileName)
        End If
      Next FileCounter
      If u = 0 Then TU vsock, "5*** You don't have any files in queue."
    Case "del", "rm"
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        If Param(Line, 2) = "" Then TU vsock, "5*** Usage: del, rm <filename or wildcard mask>": Exit Sub
        Rest = GetFileName(Right(Line, Len(Line) - Len(Param(Line, 1)) - 1))
        If InStr(Rest, "..") > 0 Then TU vsock, "5*** Usage: del, rm <filename or wildcard mask>": Exit Sub
        If MatchFlags(SocketItem(vsock).Flags, "-s") Then
          If LCase(SocketItem(vsock).FileAreaDir) = "\logs" Then
            TU vsock, "5*** You don't have the permission to delete logs."
            Exit Sub
          End If
        End If
        On Local Error Resume Next
        If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + Rest Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & Rest
        FileName = Dir(Rest)
        If FileName = "" Then TU vsock, "5*** No matching files found.": Exit Sub
        Do
          If FileName <> "" Then
            If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + FileName Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & FileName
            Err.Clear
            SetAttr Rest, vbNormal
            Kill Rest
            If Err.Number = 0 Then
              FileCounter = FileCounter + 1
              For u = 1 To ScriptCount
                If LCase(FileAreaHome & "Scripts\" & Scripts(u).Name) = LCase(Rest) Then
                  TU vsock, "14*** Unloading script: " & GetFileName(Rest) & ""
                  RemScript u
                  Exit For
                End If
              Next u
              TU vsock, "14*** Deleted: " & GetFileName(Rest) & ""
              DeletePPString SocketItem(vsock).FileAreaDir, GetFileName(Rest), HomeDir & "Files.ini"
            Else
              TU vsock, "5*** Couldn't delete '" & GetFileName(Rest) & "': " & Err.Description & ""
            End If
          Else
            Exit Do
          End If
          FileName = Dir
        Loop
        If FileCounter > 0 Then
          TU vsock, "3*** Deleted " & CStr(FileCounter) & " matching file" & IIf(FileCounter = 1, "", "s") & "."
        End If
      Else
        TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      End If
      Err.Clear
      On Error GoTo 0
    Case "move", "mv"
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        On Local Error Resume Next
        If Param(Line, 3) = "" Then TU vsock, "5*** Usage: move, mv <filename or wildcard mask (no spaces please!)> <new directory>": Exit Sub
        ToDir = GetNewDir(SocketItem(vsock).FileAreaDir, Right(Line, Len(Line) - Len(Param(Line, 1)) - Len(Param(Line, 2)) - 2), "\")
        ToFADir = ToDir
        If Right(ToDir, 1) = "\" Then ToDir = FileAreaHome + ToDir Else ToDir = FileAreaHome + ToDir & "\"
        If Right(ToFADir, 1) <> "\" Then ToFADir = ToFADir & "\"
        If DirExist(ToDir) = False Then
          TU vsock, "5*** Target directory not found.": NotDid = True
        Else
          Rest = GetFileName(Param(Line, 2))
          If InStr(Rest, "..") > 0 Then TU vsock, "5*** Usage: move, mv <filename or wildcard mask (no spaces please!)> <new directory>": Exit Sub
          If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + Rest Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & Rest
          If MatchFlags(SocketItem(vsock).Flags, "-s") Then
            If LCase(SocketItem(vsock).FileAreaDir) = "\logs" Then
              TU vsock, "5*** You don't have the permission to move logs."
              Exit Sub
            End If
          End If
          FileName = Dir(Rest)
          If FileName = "" Then TU vsock, "5*** No matching files found.": Exit Sub
          Do
            If FileName <> "" Then
              If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + FileName Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & FileName
              Err.Clear
              SetAttr Rest, vbNormal
              Name Rest As ToDir + FileName
              If Err.Number = 0 Then
                FileCounter = FileCounter + 1
                For u = 1 To ScriptCount
                  If LCase(FileAreaHome & "Scripts\" & Scripts(u).Name) = LCase(Rest) Then
                    TU vsock, "14*** Unloading script: " & GetFileName(Rest) & ""
                    RemScript u
                    Exit For
                  End If
                Next u
                TU vsock, "14*** Moved: " & GetFileName(Rest) & " -> " & ToFADir + FileName & ""
                SwapStr = GetPPString(SocketItem(vsock).FileAreaDir, GetFileName(Rest), "", HomeDir & "Files.ini")
                If SwapStr <> "" Then
                  DeletePPString SocketItem(vsock).FileAreaDir, GetFileName(Rest), HomeDir & "Files.ini"
                  WritePPString IIf(Len(ToFADir) > 1, Left(ToFADir, Len(ToFADir) - 1), ToFADir), FileName, SwapStr, HomeDir & "Files.ini"
                End If
              Else
                TU vsock, "5*** Couldn't move '" & GetFileName(Rest) & "': " & Err.Description & ""
              End If
            Else
              Exit Do
            End If
            FileName = Dir
          Loop
          If FileCounter > 0 Then
            TU vsock, "3*** Moved " & CStr(FileCounter) & " matching file" & IIf(FileCounter = 1, "", "s") & "."
          End If
        End If
      Else
        TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      End If
      Err.Clear
      On Error GoTo 0
    Case "copy", "cp"
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        On Local Error Resume Next
        If Param(Line, 3) = "" Then TU vsock, "5*** Usage: copy, cp <filename or wildcard mask (no spaces please!)> <new directory>": Exit Sub
        ToDir = GetNewDir(SocketItem(vsock).FileAreaDir, Right(Line, Len(Line) - Len(Param(Line, 1)) - Len(Param(Line, 2)) - 2), "\")
        ToFADir = ToDir
        If MatchFlags(SocketItem(vsock).Flags, "-s") Then
          If LCase(ToFADir) = "\logs" Then
            TU vsock, "5*** You don't have the permission to copy something into the logs directory."
            Exit Sub
          End If
        End If
        If Right(ToDir, 1) = "\" Then ToDir = FileAreaHome + ToDir Else ToDir = FileAreaHome + ToDir & "\"
        If Right(ToFADir, 1) <> "\" Then ToFADir = ToFADir & "\"
        If DirExist(ToDir) = False Then
          TU vsock, "5*** Target directory not found.": NotDid = True
        Else
          Rest = GetFileName(Param(Line, 2))
          If InStr(Rest, "..") > 0 Then TU vsock, "5*** Usage: copy, cp <filename or wildcard mask (no spaces please!)> <new directory>": Exit Sub
          If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + Rest Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & Rest
          FileName = Dir(Rest)
          If FileName = "" Then TU vsock, "5*** No matching files found.": Exit Sub
          Do
            If FileName <> "" Then
              If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + FileName Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & FileName
              Err.Clear
              FileCopy Rest, ToDir + FileName
              If Err.Number = 0 Then
                FileCounter = FileCounter + 1
                TU vsock, "14*** Copied: " & GetFileName(Rest) & " -> " & ToFADir + FileName & ""
                SwapStr = GetPPString(SocketItem(vsock).FileAreaDir, GetFileName(Rest), "", HomeDir & "Files.ini")
                If SwapStr <> "" Then
                  WritePPString IIf(Len(ToFADir) > 1, Left(ToFADir, Len(ToFADir) - 1), ToFADir), FileName, SwapStr, HomeDir & "Files.ini"
                End If
              Else
                TU vsock, "5*** Couldn't copy '" & GetFileName(Rest) & "': " & Err.Description & ""
              End If
            Else
              Exit Do
            End If
            FileName = Dir
          Loop
          If FileCounter > 0 Then
            TU vsock, "3*** Copied " & CStr(FileCounter) & " matching file" & IIf(FileCounter = 1, "", "s") & "."
          End If
        End If
      Else
        TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      End If
      Err.Clear
      On Error GoTo 0
    Case "ren"
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        On Local Error Resume Next
        If Param(Line, 3) = "" Then TU vsock, "5*** Usage: ren <filename> <new filename>": Exit Sub
        Rest = GetFileName(Param(Line, 2))
        ToDir = GetFileName(Right(Line, Len(Line) - Len(Param(Line, 1)) - Len(Param(Line, 2)) - 2))
        If InStr(ToDir, "*") > 0 Or InStr(ToDir, "?") > 0 Then TU vsock, "5*** You can't use wildcards ('*','?') in the destination filename.": Exit Sub
        ToFADir = ToDir
        If InStr(Rest & " " & ToDir, "..") > 0 Then TU vsock, "5*** Usage: ren <filename> <new filename>": Exit Sub
        If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + Rest Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & Rest
        If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then ToDir = FileAreaHome + SocketItem(vsock).FileAreaDir + ToDir Else ToDir = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & ToDir
        FileName = Dir(Rest)
        If FileName = "" Then TU vsock, "5*** File not found.": Exit Sub
        If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then Rest = FileAreaHome + SocketItem(vsock).FileAreaDir + FileName Else Rest = FileAreaHome + SocketItem(vsock).FileAreaDir & "\" & FileName
        Err.Clear
        SetAttr Rest, vbNormal
        Name Rest As ToDir
        If Err.Number = 0 Then
          TU vsock, "3*** Renamed: " & GetFileName(Rest) & " -> " & ToFADir & ""
          For u = 1 To ScriptCount
            If LCase(FileAreaHome & "Scripts\" & Scripts(u).Name) = LCase(Rest) Then
              TU vsock, "14*** Updating script filename: " & GetFileName(Rest) & " -> " & GetFileName(ToFADir) & ""
              RenScript u, GetFileName(ToFADir)
              Exit For
            End If
          Next u
          SwapStr = GetPPString(SocketItem(vsock).FileAreaDir, GetFileName(Rest), "", HomeDir & "Files.ini")
          If SwapStr <> "" Then
            DeletePPString SocketItem(vsock).FileAreaDir, GetFileName(Rest), HomeDir & "Files.ini"
            WritePPString SocketItem(vsock).FileAreaDir, GetFileName(ToFADir), SwapStr, HomeDir & "Files.ini"
          End If
        Else
          TU vsock, "5*** Couldn't rename '" & GetFileName(Rest) & "': " & Err.Description & "": NotDid = True
        End If
      Else
        TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      End If
      Err.Clear
      On Error GoTo 0
    Case "mkdir", "md"
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        On Local Error Resume Next
        If Param(Line, 2) = "" Then TU vsock, "5*** Usage: mkdir, md <new directory>": Exit Sub
        ToDir = GetNewDir(SocketItem(vsock).FileAreaDir, Right(Line, Len(Line) - Len(Param(Line, 1)) - 1), "\")
        ToFADir = ToDir
        If Right(ToDir, 1) = "\" Then ToDir = FileAreaHome + ToDir Else ToDir = FileAreaHome + ToDir & "\"
        If Right(ToFADir, 1) = "\" Then ToFADir = Left(ToFADir, Len(ToFADir) - 1)
        If DirExist(ToDir) = True Then
          TU vsock, "5*** Target directory already exists.": NotDid = True
        Else
          MkDir ToDir
          If Err.Number = 0 Then
            TU vsock, "3*** Created: " & ToFADir & ""
            WritePPString OneDirBack(ToFADir), GetFileName(Right(ToFADir, Len(ToFADir) - 1)), SocketItem(vsock).RegNick, HomeDir & "Files.ini"
          Else
            TU vsock, "5*** Couldn't create '" & ToFADir & "': " & Err.Description & "": NotDid = True
          End If
        End If
      Else
        TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      End If
      Err.Clear
      On Error GoTo 0
    Case "rmdir", "rd"
      If MatchFlags(SocketItem(vsock).Flags, "+j") Then
        On Local Error Resume Next
        If Param(Line, 2) = "" Then TU vsock, "5*** Usage: rmdir, rd <directory>": Exit Sub
        ToDir = GetNewDir(SocketItem(vsock).FileAreaDir, Right(Line, Len(Line) - Len(Param(Line, 1)) - 1), "\")
        ToFADir = ToDir
        If Right(ToDir, 1) = "\" Then ToDir = FileAreaHome + ToDir Else ToDir = FileAreaHome + ToDir & "\"
        If Right(ToFADir, 1) = "\" Then ToFADir = Left(ToFADir, Len(ToFADir) - 1)
        If DirExist(ToDir) = False Then
          TU vsock, "5*** Target directory not found.": NotDid = True
        Else
          If Dir(ToDir & "*.*") <> "" Then TU vsock, "5*** Target directory has to be empty!": Exit Sub
          RmDir ToDir
          If Err.Number = 0 Then
            TU vsock, "3*** Removed: " & ToFADir & ""
            DeletePPString OneDirBack(ToFADir), GetFileName(Right(ToFADir, Len(ToFADir) - 1)), HomeDir & "Files.ini"
          Else
            TU vsock, "5*** Couldn't remove '" & ToFADir & "': " & Err.Description & "": NotDid = True
          End If
        End If
      Else
        TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      End If
      Err.Clear
      On Error GoTo 0
    Case Else
      TU vsock, MakeMsg(MSG_FA_LookHelp, Line): NotDid = True
      TU vsock, "5*** If you want to use bot commands, leave the file area first."
  End Select
  If Not NotDid Then SpreadLevelFileAreaMessage vsock, "14[" & Time & "] +++ " & SocketItem(vsock).RegNick & " did: " & Line
End Sub

'List file area directory
Public Sub ListFiles(vsock As Long, ByVal Mask As String) ' : AddStack "FileAreaHandler_ListFiles(" & vsock & ", " & Mask & ")"
Dim Nr As Long, erg As Long, Attr As WIN32_FIND_DATA, FileSize As Currency
Dim FileName As String, FoundFiles() As FAFiles, FoundFileCount As Long, BaseDir As String
Dim u As Long, AddAt As Long, FileCount As Long, DirCount As Long, TotalSize As Currency
Dim FADir As String
  ReDim Preserve FoundFiles(5)
  On Local Error Resume Next
  If InStr(Mask, "\") = 0 Then
    If Right(SocketItem(vsock).FileAreaDir, 1) = "\" Then
      BaseDir = FileAreaHome + SocketItem(vsock).FileAreaDir
      If Left(Mask, 1) <> "\" Then
        FADir = SocketItem(vsock).FileAreaDir + Mask
      Else
        FADir = SocketItem(vsock).FileAreaDir + Mid(Mask, 2)
      End If
    Else
      BaseDir = FileAreaHome + SocketItem(vsock).FileAreaDir & "\"
      If Left(Mask, 1) <> "\" Then
        FADir = SocketItem(vsock).FileAreaDir & "\" & Mask
      Else
        FADir = SocketItem(vsock).FileAreaDir + Mask
      End If
    End If
    Nr = kernel32_FindFirstFileA(BaseDir + Mask, Attr)
  Else
    If Right(Mask, 1) = "\" Then Mask = Mask & "*.*"
    BaseDir = FileAreaHome + Mask
    FADir = Mask
    Nr = kernel32_FindFirstFileA(BaseDir, Attr)
  End If
  If Nr <> -1 Then
    Do
      If InStr(Attr.cFileName, Chr(0)) > 0 Then FileName = Left(Attr.cFileName, InStr(Attr.cFileName, Chr(0)) - 1) Else FileName = Attr.cFileName
      If (FileName <> ".") And Not (FileName = ".." And SocketItem(vsock).FileAreaDir = "\") And Not (FileName = "") Then
        If (Attr.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
          FoundFileCount = FoundFileCount + 1
          If FoundFileCount > UBound(FoundFiles()) Then ReDim Preserve FoundFiles(UBound(FoundFiles()) + 5)
          AddAt = FoundFileCount
          For u = 1 To FoundFileCount - 1
            If ((UCase(FoundFiles(u).FileName) > UCase(FileName)) And (FoundFiles(u).Directory = True)) Or (FoundFiles(u).Directory = False) Then AddAt = u: Exit For
          Next u
          For u = FoundFileCount To AddAt + 1 Step -1
            FoundFiles(u) = FoundFiles(u - 1)
          Next u
          FoundFiles(AddAt).FileName = FileName
          FoundFiles(AddAt).FileSize = "<DIR>"
          FoundFiles(AddAt).Directory = True
          DirCount = DirCount + 1
        Else
          FoundFileCount = FoundFileCount + 1
          If FoundFileCount > UBound(FoundFiles()) Then ReDim Preserve FoundFiles(UBound(FoundFiles()) + 5)
          AddAt = FoundFileCount
          For u = 1 To FoundFileCount - 1
            If ((UCase(FoundFiles(u).FileName) > UCase(FileName)) And (FoundFiles(u).Directory = False)) And Not (FoundFiles(u).Directory = True) Then AddAt = u: Exit For
          Next u
          For u = FoundFileCount To AddAt + 1 Step -1
            FoundFiles(u) = FoundFiles(u - 1)
          Next u
          FoundFiles(AddAt).FileName = FileName
          FileSize = Attr.nFileSizeHigh * 65535 + Attr.nFileSizeLow
          FoundFiles(AddAt).FileSize = SizeToString(FileSize)
          FoundFiles(AddAt).Directory = False
          FileCount = FileCount + 1
          TotalSize = TotalSize + FileSize
        End If
      End If
      erg = kernel32_FindNextFileA(Nr, Attr)
      If erg = 0 Then Exit Do
    Loop
    kernel32_FindClose Nr
  End If
  If (FoundFileCount = 1) And (LCase(GetFileName(Mask)) = LCase(FoundFiles(1).FileName)) And (FoundFiles(1).Directory = True) Then
    ListFiles vsock, Mask & "\*.*"
    Exit Sub
  End If
  If FoundFileCount > 0 Then
    If GetSockFlag(vsock, SF_Colors) = SF_YES Then
      TU vsock, "2*** Directory of '" & FADir & "'"
      TU vsock, "0,1 File or directory name:                  | Size:  | Created by:  "
    Else
      TU vsock, "*** Directory of '" & FADir & "'"
      TU vsock, " File or directory name:                    Size:    Created by:  "
      TU vsock, " -----------------------------------------  -------  -------------"
    End If
    For u = 1 To FoundFileCount
      FoundFiles(u).CreatedBy = GetPPString(SocketItem(vsock).FileAreaDir, FoundFiles(u).FileName, "14" & BotNetNick & "", HomeDir & "Files.ini")
      If GetSockFlag(vsock, SF_Colors) = SF_YES Then
        If FoundFiles(u).Directory Then
          If Len(FoundFiles(u).FileName) > 40 Then
            TU vsock, "3 " & FoundFiles(u).FileName
            TU vsock, Space(40) & "14 |3  <DIR>14 |3 " & Strip(FoundFiles(u).CreatedBy)
          Else
            TU vsock, "3 " & Spaces2(40, FoundFiles(u).FileName) & "14 |3  <DIR>14 |3 " & Strip(FoundFiles(u).CreatedBy)
          End If
        Else
          If Len(FoundFiles(u).FileName) > 40 Then
            TU vsock, " " & FoundFiles(u).FileName
            TU vsock, Space(40) & "14 | " & Spaces(6, FoundFiles(u).FileSize) + FoundFiles(u).FileSize & "14 | " & FoundFiles(u).CreatedBy
          Else
            TU vsock, " " & Spaces2(40, FoundFiles(u).FileName) & "14 | " & Spaces(6, FoundFiles(u).FileSize) + FoundFiles(u).FileSize & "14 | " & FoundFiles(u).CreatedBy
          End If
        End If
      Else
        If FoundFiles(u).Directory Then
          If Len(FoundFiles(u).FileName) > 40 Then
            TU vsock, " " & FoundFiles(u).FileName
            TU vsock, Space(40) & "    <DIR>   " & FoundFiles(u).CreatedBy
          Else
            TU vsock, " " & Spaces2(40, FoundFiles(u).FileName) & "    <DIR>   " & FoundFiles(u).CreatedBy
          End If
        Else
          If Len(FoundFiles(u).FileName) > 40 Then
            TU vsock, " " & FoundFiles(u).FileName
            TU vsock, Space(40) & "   " & Spaces(6, FoundFiles(u).FileSize) + FoundFiles(u).FileSize & "   " & FoundFiles(u).CreatedBy
          Else
            TU vsock, " " & Spaces2(40, FoundFiles(u).FileName) & "   " & Spaces(6, FoundFiles(u).FileSize) + FoundFiles(u).FileSize & "   " & FoundFiles(u).CreatedBy
          End If
        End If
      End If
    Next u
    TU vsock, "2*** " & CStr(FileCount) & " file" & IIf(FileCount = 1, "", "s") & ", " & CStr(DirCount) & " director" & IIf(DirCount = 1, "y", "ies") & "  (Total size: " & SizeToString(TotalSize) & ")"
    TU vsock, EmptyLine
  Else
    TU vsock, "5*** No matching files found."
  End If
End Sub

