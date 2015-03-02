Attribute VB_Name = "Scripting_Scripts"
',-======================- ==-- -  -
'|   AnGeL - Scripting - Scripts
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public Type BindType
  Chan_msg As Boolean
  Priv_msg As Boolean
  Chan_act As Boolean
  Priv_act As Boolean
  Priv_ctcp As Boolean
  Priv_ctcpreply As Boolean
  Priv_notice As Boolean
  Chan_notice As Boolean
  Chan_ctcp As Boolean
  Server_notice As Boolean
  Join As Boolean
  Part As Boolean
  Quit As Boolean
  Kick As Boolean
  Nick As Boolean
  Commands As Boolean
  Op As Boolean
  Deop As Boolean
  Mode As Boolean
  ModeEnd As Boolean
  Topic As Boolean
  Raw As Boolean
  RawFilter As String
  Numerics As Boolean
  Botnet As Boolean
  BN_Msg As Boolean
  Whois As Boolean
  Resolves As Boolean
  Party_out As Boolean
  AddedUser As Boolean
  RemovedUser As Boolean
  AddedHost As Boolean
  RemovedHost As Boolean
  ChangedNick As Boolean
  PLJoin As Boolean
  Unload As Boolean
  Srv_Connect As Boolean
  fa_downloadbegin As Boolean
  fa_downloadcomplete As Boolean
  fa_uploadbegin As Boolean
  fa_uploadcomplete As Boolean
  fa_userjoin As Boolean
  fa_userleft As Boolean
  fa_command As Boolean
  seen As Boolean
  KI As Boolean
  Ban As Boolean
  UnBan As Boolean
End Type


Public Type ScriptType
  Name As String
  Description As String
  Silent As Boolean
  Script As Object
  Hooks As BindType
  SecurityViolation As Boolean
End Type


Public Scripts() As ScriptType
Public ScriptCount As Long
Public DisableScripts As Boolean
Public HaltDefault As Boolean
Public CalledByScript As Boolean
Public DontLoadScript As String
Public RunScriptIndex As Long


Sub Scripts_Load()
  ReDim Preserve Scripts(5)
End Sub


Sub Scripts_Unload()
'
End Sub

Public Sub RunScript(ScNum As Long, Command As String)
  Dim msg As String, FileNum As Integer, u As Long, DummyVal As String
  On Error GoTo ScriptErr
  ScriptCMDs.CurrentScript = ScNum
  RunScriptIndex = RunScriptIndex + 1
  Scripts(ScNum).Script.executestatement Command
  RunScriptIndex = RunScriptIndex - 1
  If RunScriptIndex = 0 Then ScriptCMDs.CurrentScript = 0
  If Scripts(ScNum).SecurityViolation = True Then
    RemScript ScNum
  End If
  Exit Sub
ScriptErr:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  RunScriptIndex = RunScriptIndex - 1
  If RunScriptIndex = 0 Then ScriptCMDs.CurrentScript = 0
  If Scripts(ScNum).Script.error.Number > 0 Then
    SpreadFlagMessage 0, "+n", "3*** Scripting error in '" & Scripts(ScNum).Name & "':"
    If Scripts(ScNum).Script.error.Number & " (" & Scripts(ScNum).Script.error.Description & ")" <> "0 ()" Then
      SpreadFlagMessage 0, "+n", "10    Error   : " & Scripts(ScNum).Script.error.Number & " (" & Scripts(ScNum).Script.error.Description & ")"
      SpreadFlagMessage 0, "+n", "10    Position: Line " & Scripts(ScNum).Script.error.Line & ", Column " & Scripts(ScNum).Script.error.Column
      FileNum = FreeFile
      On Local Error Resume Next
      Open FileAreaHome & "Scripts\" & Scripts(ScNum).Name For Input As #FileNum
      If Err.Number = 0 Then
        For u = 1 To Scripts(ScNum).Script.error.Line
          If Not EOF(FileNum) Then Line Input #FileNum, DummyVal
        Next u
        SpreadFlagMessage 0, "+n", "10    Excerpt : " & DummyVal
      End If
      Close #FileNum
      SpreadFlagMessage 0, "+n", "10    Command : " & Command
    Else
      SpreadFlagMessage 0, "+n", "10    Error   : The script took too long to execute and was aborted."
    End If
    SpreadFlagMessage 0, "+n", "3*** End of error message"
  Else
    SpreadFlagMessage 0, "+n", "3*** Error while executing script '" & Scripts(ScNum).Name & "':"
    SpreadFlagMessage 0, "+n", "10    Error   : " & CStr(ErrNumber) & " (" & ErrDescription & ")"
    SpreadFlagMessage 0, "+n", "10    Command : " & Command
    SpreadFlagMessage 0, "+n", "3*** End of error message"
  End If
  If Scripts(ScNum).SecurityViolation = True Then
    RemScript ScNum
  End If
  If Err.Number > 0 Then Err.Clear
End Sub

Public Sub RunScriptX(ScNum As Long, Command As String, ParamArray Parameter() As Variant)
  Dim msg As String, FileNum As Integer, u As Long, DummyVal As String
  On Error GoTo ScriptErr
  ScriptCMDs.CurrentScript = ScNum
  RunScriptIndex = RunScriptIndex + 1
  Select Case UBound(Parameter)
    Case 0
      Scripts(ScNum).Script.Run Command, Parameter(0)
    Case 1
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1)
    Case 2
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2)
    Case 3
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3)
    Case 4
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3), Parameter(4)
    Case 5
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3), Parameter(4), Parameter(5)
    Case 6
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3), Parameter(4), Parameter(5), Parameter(6)
    Case 7
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3), Parameter(4), Parameter(5), Parameter(6), Parameter(7)
    Case 8
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3), Parameter(4), Parameter(5), Parameter(6), Parameter(7), Parameter(8)
    Case 9
      Scripts(ScNum).Script.Run Command, Parameter(0), Parameter(1), Parameter(2), Parameter(3), Parameter(4), Parameter(5), Parameter(6), Parameter(7), Parameter(8), Parameter(9)
    Case Else
      If UBound(Parameter) > 9 Then Trace UBound(Parameter)
      Scripts(ScNum).Script.Run Command
  End Select
  RunScriptIndex = RunScriptIndex - 1
  If RunScriptIndex = 0 Then ScriptCMDs.CurrentScript = 0
  If Scripts(ScNum).SecurityViolation = True Then
    RemScript ScNum
  End If
  Exit Sub
ScriptErr:
  Dim ErrNumber As Long, ErrDescription As String, CommandLine As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  If UBound(Parameter) = -1 Then
    CommandLine = Command
  Else
    CommandLine = Command & " (" & Join(Parameter, " | ") & ")"
  End If
  RunScriptIndex = RunScriptIndex - 1
  If RunScriptIndex = 0 Then ScriptCMDs.CurrentScript = 0
  If Scripts(ScNum).Script.error.Number > 0 Then
    SpreadFlagMessage 0, "+n", "3*** Scripting error in '" & Scripts(ScNum).Name & "':"
    If Scripts(ScNum).Script.error.Number & " (" & Scripts(ScNum).Script.error.Description & ")" <> "0 ()" Then
      SpreadFlagMessage 0, "+n", "10    Error   : " & Scripts(ScNum).Script.error.Number & " (" & Scripts(ScNum).Script.error.Description & ")"
      SpreadFlagMessage 0, "+n", "10    Position: Line " & Scripts(ScNum).Script.error.Line & ", Column " & Scripts(ScNum).Script.error.Column
      FileNum = FreeFile
      On Local Error Resume Next
      Open FileAreaHome & "Scripts\" & Scripts(ScNum).Name For Input As #FileNum
      If Err.Number = 0 Then
        For u = 1 To Scripts(ScNum).Script.error.Line
          If Not EOF(FileNum) Then Line Input #FileNum, DummyVal
        Next u
        SpreadFlagMessage 0, "+n", "10    Excerpt : " & DummyVal
      End If
      Close #FileNum
      SpreadFlagMessage 0, "+n", "10    Command : " & CommandLine
    Else
      SpreadFlagMessage 0, "+n", "10    Error   : The script took too long to execute and was aborted."
    End If
    SpreadFlagMessage 0, "+n", "3*** End of error message"
  Else
    SpreadFlagMessage 0, "+n", "3*** Error while executing script '" & Scripts(ScNum).Name & "':"
    SpreadFlagMessage 0, "+n", "10    Error   : " & ErrNumber & " (" & ErrDescription & ")"
    SpreadFlagMessage 0, "+n", "10    Command : " & CommandLine
    SpreadFlagMessage 0, "+n", "3*** End of error message"
  End If
  If Scripts(ScNum).SecurityViolation = True Then
    RemScript ScNum
  End If
  If Err.Number > 0 Then Err.Clear
End Sub

'Loads a script and calls 'Sub Init'
Public Function LoadScript(FileName As String) As Boolean ' : AddStack "Scripting_LoadScript(" & FileName & ")"
  Dim Commands As String, FileNum As Integer, ScNum As Long
  Dim x As Integer, Y As Integer, Commands2 As String, Commands3 As String
  Dim msg As String
  On Local Error Resume Next
  FileNum = FreeFile
  
  If Dir(FileName) = "" Then SpreadFlagMessage 0, "+n", "5*** LoadScript failed (file not found: " & GetFileName(FileName) & ")": Close #FileNum: LoadScript = False: Exit Function
  Open FileName For Binary As #FileNum
    If Err.Number <> 0 Then SpreadFlagMessage 0, "+n", "5*** LoadScript failed (file not accessible: " & GetFileName(FileName) & ")": Close #FileNum: LoadScript = False: Exit Function
    If LOF(FileNum) > 1000000 Then SpreadFlagMessage 0, "+n", "5*** LoadScript failed (file too long)": Close #FileNum: LoadScript = False: Exit Function
    Commands = Space(LOF(FileNum))
    Get #FileNum, , Commands
  Close #FileNum
  
  If Right(LCase(FileName), 5) = ".perl" Then
    ScNum = AddScript("PerlScript")
  Else
    ScNum = AddScript("VBScript")
  End If
  
  While InStr(1, LCase(Commands), "!include " & Chr(34)) > 0
    x = InStr(1, LCase(Commands), "!include " & Chr(34))
    Y = InStr(x + 10, LCase(Commands), Chr(34) & "!")
    Commands2 = Mid(Commands, x + 10, Y - x - 10)
    Commands = Left(Commands, x - 1) & Mid(Commands, Y + 2)
    If Commands2 <> "" Then
      FileNum = FreeFile
      Open OneDirBack(FileName) & "\" & Commands2 & ".inc" For Binary As #FileNum
      If Err.Number = 0 Then
        Commands2 = Space(LOF(FileNum))
        Get #FileNum, , Commands2
      Else
        Err.Clear
      End If
      Close #FileNum
    End If
    Commands = Commands & vbCrLf & Commands2
  Wend
  
  If ScNum > 0 Then
    Scripts(ScNum).Name = GetFileName(FileName)
    On Error GoTo Blah
    Scripts(ScNum).Script.AddCode Commands
    If Not IsProcedure(ScNum, "Init") Then
      SpreadFlagMessage 0, "+n", "5*** LoadScript failed (no sub 'Init' found)"
      RemScript ScNum
      LoadScript = False
      Exit Function
    End If
    DontLoadScript = ""
    RunScriptX ScNum, "Init"
    If DontLoadScript <> "" Then
      SpreadFlagMessage 0, "+n", "5*** LoadScript failed (" & DontLoadScript & ")"
      RemScript ScNum
      LoadScript = False
      Exit Function
    End If
  Else
    SpreadFlagMessage 0, "+n", "5*** LoadScript failed (scripts are disabled)"
    LoadScript = False
    Exit Function
  End If
  LoadScript = True

Exit Function
Blah:
  Dim ErrNumber As Long, ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  Err.Clear
  If Scripts(ScNum).Script.error.Number = 0 Then
    SpreadFlagMessage 0, "+n", "3*** Scripting error in '" & Scripts(ScNum).Name & "':"
    SpreadFlagMessage 0, "+n", "10    Error   : " & ErrNumber & " (" & ErrDescription & ")"
    SpreadFlagMessage 0, "+n", "3*** End of error message"
  Else
    SpreadFlagMessage 0, "+n", "3*** Scripting error in '" & Scripts(ScNum).Name & "':"
    SpreadFlagMessage 0, "+n", "10    Error   : " & Scripts(ScNum).Script.error.Number & " (" & Scripts(ScNum).Script.error.Description & ")"
    SpreadFlagMessage 0, "+n", "10    Position: Line " & Scripts(ScNum).Script.error.Line & ", Column " & Scripts(ScNum).Script.error.Column
    SpreadFlagMessage 0, "+n", "10    Context : """ & Scripts(ScNum).Script.error.Text & """"
    SpreadFlagMessage 0, "+n", "3*** End of error message"
  End If
  RemScript ScNum
  LoadScript = False
End Function

Public Sub RenScript(ScNum As Long, NewFile As String) ' : AddStack "Scripting_RenScript(" & scNum & ", " & NewFile & ")"
Dim Rest As String, FileName As String, NewLine As String, u As Long
  Rest = GetPPString("Scripts", "Load", "", AnGeL_INI)
  For u = 1 To ParamCount(Rest)
    FileName = Param(Rest, u)
    If LCase(FileName) = LCase(Scripts(ScNum).Name) Then
      If NewLine <> "" Then NewLine = NewLine & " " & NewFile Else NewLine = NewFile
    Else
      If NewLine <> "" Then NewLine = NewLine & " " & FileName Else NewLine = FileName
    End If
  Next u
  WritePPString "Scripts", "Load", NewLine, AnGeL_INI
  'Keep track of scripted commands
  For u = 1 To CommandCount
    If Commands(u).Script = Scripts(ScNum).Name Then Commands(u).Script = NewFile
  Next u
  Scripts(ScNum).Name = NewFile
End Sub

Public Function IsProcedure(ScNum As Long, ProcName As String) As Boolean ' : AddStack "Scripting_IsProcedure(" & scNum & ", " & ProcName & ")"
  Dim u As Long
  If Scripts(ScNum).Script.Language = "PerlScript" Then IsProcedure = True: Exit Function
  For u = 1 To Scripts(ScNum).Script.Modules(1).Procedures.Count
    If LCase(Scripts(ScNum).Script.Modules(1).Procedures(u).Name) = LCase(ProcName) Then IsProcedure = True: Exit Function
  Next u
  IsProcedure = False
End Function

Public Function AddScript(Optional Language As String = "VBScript") ' : AddStack "Scripting_AddScript()"
  ScriptCount = ScriptCount + 1: If ScriptCount > UBound(Scripts()) Then ReDim Preserve Scripts(UBound(Scripts()) + 5)
  On Local Error Resume Next
  Set Scripts(ScriptCount).Script = CreateObject("ScriptControl")
  If Err.Number <> 0 Then
    Status "*** Couldn't load MSSCRIPT.OCX - all scripts were disabled." & vbCrLf
    SpreadFlagMessage 0, "+m", "5*** I couldn't load MSSCRIPT.OCX! Get the scripting engine here:"
    SpreadFlagMessage 0, "+m", "2    http://www.angel-bot.de/files/msscript.ocx"
    SpreadFlagMessage 0, "+m", "5    After finishing your download, DCC send this file to the bot"
    SpreadFlagMessage 0, "+m", "5    and type '.iscripts' on the party line. You can then try to"
    SpreadFlagMessage 0, "+m", "5    add this script again."
    DisableScripts = True
    AddScript = 0: ScriptCount = ScriptCount - 1
    Err.Clear
    Exit Function
  End If
  Scripts(ScriptCount).Script.Language = Language
  If Err.Number > 0 Then
    SpreadFlagMessage 0, "+m", "5*** Cant use " & Language & " - Be sure everything installed."
    AddScript = 0: ScriptCount = ScriptCount - 1
    Err.Clear
    Exit Function
  End If
  Scripts(ScriptCount).Script.AddObject "BotScriptCommands", ScriptCMDs, True
  Scripts(ScriptCount).Script.AllowUI = IIf(Language = "PerlScript", True, False)
  Scripts(ScriptCount).Script.UseSafeSubset = IIf(Language = "PerlScript", False, AnGeLFiles.CommandAllowed("Objects"))
  Scripts(ScriptCount).Description = ""
  Scripts(ScriptCount).SecurityViolation = False
  Scripts(ScriptCount).Silent = False
  Scripts(ScriptCount).Hooks.Chan_msg = False
  Scripts(ScriptCount).Hooks.Priv_msg = False
  Scripts(ScriptCount).Hooks.Chan_act = False
  Scripts(ScriptCount).Hooks.Priv_act = False
  Scripts(ScriptCount).Hooks.Priv_ctcp = False
  Scripts(ScriptCount).Hooks.Priv_ctcpreply = False
  Scripts(ScriptCount).Hooks.Chan_ctcp = False
  Scripts(ScriptCount).Hooks.Chan_notice = False
  Scripts(ScriptCount).Hooks.Priv_notice = False
  Scripts(ScriptCount).Hooks.Server_notice = False
  Scripts(ScriptCount).Hooks.Join = False
  Scripts(ScriptCount).Hooks.Part = False
  Scripts(ScriptCount).Hooks.Quit = False
  Scripts(ScriptCount).Hooks.Nick = False
  Scripts(ScriptCount).Hooks.Kick = False
  Scripts(ScriptCount).Hooks.Commands = False
  Scripts(ScriptCount).Hooks.Op = False
  Scripts(ScriptCount).Hooks.Deop = False
  Scripts(ScriptCount).Hooks.Mode = False
  Scripts(ScriptCount).Hooks.ModeEnd = False
  Scripts(ScriptCount).Hooks.Topic = False
  Scripts(ScriptCount).Hooks.Raw = False
  Scripts(ScriptCount).Hooks.RawFilter = ""
  Scripts(ScriptCount).Hooks.Numerics = False
  Scripts(ScriptCount).Hooks.Botnet = False
  Scripts(ScriptCount).Hooks.BN_Msg = False
  Scripts(ScriptCount).Hooks.Whois = False
  Scripts(ScriptCount).Hooks.Resolves = False
  Scripts(ScriptCount).Hooks.Party_out = False
  Scripts(ScriptCount).Hooks.AddedUser = False
  Scripts(ScriptCount).Hooks.RemovedUser = False
  Scripts(ScriptCount).Hooks.AddedHost = False
  Scripts(ScriptCount).Hooks.RemovedHost = False
  Scripts(ScriptCount).Hooks.ChangedNick = False
  Scripts(ScriptCount).Hooks.PLJoin = False
  Scripts(ScriptCount).Hooks.Unload = False
  Scripts(ScriptCount).Hooks.Srv_Connect = False
  Scripts(ScriptCount).Hooks.Ban = False
  Scripts(ScriptCount).Hooks.UnBan = False
  Scripts(ScriptCount).Hooks.fa_uploadbegin = False
  Scripts(ScriptCount).Hooks.fa_uploadcomplete = False
  Scripts(ScriptCount).Hooks.fa_downloadbegin = False
  Scripts(ScriptCount).Hooks.fa_downloadcomplete = False
  Scripts(ScriptCount).Hooks.fa_command = False
  Scripts(ScriptCount).Hooks.seen = False
  Scripts(ScriptCount).Hooks.fa_userjoin = False
  Scripts(ScriptCount).Hooks.fa_userleft = False
  Scripts(ScriptCount).Hooks.KI = False
  
  AddScript = ScriptCount
End Function

Public Sub RemScript(ScNum As Long) ' : AddStack "Scripting_RemScript(" & scNum & ")"
Dim u As Long, RemovedOne As Boolean, CommandName As String
  If ScNum = 0 Or ScNum > ScriptCount Then Exit Sub
  'Check unload hook
  If Scripts(ScNum).Hooks.Unload Then
    RunScriptX ScNum, "Unload"
  End If
  'Remove commands added by the script
  Do
    RemovedOne = False
    For u = 1 To CommandCount
      CommandName = Commands(u).Name
      If Commands(u).Script = Scripts(ScNum).Name Then RemCommand CommandName: RemovedOne = True: Exit For
    Next u
  Loop While RemovedOne
  'Remove timers started by the script


  For u = 1 To EventCount
    If Param(Events(u).DoThis, 1) = "CallScript" Then
      If Param(Events(u).DoThis, 2) = Scripts(ScNum).Name Then Events(u).DoThis = ""
    End If
  Next u
  'Remove sockets opened by the script
  For u = 1 To SocketCount
    If IsValidSocket(u) Then
      If GetSockFlag(u, SF_Status) = SF_Status_ScriptSocket Then
        If Scripts(ScNum).Name = SocketItem(u).SetupChan Then
          DisconnectSocket u
        End If
      End If
    End If
  Next u
  Set Scripts(ScNum).Script = Nothing
  For u = ScNum To ScriptCount - 1
    Scripts(u) = Scripts(u + 1)
  Next u
  ScriptCount = ScriptCount - 1
End Sub

'Converts " to "" to make a string script-friendly
Function MSS(Line As String) As String
Dim u As Long, a As Boolean, b As Boolean, c As Boolean, d As Boolean, E As Boolean
Dim NewLine As String, Char As String
  a = (InStr(Line, """") > 0)
  b = (InStr(Line, vbCrLf) > 0)
  c = (InStr(Line, vbCr) > 0)
  d = (InStr(Line, vbLf) > 0)
  E = (InStr(Line, Chr(0)) > 0)
  If (a = False) And (b = False) And (c = False) And (d = False) And (E = False) Then MSS = Line: Exit Function
  If a = True Then
    For u = 1 To Len(Line)
      Char = Mid(Line, u, 1)
      If Char = """" Then Char = """"""
      NewLine = NewLine + Char
    Next u
  Else
    NewLine = Line
  End If
  If b = True Then
    NewLine = Replace(NewLine, vbCrLf, """ & vbCrLf & """)
  End If
  If c = True Then
    NewLine = Replace(NewLine, vbCr, """ & vbCr & """)
  End If
  If d = True Then
    NewLine = Replace(NewLine, vbLf, """ & vbLf & """)
  End If
  If E = True Then
    NewLine = Replace(NewLine, Chr(0), """ & Chr(0) & """)
  End If
  MSS = NewLine
End Function

