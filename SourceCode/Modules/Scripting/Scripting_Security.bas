Attribute VB_Name = "Scripting_Security"
',-======================- ==-- -  -
'|   AnGeL - Scripting - Security
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Sub PolError(ScNum As Long, ErrorX As String)
  SpreadFlagMessage 0, "+n", "3*** Security-Error while executing script '" & Scripts(ScNum).Name & "':"
  SpreadFlagMessage 0, "+n", "10    Error   : " & Err.Number & " (" & Err.Description & ")"
  SpreadFlagMessage 0, "+n", "10    Command : " & ErrorX
  SpreadFlagMessage 0, "+n", "10    "
  SpreadFlagMessage 0, "+n", "4    Unloading script!"
  Scripts(ScNum).SecurityViolation = True
  SpreadFlagMessage 0, "+n", "3*** End of error message"
End Sub

Function CheckForAllow(Section, Entry, Value) As Boolean
  If (InStr(LCase(Section), "allowcreate")) Then
    CheckForAllow = False
  ElseIf (InStr(LCase(Entry), "allowcreate")) Then
    CheckForAllow = False
  ElseIf (InStr(LCase(Value), "allowcreate")) Then
    CheckForAllow = False
  Else
    CheckForAllow = True
  End If
End Function

Function CheckForAllow2(Section, Entry, Value) As Boolean
  If (InStr(LCase(Section), "allowrun")) Then
    CheckForAllow2 = False
  ElseIf (InStr(LCase(Entry), "allowrun")) Then
    CheckForAllow2 = False
  ElseIf (InStr(LCase(Value), "allowrun")) Then
    CheckForAllow2 = False
  Else
    CheckForAllow2 = True
  End If
End Function

