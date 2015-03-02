Attribute VB_Name = "OpSys_ThreadPriority"
',-======================- ==-- -  -
'|   AnGeL - OpSys - ThreadPriority
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Sub ChangePriority()
  Dim hProcess As Long
  Dim ret As Long, pid As Long
  pid = kernel32_GetCurrentProcessId()
  hProcess = kernel32_OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION, False, pid)
  If hProcess = 0 Then Exit Sub
  ret = kernel32_SetPriorityClass(hProcess, BELOW_NORMAL_PRIORITY_CLASS)
  Call CloseHandle(hProcess)
End Sub
