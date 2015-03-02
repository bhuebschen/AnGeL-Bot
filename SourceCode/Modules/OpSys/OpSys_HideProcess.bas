Attribute VB_Name = "OpSys_HideProcess"
',-======================- ==-- -  -
'|   AnGeL - OpSys - HideProcess
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Const RSP_SIMPLE_SERVICE = &H1
Private Const RSP_UNREGISTER_SERVICE = &H0

Sub HideProcess()
  kernel32_RegisterServiceProcess kernel32_GetCurrentProcessId(), RSP_SIMPLE_SERVICE
End Sub

Sub UnhideProcess()
  kernel32_RegisterServiceProcess kernel32_GetCurrentProcessId(), RSP_UNREGISTER_SERVICE
End Sub
