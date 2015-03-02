Attribute VB_Name = "OpSys_Main"
',-======================- ==-- -  -
'|   AnGeL - OpSys - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Sub OpSys_Load()
  GetWinInfo
  ChangePriority
  If Not DebugMode Then kernel32_SetUnhandledExceptionFilter AddressOf ExceptionFilter
End Sub


Sub OpSys_Unload()
  kernel32_SetUnhandledExceptionFilter 0&
End Sub



