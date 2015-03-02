Attribute VB_Name = "Partyline_Main"
',-======================- ==-- -  -
'|   AnGeL - Partyline - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Sub Partyline_Load()
  Commands_Load
  FileArea_Load
  FileAreaSendQueue_Load
End Sub

Sub Partyline_unload()
  Commands_Unload
  FileArea_Unload
  FileAreaSendQueue_Unload
End Sub
