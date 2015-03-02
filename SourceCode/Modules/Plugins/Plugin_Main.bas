Attribute VB_Name = "Plugin_Main"
',-======================- ==-- -  -
'|   AnGeL - Plugins - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Sub Plugins_Load()
  KI_Load
  Uptime_Load
  Whatis_Load
  Notes_Load
  Seen_Load
End Sub

Sub Plugins_Unload()
  KI_Unload
  Uptime_Unload
  Whatis_Unload
  Notes_Unload
  Seen_Unload
End Sub
