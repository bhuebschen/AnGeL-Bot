Attribute VB_Name = "Server_Main"
',-======================- ==-- -  -
'|   AnGeL - Server - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Sub Server_Load()
  SendQueue_Load
  SplitDetection_Load
  Channels_Load
  IgnoreList_Load
  BanList_Load
  Features_Load
End Sub


Sub Server_Unload()
  SendQueue_Unload
  SplitDetection_Unload
  Channels_Unload
  IgnoreList_Unload
  BanList_Unload
  Features_Unload
End Sub
