Attribute VB_Name = "TimedEvents_Main"
',-======================- ==-- -  -
'|   AnGeL - TimedEvents - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Sub TimedEvents_Load()
  Events_Load
  Orders_Load
End Sub

Sub TimedEvents_Unload()
  Events_Unload
  Orders_Unload
End Sub
