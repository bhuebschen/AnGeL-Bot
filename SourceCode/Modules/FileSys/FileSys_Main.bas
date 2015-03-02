Attribute VB_Name = "FileSys_Main"
',-======================- ==-- -  -
'|   AnGeL - FileSys - Main
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public HomeDir          As String
Public FileAreaHome     As String
Public AnGeL_INI        As String
Public NET_INI          As String


Sub FileSys_Load()
  If Right(App.Path, 1) = "\" Then HomeDir = App.Path Else HomeDir = App.Path & "\"
  
  FileAreaHome = HomeDir & "FileArea\"
  AnGeL_INI = HomeDir & "AnGeL.ini"
  NET_INI = HomeDir & "net.conf"
  
  NTFS_Load
End Sub


Sub FileSys_Unload()
  NTFS_Unload
End Sub
