Attribute VB_Name = "Server_Features"
',-======================- ==-- -  -
'|   AnGeL - Server - Features
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Type TServerType
  Network As String
  SupportsMultiChanJoin As Boolean
  SupportsMultiChanWho As Boolean
  SupportsMultiChanMode As Boolean
  SupportsServersChan As Boolean
  SupportsMultiKicks As Boolean
  HidesHosts As Boolean
  MaxNickLength As Long
End Type


Public ServerNetwork As String
Public ServerNickLen As Long
Public ServerUseFullAdress As Boolean
Public ServerNumberOfModes As Byte
Public ServerMaxChannels As Byte
Public ServerSplitDetection As Boolean
Public ServerChannelPrefixes As String
Public ServerChannelModes As String
Public ServerTopicLen As Integer


Public ServerInfo As TServerType
Public AutoNetSetup As Boolean


Sub Features_Load()
'
End Sub


Sub Features_Unload()
'
End Sub
