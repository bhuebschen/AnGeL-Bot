Attribute VB_Name = "Server_SplitDetection"
',-======================- ==-- -  -
'|   AnGeL - Server - SplitDetection
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

    
Private Type SplitServer
  Name As String
  SplittedAt As Date
End Type
    
    
Public SplitServerCount As Long
Public SplitServers() As SplitServer
    
    
Sub SplitDetection_Load()
  ReDim Preserve SplitServers(5)
  ReDim Preserve SplitServers(5)
End Sub

Sub SplitDetection_Unload()
'
End Sub


Public Sub AddSplitServer(ServerName As String)
  Dim u As Long
  For u = 1 To SplitServerCount
    If (LCase(SplitServers(u).Name) Like LCase(ServerName)) Then Exit Sub
  Next u
  SplitServerCount = SplitServerCount + 1: If SplitServerCount > UBound(SplitServers()) Then ReDim Preserve SplitServers(UBound(SplitServers()) + 5)
  SplitServers(SplitServerCount).Name = ServerName
  SplitServers(SplitServerCount).SplittedAt = Now
End Sub


Public Sub RemoveSplitServer(ServerName As String)
  Dim u As Long, s As Long
  For u = 1 To SplitServerCount
    If (LCase(SplitServers(u).Name) Like LCase(ServerName)) Or (LCase(ServerName) Like LCase(SplitServers(u).Name)) Then
      For s = u To SplitServerCount - 1
        SplitServers(s) = SplitServers(s + 1)
      Next s
      SplitServerCount = SplitServerCount - 1
      Exit For
    End If
  Next u
  ReDim Preserve SplitServers(((SplitServerCount \ 5) + 1) * 5)
End Sub

