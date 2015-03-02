Attribute VB_Name = "Server_AuthIdentify"
',-======================- ==-- -  -
'|   AnGeL - Server - AuthIdentify
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


' -= Service Name, Kommando und Parameter
Public AuthTarget As String
Public AuthCommand As String
Public AuthParam1 As String
Public AuthParam2 As String
Public AuthReAuth As Boolean
Public AuthJust As Boolean


Sub CommandAuth(vSocket As Long, Line As String)
  If AuthTarget <> "" And AuthCommand <> "" And AuthParam1 <> "" And AuthJust = False Then
    SendLine "PRIVMSG " & AuthTarget & " :" & AuthCommand & " " & AuthParam1 & IIf(AuthParam2 <> "", " " & AuthParam2, ""), 1
    SpreadFlagMessage 0, "+s", "14*** Trying to AUTH."
    AuthJust = True
  ElseIf AuthJust = True Then
    TU vSocket, "5*** I tried to AUTH just moments ago."
  Else
    TU vSocket, "5*** AUTH setting incomplete. Try .authsetup first."
  End If
End Sub

