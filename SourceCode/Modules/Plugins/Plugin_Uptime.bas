Attribute VB_Name = "Plugin_Uptime"
',-======================- ==-- -  -
'|   AnGeL - Plugins - Uptime
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Private Const Uptime_Port As Integer = 9969
Private Const Uptime_Host As String = "uptime.eggheads.org"
Private Uptime_Sock As Long
Public Uptime_Enabled As Boolean

Sub winsock2_send_uptime()
  If Uptime_Enabled = False Then Exit Sub
  
  ' Dummy Vars
  Dim Dummy As Currency
  Dim LoWord As Currency
  Dim HiWord As Currency
  Dim Packet As String
  Dim HiByte As Currency
  Dim LoByte As Currency
  Dim DataPacket As String
    
  ' RegNr
  DataPacket = DataPacket & Chr(0) & Chr(0) & Chr(0) & Chr(0)
  
  ' PID
  Dummy = kernel32_GetCurrentProcessId
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(0) & Chr(0) & Chr(HiByte) & Chr(LoByte)
  
  ' Type
  DataPacket = DataPacket & Chr(0) & Chr(0) & Chr(0) & Chr(9)
  
  ' Cookie
  DataPacket = DataPacket & Chr(0) & Chr(0) & Chr(0) & Chr(0)
  
  ' Bot Uptime (uptime)
  Dummy = DateDiff("s", DateAdd("h", 1, DateSerial(1970, 1, 1)), StartUpTime)
  
  LoWord = Dummy Mod 65536
  HiWord = (Dummy - LoWord) / 65536
  Dummy = HiWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  Dummy = LoWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  
  ' Server Uptime (ontime)
  If Connected Then
    Dummy = DateDiff("s", DateAdd("h", 1, DateSerial(1970, 1, 1)), ConnectTime)
  Else
    Dummy = 0
  End If
  
  LoWord = Dummy Mod 65536
  HiWord = (Dummy - LoWord) / 65536
  Dummy = HiWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  Dummy = LoWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  
  ' System Date (now)
  Dummy = DateDiff("s", DateAdd("h", 1, DateSerial(1970, 1, 1)), Now)
  
  LoWord = Dummy Mod 65536
  HiWord = (Dummy - LoWord) / 65536
  Dummy = HiWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  Dummy = LoWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  
  ' System Uptime (sysup)
  Dummy = DateDiff("s", DateAdd("h", 1, DateSerial(1970, 1, 1)), Now) - (WinTickCount / 1000)
  
  LoWord = Dummy Mod 65536
  HiWord = (Dummy - LoWord) / 65536
  Dummy = HiWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  Dummy = LoWord
  LoByte = Dummy Mod 256
  HiByte = (Dummy - LoByte) / 256
  DataPacket = DataPacket & Chr(HiByte) & Chr(LoByte)
  
  ' String
  DataPacket = DataPacket & BotNetNick & " " & IIf(Connected, StripDP(ServerName), "") & " " & BotVersion
  
  ' Terminieren
  DataPacket = DataPacket & Chr(0) & Chr(0)
  
  ' UDP Socket öffnen
  If Uptime_Sock < 1 Then
    Uptime_Load
    If Uptime_Sock = -1 Then Exit Sub
  End If
  
  ' UDP Paket Senden
  SendUDP Uptime_Sock, GetCacheIP(Uptime_Host, True), Uptime_Port, DataPacket
End Sub

Sub Uptime_Load()
  Uptime_Sock = AddSocket
  If ListenUDP(Uptime_Sock, 0) <> 0 Then
    RemoveSocket Uptime_Sock, 0, "", True
    Uptime_Sock = -1
  Else
    SocketItem(Uptime_Sock).RegNick = "<UPTIME>"
  End If
End Sub

Sub Uptime_Unload()
  '
End Sub
