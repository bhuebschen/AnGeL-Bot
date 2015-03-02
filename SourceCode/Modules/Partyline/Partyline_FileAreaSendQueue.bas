Attribute VB_Name = "Partyline_FileAreaSendQueue"
',-======================- ==-- -  -
'|   AnGeL - Partyline - FileAreaSendQueue
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Type FileSendQueueFile
  RegNick As String
  IRCNick As String
  FileName As String
End Type


Public FileSendQueue() As FileSendQueueFile
Public QueuedFiles As Long


Sub FileAreaSendQueue_Load()
  ReDim Preserve FileSendQueue(5)
End Sub


Sub FileAreaSendQueue_Unload()
'
End Sub

Sub AfterFileCompletion(ByVal TheSock As Long) ' : AddStack "Routines_AfterFileCompletion(" & TheSock & ")"
  If LCase(SocketItem(TheSock).FileName) = "angel.exe" Then
    If InStr(GetFileDescription(FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName), "--- AnGeL IRC Bot ---") > 0 Then
      FileCopy FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName, HomeDir & "Update.exe"
      SpreadFlagMessage 0, "+m", "14[" & Time & "] 3*** Received AutoUpdate file - type '.update' to launch it!"
    Else
      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** ERROR: Received invalid AutoUpdate file!"
    End If
    Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
  ElseIf LCase(SocketItem(TheSock).FileName) = "motd.txt" Then
    If SocketItem(TheSock).FileSize > 0 Then
      FileCopy FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName, HomeDir & "Motd.txt"
      Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Updated MOTD."
    Else
      Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
      Kill HomeDir & "Motd.txt"
      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Deleted MOTD."
    End If
  ElseIf LCase(Right(SocketItem(TheSock).FileName, 4)) = ".asc" Then
    SpreadFlagMessage 0, "+s", "3*** You can load this script now: .+script " & SocketItem(TheSock).FileName
  ElseIf LCase(Right(SocketItem(TheSock).FileName, 5)) = ".seen" Then
    If SocketItem(TheSock).FileSize > 0 Then
      DeletePPString "\Incoming", SocketItem(TheSock).FileName, HomeDir & "Files.ini"
      WritePPString "\", SocketItem(TheSock).FileName, SocketItem(TheSock).RegNick, HomeDir & "Files.ini"
      FileCopy FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName, FileAreaHome & "" & SocketItem(TheSock).FileName
      Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Added seen list '" & LCase(Left(SocketItem(TheSock).FileName, Len(SocketItem(TheSock).FileName) - 5)) & "'."
    Else
      DeletePPString "\Incoming", SocketItem(TheSock).FileName, HomeDir & "Files.ini"
      DeletePPString "\", SocketItem(TheSock).FileName, HomeDir & "Files.ini"
      Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
      Kill FileAreaHome & "" & SocketItem(TheSock).FileName
      SpreadFlagMessage 0, "+m", "14[" & Time & "] *** Deleted seen list."
    End If
  ElseIf LCase(Right(SocketItem(TheSock).FileName, 4)) = ".lng" Then
    DeletePPString "\Incoming", SocketItem(TheSock).FileName, HomeDir & "Files.ini"
    WritePPString "\", SocketItem(TheSock).FileName, SocketItem(TheSock).RegNick, HomeDir & "Files.ini"
    FileCopy FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName, FileAreaHome & "" & SocketItem(TheSock).FileName
    Kill FileAreaHome & "Incoming\" & SocketItem(TheSock).FileName
    SpreadFlagMessage 0, "+m", "14[" & Time & "] *** The language '" & LCase(Left(SocketItem(TheSock).FileName, Len(SocketItem(TheSock).FileName) - 4)) & "' is now available in '.botsetup'."
  End If
End Sub

