Attribute VB_Name = "Sockets_SocketList"
',-======================- ==-- -  -
'|   AnGeL - Sockets - Array Steuerung für Sockets
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


',-======================- ==-- -  -
'|   Konstanten
'`-=====================-- -===- ==- -- -
Public Const SD_In As Byte = 0
Public Const SD_Out As Byte = 1
Public Const SS_Idle As Byte = 0
Public Const SS_Connecting As Byte = 1
Public Const SS_Disconnecting As Byte = 2
Public Const SS_Connected As Byte = 3
Public Const SS_Listening As Byte = 4
Public Const SocketStepping As Long = 5


',-======================- ==-- -  -
'|   Variablen
'`-=====================-- -===- ==- -- -
Public SocketItem() As Sockets_SocketItem
Public SocketCount As Long
Public SocketUse As Long


Private Function SocketList_Add() As Long
  ',-= Zahl erhöhen
  SocketCount = SocketCount + 1
  SocketUse = SocketUse + 1
  
  ',-= Array prüfen und ggf. anpassen
  If SocketCount > UBound(SocketItem) Then
    ReDim Preserve SocketItem(((SocketCount / SocketStepping) + 1) * SocketStepping)
  End If
  
  ',-= Objekt erstellen und Index übergeben
  Set SocketItem(SocketCount) = New Sockets_SocketItem
  
  ',-= Index zurückgeben
  SocketList_Add = SocketCount
End Function

Private Function SocketList_Del(Index As Long) As Long
  Dim Counter As Long
  
  If SocketItem(Index) Is Nothing Then Exit Function
  
  ',-= Zahl vermindern
  SocketUse = SocketUse - 1
  
  ',-= Objekt löschen
  Set SocketItem(Index) = Nothing
  
  ',-= Array prüfen und ggf. anpassen
  If Index = SocketCount Then
    For Counter = SocketCount To 1 Step -1
      If Not IsValidSocket(Counter) Then
        SocketCount = SocketCount - 1
      Else
        Exit For
      End If
    Next Counter
    If UBound(SocketItem) - SocketCount >= SocketStepping Then
      ReDim Preserve SocketItem(((SocketCount / SocketStepping) + 1) * SocketStepping)
    End If
  End If
End Function

Public Function AddSocket() As Long
  Dim Index As Long
  
  ',-= Zähler vergleichen
  If SocketUse = SocketCount Then
    ',-= Neues Objekt erstellen und Wert zurückgeben
    AddSocket = SocketList_Add
  Else
    ',-= Nach altem Objekt suchen
    For Index = 1 To SocketCount
      If Not IsValidSocket(Index) Then
        ',-= Neues Objekt erstellen und Wert zurückgeben
        Set SocketItem(Index) = New Sockets_SocketItem
        SocketUse = SocketUse + 1
        AddSocket = Index
        Exit For
      End If
    Next Index
  End If
End Function

Public Sub DisconnectSocket(ByVal Index As Long)
  If SocketItem(Index).SocketStatus = SS_Disconnecting Or SocketItem(Index).SocketStatus = SS_Listening Then
    RemoveSocket Index, -1, "", True
  Else
    SocketItem(Index).SocketStatus = SS_Disconnecting
    winsock2_shutdown SocketItem(Index).SocketNumber, 1
  End If
End Sub

Public Sub RemoveSocket(ByVal Index As Long, ByVal ErrorCode As Long, ByVal Reason As String, ByVal SurpressClosed As Boolean)
  If SocketItem(Index).SocketNumber <> -1 Then
    winsock2_shutdown SocketItem(Index).SocketNumber, 1
    winsock2_closesocket SocketItem(Index).SocketNumber
  End If
  If ((GetSockFlag(Index, SF_Status) = SF_Status_Bot) Or (GetSockFlag(Index, SF_Status) = SF_Status_BotLinking)) Then RemBot SocketItem(Index).RegNick, Index, Reason
  If Not SurpressClosed Then Socket_Closed Index, ErrorCode, Reason
  SocketList_Del Index
End Sub

Public Function IsValidSocket(Index As Long)
  If Index > 0 Then
    If Index <= SocketCount Then
      If Not SocketItem(Index) Is Nothing Then
        IsValidSocket = True
        Exit Function
      End If
    End If
  End If
  IsValidSocket = False
End Function

Public Function GetSockFlag(TheSock As Long, SF As Byte)
  If IsValidSocket(TheSock) Then
    GetSockFlag = SocketItem(TheSock).SocketFlag(SF)
  End If
End Function

Public Sub SetSockFlag(TheSock As Long, SF As Byte, Value As String)
  If IsValidSocket(TheSock) Then
    SocketItem(TheSock).SocketFlag(SF) = Value
  End If
End Sub

Public Sub TU(vsock As Long, ByVal What As String)
  Dim StrToSend As String, DidntSend As Boolean, ScNum As Long
  If Not IsValidSocket(vsock) Then Exit Sub
  If GetSockFlag(vsock, SF_Colors) = SF_YES Then
    StrToSend = What & IIf(GetSockFlag(vsock, SF_LF_ONLY) = SF_YES, vbCrLf, vbCrLf)
    If GetSockFlag(vsock, SF_Telnet) = SF_YES Then
      StrToSend = MakeTelnetColor(StrToSend, vsock)
    End If
  Else
    StrToSend = Strip(What) & IIf(GetSockFlag(vsock, SF_LF_ONLY) = SF_YES, vbCrLf, vbCrLf)
  End If
  If CalledByScript = False Then
    HaltDefault = False
    For ScNum = 1 To ScriptCount
      If Scripts(ScNum).Hooks.Party_out Then
        RunScriptX ScNum, "Party_out", vsock, What
      End If
    Next ScNum
    If HaltDefault Then Exit Sub
  End If
  If SocketItem(vsock).IRCNick = "²*SCRIPT*²" Then
    If SocketItem(vsock).CurrentQuestion <> "" Then
      For ScNum = 1 To ScriptCount
        If Scripts(ScNum).Name = SocketItem(vsock).SetupChan Then
          RunScriptX ScNum, SocketItem(vsock).CurrentQuestion, What
        End If
      Next ScNum
    End If
    Exit Sub
  End If
  DidntSend = False
  If SocketItem(vsock).SendQLines > 0 Then
    DidntSend = True
  Else
    If SendTCP(vsock, StrToSend) = -1 Then DidntSend = True
  End If
  'Couldn't winsock2_send packet - store in SendQ
  If DidntSend Then
    If SocketItem(vsock).SendQLines > 200 Then Exit Sub 'Too many lines in SendQ
    Output "Socket [" & CStr(vsock) & "]: Stored one line in SendQ." & vbCrLf
    SocketItem(vsock).SendQLines = SocketItem(vsock).SendQLines + 1
    SocketItem(vsock).SendQ(SocketItem(vsock).SendQLines) = StrToSend
    If Not IsTimed("PutSendQ " & CStr(vsock)) Then
      TimedEvent "PutSendQ " & CStr(vsock), 2
    End If
  End If
End Sub

Public Sub TUEx(vsock As Long, SockFlag As Byte, ByVal What As String)
  If GetSockFlag(vsock, SockFlag) = SF_YES Then TU vsock, What
End Sub

Public Sub RTU(vsock As Long, ByVal What As String)
  Dim StrToSend As String, DidntSend As Boolean
  If Not IsValidSocket(vsock) Then Exit Sub
  StrToSend = What & IIf(GetSockFlag(vsock, SF_LF_ONLY) = SF_YES, vbLf, vbCrLf)
  DidntSend = False
  If SocketItem(vsock).SendQLines > 0 Then
    DidntSend = True
  Else
    If SendTCP(vsock, StrToSend) = -1 Then DidntSend = True
  End If
  'Couldn't winsock2_send packet - store in SendQ
  If DidntSend Then
    Output "Socket [" & CStr(vsock) & "]: Stored one line in SendQ." & vbCrLf
    SocketItem(vsock).SendQLines = SocketItem(vsock).SendQLines + 1
    SocketItem(vsock).SendQ(SocketItem(vsock).SendQLines) = StrToSend
    If Not IsTimed("PutSendQ " & CStr(vsock)) Then
      TimedEvent "PutSendQ " & CStr(vsock), 2
    End If
  End If
End Sub

Public Sub SS(Sock As Long, ByVal What As String)
  SendTCP Sock, What & vbCrLf
End Sub

Public Function FindFreeOrderSign() As String
  Dim u As Long, AlreadyUsed As String
  AlreadyUsed = ","
  For u = 1 To SocketCount
    If IsValidSocket(u) Then If SocketItem(u).OrderSign <> "" Then AlreadyUsed = AlreadyUsed & SocketItem(u).OrderSign & ","
  Next u
  u = 11
  While InStr(AlreadyUsed, "," & MakeOrderSign(u) & ",") <> 0
    u = u + 1
  Wend
  FindFreeOrderSign = MakeOrderSign(u)
End Function

Public Function MakeOrderSign(vsock As Long) As String
  Dim NewSock As Long, prefix As Long
  NewSock = vsock
  While NewSock > 52
    NewSock = NewSock - 52
    prefix = prefix + 1
  Wend
  If NewSock < 27 Then
    If prefix = 0 Then
      MakeOrderSign = Chr(Asc("A") + NewSock - 1)
    Else
      MakeOrderSign = CStr(prefix) & Chr(Asc("A") + NewSock - 1)
    End If
  Else
    If prefix = 0 Then
      MakeOrderSign = Chr(Asc("a") + NewSock - 27)
    Else
      MakeOrderSign = CStr(prefix) & Chr(Asc("a") + NewSock - 27)
    End If
  End If
End Function

