Attribute VB_Name = "Sockets_Winsock2"
',-======================- ==-- -  -
'|   AnGeL - Sockets - Winsock2
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


',-======================- ==-- -  -
'|   Variablen
'`-=====================-- -===- ==- -- -
Public AddressIPv6 As String
Public AddressIPv4 As String
Public AddressDefault As String
Public WSAWindowHandle As Long
Public WSAData As LPWSADATA
Public IPv6Caps As Byte
Public WSAOldWindowProc As Long
Public PortRange As String
Public LocalAddress As String
Public WM_Sockets As Long


',-= IsListenPortInUse - Gibt aus ob eine IP lokal ist
Public Function IsValidLocalIP(IP As String) As Boolean
  Dim TestSock As Long
  TestSock = AddSocket
  If ListenTCP(TestSock, 135, IP) = 10048 Then
    IsValidLocalIP = True
  Else
    IsValidLocalIP = False
  End If
  RemoveSocket TestSock, 0, "", True
End Function

',-= IsListenPortInUse - Gibt aus ob ein Port belegt ist
Public Function IsListenPortInUse(Port As Long) As Boolean
  Dim TestSock As Long
  TestSock = AddSocket
  If ListenTCP(TestSock, Port) = 0 Then
    IsListenPortInUse = False
  Else
    IsListenPortInUse = True
  End If
  RemoveSocket TestSock, 0, "", True
End Function

',-= IsValidIP - Gibt aus ob ein Host eine IP ist
Public Function IsValidIP(Host As String) As Boolean
  If IsValidIPv4(Host) = True Then
    IsValidIP = True
    Exit Function
  ElseIf IsValidIPv6(Host) = True Then
    IsValidIP = True
    Exit Function
  Else
    IsValidIP = False
  End If
End Function

',-= IsValidIPv4 - Gibt aus ob ein Host eine IPv4 ist
Public Function IsValidIPv4(Host As String) As Boolean
  If winsock2_inet_addr(Host) = -1 Then
    IsValidIPv4 = False
  Else
    IsValidIPv4 = True
  End If
End Function

',-= IsValidIPv6 - Gibt aus ob ein Host eine IPv6 ist
Public Function IsValidIPv6(Host As String) As Boolean
  '+'
  IsValidIPv6 = (InStr(1, Host, ":", vbBinaryCompare) <> 0)
End Function

',-= WSAGetLastError - Gibt einen FehlerCode zurück
Public Function WSAGetLastError() As Long
  If Err.LastDllError = 0 Then
    WSAGetLastError = -1
  Else
    WSAGetLastError = Err.LastDllError
  End If
End Function

',-= WSAASyncGetHost - Startet eine (asynchrone) DNS Abfrage
Public Function WSAASyncGetHost(ByVal Host As String) As Long
  Dim retip As Long
  Dim RetCode As Long
  
  ',-= Puffer zurücksetzen
  HostEntBuffer = HostEntZero
  
  ',-= IP Version prüfen
  If IsValidIPv4(Host) = True Then
    ',-= IPv4 auflösen
    retip = winsock2_inet_addr(Host)
    RetCode = winsock2_WSAAsyncGetHostByAddr(WSAWindowHandle, WM_Resolve, retip, 4, AF_INET, HostEntBuffer, hostentasync_size)
    If RetCode = -1 Then
      WSAASyncGetHost = WSAGetLastError
    Else
      WSAASyncGetHost = RetCode
    End If
  ElseIf IsValidIPv6(Host) = True Then
    '2do',-= IPv6 auflösen
    WSAASyncGetHost = 10045
  Else
    ',-= Hostnamen auflösen
    RetCode = winsock2_WSAAsyncGetHostByName(WSAWindowHandle, WM_Resolve, Host, HostEntBuffer, hostentasync_size)
    If RetCode = -1 Then
      WSAASyncGetHost = WSAGetLastError
    Else
      WSAASyncGetHost = RetCode
    End If
  End If
End Function

',-= GetPortFromSocketAddress - Gibt den Port aus einer SocketAddress zurück
Public Function GetPortFromSocketAddress(ByVal SocketAddress As String) As Long
  Dim Index As Long
  Dim Index2 As Long
  
  ',-= Länge prüfen
  If Len(SocketAddress) < 4 Then
    GetPortFromSocketAddress = -1
    Exit Function
  End If
  
  ',-= Wert auslesen
  Index = Asc(Mid(SocketAddress, 4, 1))
  Index2 = Asc(Mid(SocketAddress, 3, 1))

  ',-= Wert ausgeben
  GetPortFromSocketAddress = (Index2 * 256) + Index
End Function

',-= SetPortInSocketAddress - Setzt den Port in einer SocketAddress
Public Sub SetPortInSocketAddress(ByRef SocketAddress As String, ByVal Port As Long)
  Dim Index As Long
  Dim Index2 As Long
  
  ',-= Länge prüfen
  If Len(SocketAddress) < 4 Then
    Err.Raise 10014, , WSAGetErrorString(10014)
    Exit Sub
  End If
  
  ',-= Wert prüfen
  Port = Port Mod 65536
  If Port < 0 Then
    Err.Raise 10014, , WSAGetErrorString(10014)
    Exit Sub
  End If
  
  ',-= Wert schreiben
  Index = Port Mod 256
  Index2 = (Port - Index) / 256
  Mid(SocketAddress, 3, 1) = Chr(Index2)
  Mid(SocketAddress, 4, 1) = Chr(Index)
End Sub

',-= FindSocket - Sucht eine SocketNummer in den SocketItems
Public Function FindSocket(SocketNumber As Long) As Long
  Dim Index As Long
  FindSocket = -1
  For Index = 1 To SocketCount
    If IsValidSocket(Index) Then
      If SocketItem(Index).SocketNumber = SocketNumber Then
        FindSocket = Index
        Exit For
      End If
    End If
  Next Index
End Function

',-= WSAGetErrorString - Gibt die Beschreibung zu einem Winsock Fehler aus
Public Function WSAGetErrorString(ByVal errnum As Long) As String
  Select Case errnum
    Case 10004
      WSAGetErrorString = "Interrupted system call."
    Case 10009
      WSAGetErrorString = "Bad file number."
    Case 10013
      WSAGetErrorString = "Permission Denied."
    Case 10014
      WSAGetErrorString = "Bad Address."
    Case 10022
      WSAGetErrorString = "Invalid Argument."
    Case 10024
      WSAGetErrorString = "Too many open files."
    Case 10035
      WSAGetErrorString = "Operation would block."
    Case 10036
      WSAGetErrorString = "Operation now in progress."
    Case 10037
      WSAGetErrorString = "Operation already in progress."
    Case 10038
      WSAGetErrorString = "Socket operation on nonsocket."
    Case 10039
      WSAGetErrorString = "Destination address required."
    Case 10040
      WSAGetErrorString = "Message too long."
    Case 10041
      WSAGetErrorString = "Protocol wrong type for Socket."
    Case 10042
      WSAGetErrorString = "Protocol not available."
    Case 10043
      WSAGetErrorString = "Protocol not supported."
    Case 10044
      WSAGetErrorString = "Socket type not supported."
    Case 10045
      WSAGetErrorString = "Operation not supported on Socket."
    Case 10046
      WSAGetErrorString = "Protocol family not supported."
    Case 10047
      WSAGetErrorString = "Address family not supported by protocol family."
    Case 10048
      WSAGetErrorString = "Address already in use."
    Case 10049
      WSAGetErrorString = "Can't assign requested address."
    Case 10050
      WSAGetErrorString = "Network is down."
    Case 10051
      WSAGetErrorString = "Network is unreachable."
    Case 10052
      WSAGetErrorString = "Network dropped connection."
    Case 10053
      WSAGetErrorString = "Software caused connection abort."
    Case 10054
      WSAGetErrorString = "Connection reset by peer."
    Case 10055
      WSAGetErrorString = "No buffer space available."
    Case 10056
      WSAGetErrorString = "Socket is already connected."
    Case 10057
      WSAGetErrorString = "Socket is not connected."
    Case 10058
      WSAGetErrorString = "Can't send after Socket shutdown."
    Case 10059
      WSAGetErrorString = "Too many references: can't splice."
    Case 10060
      WSAGetErrorString = "Connection timed out."
    Case 10061
      WSAGetErrorString = "Connection refused."
    Case 10062
      WSAGetErrorString = "Too many levels of symbolic links."
    Case 10063
      WSAGetErrorString = "File name too long."
    Case 10064
      WSAGetErrorString = "Host is down."
    Case 10065
      WSAGetErrorString = "No route to host."
    Case 10066
      WSAGetErrorString = "Directory not empty."
    Case 10067
      WSAGetErrorString = "Too many processes."
    Case 10068
      WSAGetErrorString = "Too many users."
    Case 10069
      WSAGetErrorString = "Disk quota exceeded."
    Case 10070
      WSAGetErrorString = "Stale NFS file handle."
    Case 10071
      WSAGetErrorString = "Too many levels of remote in path."
    Case 10091
      WSAGetErrorString = "Network subsystem is unusable."
    Case 10092
      WSAGetErrorString = "Winsock DLL cannot support this application."
    Case 10093
      WSAGetErrorString = "Winsock not initialized."
    Case 10101
      WSAGetErrorString = "Disconnect."
    Case 11001
      WSAGetErrorString = "Host not found."
    Case 11002
      WSAGetErrorString = "Nonauthoritative host not found."
    Case 11003
      WSAGetErrorString = "Nonrecoverable error."
    Case 11004
      WSAGetErrorString = "Valid name, no data record of requested type."
    Case Else
      WSAGetErrorString = "unknown error " & errnum
  End Select
End Function

',-= WSABuildSocketAddress - Generiert eine SocketAdress aus einer IP
Public Function WSABuildSocketAddress(IP As String) As String
  Dim Family As Long
  Dim SocketAddress As String
  Dim SocketAddressLen As Long
  Dim Address As String
  Dim RetCode As Long
  Dim FailCode As Long

  If Not IsValidIP(IP) Then
    IP = GetCacheIP(IP, True)
  End If
  
  ',-= IP Version abfragen
  If IsValidIPv4(IP) Then
    Family = AF_INET
  ElseIf IsValidIPv6(IP) Then
    Family = AF_INET6
  End If
  
  ',-= Adresse umwandeln
  SocketAddress = String(WSA_Limit_SocketAddress, 0)
  SocketAddressLen = WSA_Limit_SocketAddress
  Address = IP
  RetCode = winsock2_WSAStringToAddressA(Address, Family, 0, SocketAddress, SocketAddressLen)
   
  ',-= Rückgabe verarbeiten
  If RetCode = -1 Then
    Err.Raise WSAGetLastError, , WSAGetErrorString(WSAGetLastError)
  Else
    WSABuildSocketAddress = Left(SocketAddress, SocketAddressLen)
  End If
End Function

',-= WSAGetHostByNameAlias - Generiert eine IP aus einem Hostnamen
Public Function WSAGetHostByNameAlias(ByVal HostName As String) As String
  Dim HostEntPointer As Long
  Dim HostEntDestination As HOSTENT
  Dim ListPointer As Long
  Dim AddressPointer As Long
  Dim AddressPointer2 As Long
  Dim ReturnIP(0 To 3) As Byte
  Dim ReturnAddress As String
  Dim Dummy As String
  ReturnAddress = WSABuildSocketAddress("0.0.0.0")
  HostEntPointer = winsock2_gethostbyname(HostName)
  If HostEntPointer <> 0 Then
    kernel32_RtlMoveMemory HostEntDestination, ByVal HostEntPointer, hostent_size
    If HostEntDestination.h_length = 4 Then
      ListPointer = HostEntDestination.h_addr_list
      kernel32_RtlMoveMemory AddressPointer, ByVal ListPointer, 4
      While AddressPointer <> 0
        kernel32_RtlMoveMemory ReturnIP(0), ByVal AddressPointer, 4
        ListPointer = ListPointer + 4
        kernel32_RtlMoveMemory AddressPointer, ByVal ListPointer, 4
        Mid(ReturnAddress, 5, 4) = StrConv(ReturnIP, vbUnicode)
        Dummy = WSAGetAscIP(ReturnAddress)
        If Not Dummy = "255.255.255.255" Then
          AddCacheData HostName, Dummy
        End If
      Wend
    End If
  End If
  WSAGetHostByNameAlias = Dummy
End Function

',-= WSAGetHostByNameAlias2 - Generiert die letzte IP aus einem Hostnamen
Public Function WSAGetHostByNameAlias2(ByVal HostName As String) As String
  Dim HostEntPointer As Long
  Dim HostEntDestination As HOSTENT
  Dim ListPointer As Long
  Dim AddressPointer As Long
  Dim AddressPointer2 As Long
  Dim ReturnIP(0 To 3) As Byte
  Dim ReturnAddress As String
  Dim Dummy As String
  ReturnAddress = WSABuildSocketAddress("0.0.0.0")
  HostEntPointer = winsock2_gethostbyname(HostName)
  If HostEntPointer <> 0 Then
    kernel32_RtlMoveMemory HostEntDestination, ByVal HostEntPointer, hostent_size
    If HostEntDestination.h_length = 4 Then
      ListPointer = HostEntDestination.h_addr_list
      kernel32_RtlMoveMemory AddressPointer, ByVal ListPointer, 4
      While AddressPointer <> 0
        ListPointer = ListPointer + 4
        kernel32_RtlMoveMemory ReturnIP(0), ByVal AddressPointer, 4
        kernel32_RtlMoveMemory AddressPointer, ByVal ListPointer, 4
      Wend
      Mid(ReturnAddress, 5, 4) = StrConv(ReturnIP, vbUnicode)
      Dummy = WSAGetAscIP(ReturnAddress)
    End If
  End If
  WSAGetHostByNameAlias2 = Dummy
End Function

',-= WSAGetHostByNameAlias - Generiert eine Hostnamen aus einer IP
Public Function WSAGetHostByAddress(ByVal IP As String) As String
  Dim HostEntPointer As Long
  Dim HostLength As Long
  Dim Address As Long
  Dim HostEntDestination As HOSTENT
  Dim HostName As String
  Dim ReturnAddress As String
  Dim AddressChar() As Byte
  ReturnAddress = WSABuildSocketAddress(IP)
  AddressChar = StrConv(Mid(ReturnAddress, 5, 4), vbFromUnicode)
  kernel32_RtlMoveMemory AddressChar(0), Address, 4
  HostEntPointer = winsock2_gethostbyaddr(Address, 4, AF_INET)
  If HostEntPointer Then
      kernel32_RtlMoveMemory HostEntDestination, ByVal HostEntPointer, hostent_size
      HostName = String(256, 0)
      HostLength = kernel32_lstrlenA(HostEntDestination.h_name)
      kernel32_RtlMoveMemory ByVal HostName, ByVal HostEntDestination.h_name, HostLength
      WSAGetHostByAddress = Left(HostName, HostLength)
  Else
      WSAGetHostByAddress = "Unknown"
  End If
End Function

',-= WSAGetAscIP - Gibt die IP einer SocketAddress in Text aus
Public Function WSAGetAscIP(ByVal SocketAddress As String) As String
  On Local Error Resume Next
  Dim IP As String
  Dim IPL As Long
  Dim RetCode As Long
  
  IP = String(WSA_Limit_SocketAddress, 0)
  IPL = WSA_Limit_SocketAddress
  
  ',-= Adresse anpassen
  SetPortInSocketAddress SocketAddress, 0
  If Err Then
    WSAGetAscIP = "255.255.255.255"
    Err.Clear
    Exit Function
  End If
  
  ',-= Adresse umwandeln
  RetCode = winsock2_WSAAddressToStringA(SocketAddress, Len(SocketAddress), 0, IP, IPL)
  If RetCode <> -1 Then
    WSAGetAscIP = Left(IP, IPL - 1)
  Else
    WSAGetAscIP = "255.255.255.255"
  End If
End Function

',-= WSA_GetSelectEvent - Gibt einen EventCode zurück
Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
  ',-= Wert prüfen und ausgeben
  If (lParam And &HFFFF&) > &H7FFF Then
    WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
  Else
    WSAGetSelectEvent = lParam And &HFFFF&
  End If
End Function

',-= WSA_GetAsyncError - Gibt einen FehlerCode zurück
Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
  ',-= Wert ausgeben
  WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

',-= WSASendData - Sendet Daten auf einem Socket
Public Function WSASendData(ByVal SocketNumber As Long, ByVal Message As Variant) As Long
  Dim TheMsg() As Byte, Temp As String, RetCode As Long
  
  ',-= Typ prüfen
  Select Case VarType(Message)
    Case 8209 ' byte array
      Temp = Message
    Case 8 ' String
      Temp = StrConv(Message, vbFromUnicode)
    Case Else
      Temp = CStr(Message)
      Temp = StrConv(Message, vbFromUnicode)
  End Select
  
  ',-= Daten umwandeln
  TheMsg = Temp
  
  ',-= Daten prüfen und ggf. senden
  If UBound(TheMsg) > -1 Then
    RetCode = winsock2_send(SocketNumber, TheMsg(0), UBound(TheMsg) + 1, 0)
    If RetCode = -1 Then
      WSASendData = WSAGetLastError
    Else
      WSASendData = 0
    End If
  End If
End Function

',-= WSASendDataTo - Sendet Daten auf einem Socket an ein spezielles Ziel
Public Function WSASendDataTo(ByVal SocketNumber As Long, ByVal TargetHost As String, ByVal TargetPort As Long, ByVal Data As String) As Long
  On Local Error Resume Next
  
  Dim SockAddr As String
  Dim SockAddrL As Long
  Dim APIVal As Long
  Dim MsgL As Long
  Dim msg() As Byte
  
  If IsValidIP(TargetHost) = False Then
    TargetHost = GetCacheIP(TargetHost, True)
  End If
  
  SockAddr = WSABuildSocketAddress(TargetHost)
  If Err Then
    WSASendDataTo = Err
    Err.Clear
    Exit Function
  End If
  SockAddrL = Len(SockAddr)
  
  SetPortInSocketAddress SockAddr, TargetPort
  
  msg = StrConv(Data, vbFromUnicode)
  MsgL = UBound(msg) + 1
  APIVal = winsock2_sendto(SocketNumber, msg(0), MsgL, 0, SockAddr, SockAddrL)
  If APIVal = -1 Then
    WSASendDataTo = WSAGetLastError
  Else
    WSASendDataTo = 0
  End If
End Function

',-= WSAGetAscAddr - Gibt die IP samt Port einer SocketAddress in Text aus
Public Function WSAGetAscAddr(ByVal SocketAddress As String) As String
  Dim IP As String
  Dim IPL As Long
  Dim RetCode As Long
  
  IP = String(WSA_Limit_SocketAddress, 0)
  IPL = WSA_Limit_SocketAddress
  
  ',-= Adresse umwandeln
  RetCode = winsock2_WSAAddressToStringA(SocketAddress, Len(SocketAddress), 0, IP, IPL)
  If RetCode <> -1 Then
    WSAGetAscAddr = Replace(Left(IP, IPL - 1), "0.0.0.0", "*")
    If WSAGetAscAddr = "*" Then WSAGetAscAddr = "*:*"
  Else
    WSAGetAscAddr = ""
  End If
End Function

',-= WSASetSockLinger - Aktiviert den Linger auf einem Socket
Public Function WSASetSockLinger(ByVal SocketNumber As Long, ByVal OnOff As Long, ByVal LingerTime As Long) As Long
  Dim Linger As LingerType
  Linger.l_onoff = OnOff
  Linger.l_linger = LingerTime
  If winsock2_setsockopt(SocketNumber, SOL_Socket, SO_LINGER, Linger, 4) = -1 Then
    WSASetSockLinger = -1
    Exit Function
  Else
    If winsock2_getsockopt(SocketNumber, SOL_Socket, SO_LINGER, Linger, 4) = -1 Then
      WSASetSockLinger = -1
    Else
      WSASetSockLinger = 0
    End If
  End If
End Function


Public Function WSAWindowMessage(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If iMsg = WM_Sockets Then
    WSAWindowMessage = WSASocketsMessage(hwnd, iMsg, wParam, lParam)
  ElseIf iMsg = WM_Resolve Then
    WSAWindowMessage = WSAResolveMessage(hwnd, iMsg, wParam, lParam)
  Else
    WSAWindowMessage = user32_CallWindowProcA(WSAOldWindowProc, ByVal hwnd, ByVal iMsg, ByVal wParam, ByVal lParam)
  End If
End Function

Public Function WSASocketsMessage(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim Index As Long, Dummy As String, RetCode As Long
  Dim WSAEvent As Long, WSAError As Long
  Dim SockNum As Long, SockAddr As String, SockAddrL As Long
  Dim Buffer As String, BufferLen As Long
  
  ',-= Nach SocketNummer suchen
  Index = FindSocket(wParam)
  
  ',-= Ungültige/Unbekannte Sockets sofort schließen
  If Index = -1 Then
    winsock2_closesocket wParam
    Exit Function
  End If
  
  SocketItem(Index).LastEvent = Now
  
  ',-= Event und Fehler ermitteln
  WSAEvent = WSAGetSelectEvent(lParam)
  WSAError = WSAGetAsyncError(lParam)
  
  ',-= Event abfragen
  If WSAEvent = FD_READ Then
    If SocketItem(Index).SocketType = SOCK_STREAM Then
      ',-= TCP/Stream Daten abfragen
      Buffer = String(5000, 0)
      BufferLen = winsock2_recv(wParam, Buffer, Len(Buffer), 0)
      While BufferLen > 0
        Buffer = Left(Buffer, BufferLen)
        SocketItem(Index).TrafficIn = SocketItem(Index).TrafficIn + BufferLen
        AddBytesIn Len(BufferLen)
        Socket_DataTCP Index, Buffer
        Buffer = String(5000, 0)
        BufferLen = winsock2_recv(wParam, Buffer, Len(Buffer), 0)
      Wend
    Else
      ',-= UDP/Datagram Daten abfragen
      SockAddr = String(WSA_Limit_SocketAddress, 0)
      SockAddrL = WSA_Limit_SocketAddress
      Buffer = String(5000, 0)
      BufferLen = winsock2_recvfrom(wParam, Buffer, Len(Buffer), 0, SockAddr, SockAddrL)
      While BufferLen > 0
        Buffer = Left(Buffer, BufferLen)
        SockAddr = Left(SockAddr, SockAddrL)
        Dummy = WSAGetAscIP(SockAddr)
        SockAddrL = (Asc(Mid(SockAddr, 3, 1)) * 256) + Asc(Mid(SockAddr, 4, 1))
        Socket_DataUDP Index, Buffer, Dummy, SockAddrL
        SockAddr = String(WSA_Limit_SocketAddress, 0)
        SockAddrL = WSA_Limit_SocketAddress
        Buffer = String(5000, 0)
        BufferLen = winsock2_recvfrom(wParam, Buffer, Len(Buffer), 0, SockAddr, SockAddrL)
      Wend
    End If
  ElseIf WSAEvent = FD_WRITE Then
    ',-= Nichts tun und Däumchen drehen? *G*
  ElseIf WSAEvent = FD_CLOSE Then
    RemoveSocket Index, WSAError, WSAGetErrorString(WSAError), False
  ElseIf WSAEvent = FD_winsock2_accept Then
    If WSAError = 0 Then
      ',-= Freien Index ermitteln
      BufferLen = AddSocket
    
      ',-= Verbindung annehmen
      SockAddr = String(WSA_Limit_SocketAddress, 0)
      SockAddrL = WSA_Limit_SocketAddress
      SockNum = winsock2_accept(wParam, SockAddr, SockAddrL)
      SockAddr = Left(SockAddr, SockAddrL)
      SocketItem(BufferLen).SocketType = SocketItem(Index).SocketType
      SocketItem(BufferLen).RemoteAddress = SockAddr
      SocketItem(BufferLen).SocketNumber = SockNum
      SocketItem(BufferLen).SocketStatus = SS_Connected
      SocketItem(BufferLen).SocketDirection = SD_In
      
      
      ',-= Lokale Adresse ermitteln
      SockAddr = String(WSA_Limit_SocketAddress, 0)
      SockAddrL = WSA_Limit_SocketAddress
      If winsock2_getsockname(SockNum, SockAddr, SockAddrL) = 0 Then
        SockAddr = Left(SockAddr, SockAddrL)
        SocketItem(BufferLen).LocalAddress = SockAddr
      End If
      Buffer = WSAGetAscIP(SocketItem(BufferLen).RemoteAddress)
      SockAddrL = GetPortFromSocketAddress(SocketItem(BufferLen).RemoteAddress)
      Socket_Incoming Index, BufferLen, Buffer, SockAddrL
    Else
      RemoveSocket Index, WSAError, WSAGetErrorString(WSAError), True
    End If
  ElseIf WSAEvent = FD_connect Then
    If WSAError = 0 Then
      ',-= Verbindung hergestellt
      
      ',-= Peer Adresse ermitteln
      SockAddr = String(WSA_Limit_SocketAddress, 0)
      SockAddrL = WSA_Limit_SocketAddress
      If winsock2_getpeername(wParam, SockAddr, SockAddrL) = 0 Then
        SockAddr = Left(SockAddr, SockAddrL)
        SocketItem(Index).RemoteAddress = SockAddr
      End If
      
      ',-= Lokale Adresse ermitteln
      SockAddr = String(WSA_Limit_SocketAddress, 0)
      SockAddrL = WSA_Limit_SocketAddress
      If winsock2_getsockname(wParam, SockAddr, SockAddrL) = 0 Then
        SockAddr = Left(SockAddr, SockAddrL)
        SocketItem(Index).LocalAddress = SockAddr
      End If
      SocketItem(Index).SocketStatus = SS_Connected
      SocketItem(Index).SocketDirection = SD_Out
      Socket_Connected Index
    Else
      ',-= Verbindung fehlgeschlagen
      Socket_ConnectError Index, WSAError, WSAGetErrorString(WSAError)
    End If
  ElseIf WSAEvent = FD_OOB Then
    '2do',-= Nichts tun und Däumchen drehen? *G*
  Else
    RemoveSocket Index, -1, WSAGetErrorString(WSAError), True
  End If
End Function

Public Function WSAResolveMessage(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim WSAError As Long
  Dim AddressPointer As Long
  Dim ReturnAddress As String
  Dim Index As Long
  Dim Index2 As Long
  Dim HostAddress As String
  Dim HostAlias As String
  Dim ReturnIP() As Byte
  Dim ReturnLength As Long
  
  WSAError = WSAGetAsyncError(lParam)
  
  If WSAError = 0 And HostEntBuffer.h_length = 4 Then
    ' Strings vorbereiten
    ReturnAddress = AddressIPv4
    ReDim Preserve ReturnIP(0 To 3)
    
    ' Addresse kopieren
    kernel32_RtlMoveMemory AddressPointer, ByVal HostEntBuffer.h_addr_list, 4
    kernel32_RtlMoveMemory ReturnIP(0), ByVal AddressPointer, 4
    
    ' Addresse in die ReturnAddress kopieren
    Mid(ReturnAddress, 5, 4) = StrConv(ReturnIP, vbUnicode)

    ' Aufgelöste IP in den Cache eintragen
    Index2 = 0
    For Index = 1 To ResolveCount
      If ResolveItem(Index).ASync = wParam Then
        HostAddress = ResolveItem(Index).Host
        HostAlias = WSAGetAscIP(ReturnAddress)
        If HostAddress = HostAlias Then
          ReturnLength = kernel32_lstrlenA(HostEntBuffer.h_name)
          If ReturnLength > 256 Then ReturnLength = 256
          ReDim ReturnIP(ReturnLength - 1)
          kernel32_RtlMoveMemory ReturnIP(0), ByVal HostEntBuffer.h_name, ReturnLength
          HostAddress = StrConv(ReturnIP, vbUnicode)
        End If
        
        Socket_ResolvedHost HostAddress, HostAlias
        Index2 = Index
        Exit For
      End If
    Next Index
  
    If Index2 > 0 Then
      ' Aufgelöste IP aus der Resolve Liste streichen
      For Index = Index2 To ResolveCount - 1
        ResolveItem(Index) = ResolveItem(Index + 1)
      Next Index
      ResolveItem(ResolveCount).ASync = -1
      ResolveItem(ResolveCount).Host = ""
      ResolveCount = ResolveCount - 1
      If ResolveCount < UBound(ResolveItem) - 5 Then ReDim Preserve ResolveItem(((ResolveCount / 5) + 1) * ResolveCount)
    End If
  Else
    ' Fehler beim Auflösen
    For Index = 1 To ResolveCount
      If ResolveItem(Index).ASync = wParam Then
        Index2 = Index
        Exit For
      End If
    Next Index
    If Index2 > 0 Then
      Socket_ResolveFailed ResolveItem(Index2).Host
      ' Anfrage aus der Resolve Liste streichen
      For Index = Index2 To ResolveCount - 1
        ResolveItem(Index) = ResolveItem(Index + 1)
      Next Index
      ResolveItem(ResolveCount).ASync = -1
      ResolveItem(ResolveCount).Host = ""
      ResolveCount = ResolveCount - 1
      If ResolveCount < UBound(ResolveItem) - 5 Then ReDim Preserve ResolveItem(((ResolveCount / 5) + 1) * ResolveCount)
    End If
  End If
  ResolveBlocking = 0
  For Index = 1 To ResolveCount
    If ResolveItem(Index).ASync = 0 And ResolveItem(Index).Host <> "" Then
      If ResolveBlocking = 0 Then
        ResolveItem(Index).ASync = WSAASyncGetHost(ResolveItem(Index).Host)
        ResolveBlocking = 1
      Else
        Exit For
      End If
    End If
  Next Index
End Function

Function WSARandomPort() As Long
  Dim PortRangeMin As Long, PortRangeMax As Long, PortRangeDif As Long
  On Local Error Resume Next
  PortRangeMin = CLng(Mid(PortRange, 1, InStr(PortRange, "-") - 1))
  PortRangeMax = CLng(Mid(PortRange, InStr(PortRange, "-") + 1))
  PortRangeDif = PortRangeMax - PortRangeMin + 2
  If Err Or PortRangeDif < 1 Then
    Err.Clear
    PortRangeMin = 1025
    PortRangeMax = 65535
    PortRangeDif = PortRangeMax - PortRangeMin + 2
  End If
  WSARandomPort = PortRangeMin + CLng((Rnd() * PortRangeDif))
End Function


Public Function SendTCP(vSocket As Long, Data As String)
  Dim Index As Long
  If Not SocketItem(vSocket) Is Nothing Then
    If SocketItem(vSocket).SocketNumber <> -1 Then
      If SocketItem(vSocket).SocketType = SOCK_STREAM Then
        Index = WSASendData(SocketItem(vSocket).SocketNumber, Data)
        If Index = 0 Then
          SendTCP = 0
          SocketItem(vSocket).Trafficout = SocketItem(vSocket).Trafficout + Len(Data)
          AddBytesOut Len(Data)
        Else
          SendTCP = Index
        End If
      Else
        SendTCP = -1
      End If
    Else
      SendTCP = -1
    End If
  Else
    SendTCP = -1
  End If
End Function

Public Function SendUDP(vSocket As Long, TargetHost As String, TargetPort As Long, Data As String)
  Dim Index As Long
  If Not SocketItem(vSocket) Is Nothing Then
    If SocketItem(vSocket).SocketNumber <> -1 Then
      If SocketItem(vSocket).SocketType = SOCK_DGRAM Then
        Index = WSASendDataTo(SocketItem(vSocket).SocketNumber, TargetHost, TargetPort, Data)
        If Index = 0 Then
          SendUDP = 0
          SocketItem(vSocket).Trafficout = SocketItem(vSocket).Trafficout + Len(Data)
          AddBytesOut Len(Data)
        Else
          SendUDP = Index
        End If
      Else
        SendUDP = -1
      End If
    Else
      SendUDP = -1
    End If
  Else
    SendUDP = -1
  End If
End Function

Public Function ConnectTCP(vSocket As Long, Host As String, Port As Long) As Long
  On Local Error Resume Next
  Dim Index As Long
  Dim Index2 As Long
  Dim Index3 As Long
  Dim Index4 As Long
  Dim Index5 As Long
  Dim Index6 As Long
  Dim RetCode As Long
  
  Dim SockAddr As String
  Dim SockAddrL As Long
  
  Dim SockAddr2 As String
  Dim SockAddr2L As String
  
  ',-= vSocket prüfen
  If Not IsValidSocket(vSocket) Then
    ConnectTCP = -1
    Exit Function
  End If
  
  ',-= SockAddr holen
  SockAddr = WSABuildSocketAddress(Host)
  If Err Then
    ConnectTCP = Err.Number
    Err.Clear
    Exit Function
  End If
  SetPortInSocketAddress SockAddr, Port
  If Err Then
    ConnectTCP = Err.Number
    Err.Clear
    Exit Function
  End If
  SockAddrL = Len(SockAddr)
  
  ',-= Familie auslesen
  Index = Asc(Mid(SockAddr, 1, 1))
  
  ',-= Default Addresse für das Ziel setzen
  If Index = AF_INET Then
    SockAddr2 = AddressIPv4
  ElseIf Index = AF_INET6 Then
    SockAddr2 = AddressIPv6
  End If
  SockAddr2L = Len(SockAddr)
  If SockAddr2 = "" Then
    ConnectTCP = 10014
    Exit Function
  End If
  
  ',-= Socket erstellen
  Index4 = winsock2_socket(Index, SOCK_STREAM, IPPROTO_TCP)
  
  If Index4 = -1 Then
    ConnectTCP = WSAGetLastError
    Exit Function
  End If
  
  Index = 0
  Index2 = -1
  While Index2 = -1
    ',-= Port generieren
    Index3 = WSARandomPort
    
    ',-= Port in die SockAddr schreiben
    SetPortInSocketAddress SockAddr2, Index3

    ',-= Socked binden
    Index2 = winsock2_bind(Index4, SockAddr2, SockAddr2L)
    
    If Index2 = -1 Then
      ',-= Versuchszähler erhöhen
      Index = Index + 1
      If Index = 25 Then
        winsock2_closesocket Index4
        ConnectTCP = 10048
        Exit Function
      End If
    End If
  Wend
  
  ',-= Linger setzen
  If WSASetSockLinger(Index4, 1, 0) = -1 Then
    winsock2_closesocket Index4
    ConnectTCP = WSAGetLastError
    Exit Function
  End If
  
  ',-= Events setzen
  Index = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_connect
  If winsock2_WSAAsyncSelect(Index4, WSAWindowHandle, WM_Sockets, Index) = -1 Then
    winsock2_closesocket Index4
    ConnectTCP = WSAGetLastError
    Exit Function
  End If

  If winsock2_connect(Index4, SockAddr, SockAddrL) = -1 Then
    SocketItem(vSocket).SocketDirection = SD_Out
    SocketItem(vSocket).SocketStatus = SS_Connecting
    SocketItem(vSocket).SocketNumber = Index4
    SocketItem(vSocket).SocketType = SOCK_STREAM
    SocketItem(vSocket).LocalAddress = SockAddr2
    SocketItem(vSocket).RemoteAddress = SockAddr
    ConnectTCP = 0
  Else
    winsock2_closesocket Index4
    ConnectTCP = WSAGetLastError
    Exit Function
  End If
End Function

',-= ListenTCP - Beginnt auf einer vSock zu lauschen
Public Function ListenTCP(vSocket As Long, Port As Long, Optional ListenIP As String = "") As Long
  Dim Index As Long
  Dim Index2 As Long
  Dim Index3 As Long
  
  Dim SockAddr As String
  Dim SockAddrL As Long
  
  On Local Error Resume Next
  
  ',-= vSocket prüfen
  If Not IsValidSocket(vSocket) Then
    ListenTCP = -1
    Exit Function
  End If
  
  ',-= SockAddr holen
  If ListenIP = "" Then
    SockAddr = AddressDefault
  Else
    SockAddr = WSABuildSocketAddress(ListenIP)
    If Err Then
      ListenTCP = Err.Number
      Err.Clear
      Exit Function
    End If
  End If
  SockAddrL = Len(SockAddr)
  
  ',-= Port in die SockAddr schreiben
  SetPortInSocketAddress SockAddr, Port
  If Err Then
    ListenTCP = Err.Number
    Err.Clear
    Exit Function
  End If
  
  ',-= Familie auslesen
  Index = Asc(Mid(SockAddr, 1, 1))
  
  ',-= Socket erstellen
  Index2 = winsock2_socket(Index, SOCK_STREAM, IPPROTO_TCP)
  If Index2 = -1 Then
    ListenTCP = WSAGetLastError
    Exit Function
  End If
    
  ',-= Socket erstellt - An SockAddr binden
  If winsock2_bind(Index2, SockAddr, SockAddrL) = -1 Then
    ListenTCP = WSAGetLastError
    winsock2_closesocket Index2
    Exit Function
  End If
    
  ',-= Events setzen
  Index3 = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_winsock2_accept
  If winsock2_WSAAsyncSelect(Index2, WSAWindowHandle, WM_Sockets, Index3) = -1 Then
    winsock2_closesocket Index2
    ListenTCP = WSAGetLastError
    Exit Function
  End If
      
  ',-= Socket gebunden - listen anfangen
  If winsock2_listen(Index2, 1) = -1 Then
    winsock2_closesocket Index2
    ListenTCP = WSAGetLastError
    Exit Function
  End If
  
  ',-= Fertig
  SocketItem(vSocket).SocketDirection = SD_In
  SocketItem(vSocket).RemoteAddress = IIf(Index = AF_INET, AddressIPv4, AddressIPv6)
  SocketItem(vSocket).SocketStatus = SS_Listening
  SocketItem(vSocket).SocketType = SOCK_STREAM
  SocketItem(vSocket).SocketNumber = Index2
  SocketItem(vSocket).LocalAddress = SockAddr
  ListenTCP = 0
End Function


Public Function ListenUDP(vSocket As Long, Port As Long, Optional ListenIP As String = "") As Long
  Dim Index As Long
  Dim Index2 As Long
  Dim Index3 As Long
  
  Dim SockAddr As String
  Dim SockAddrL As Long
  
  On Local Error Resume Next
  
  ' vSocket prüfen
  If Not IsValidSocket(vSocket) Then
    ListenUDP = -1
    Exit Function
  End If
  
  ' SockAddr holen
  If ListenIP = "" Then
    SockAddr = AddressIPv4
  Else
    SockAddr = WSABuildSocketAddress(ListenIP)
    If Err Then
      ListenUDP = Err.Number
      Err.Clear
      Exit Function
    End If
  End If
  SockAddrL = Len(SockAddr)
  
  ' Port in die SockAddr schreiben
  SetPortInSocketAddress SockAddr, Port
  
  ' Familie auslesen
  Index = Asc(Mid(SockAddr, 1, 1))
  
  ' Socket erstellen
  Index2 = winsock2_socket(Index, SOCK_DGRAM, IPPROTO_UDP)
  If Index2 = -1 Then
    ListenUDP = WSAGetLastError
    Exit Function
  End If
    
  ' Socket erstellt - An SockAddr binden
  If winsock2_bind(Index2, SockAddr, SockAddrL) = -1 Then
    winsock2_closesocket Index2
    ListenUDP = WSAGetLastError
    Exit Function
  End If
  
  ' Broadcasts aktivieren
  Index3 = -1
  If winsock2_setsockopt(Index2, SOL_Socket, SO_BROADCAST, Index, Len(Index)) = -1 Then
    winsock2_closesocket Index2
    ListenUDP = WSAGetLastError
    Exit Function
  End If
  
  ' Events setzen
  Index3 = FD_READ Or FD_WRITE
  If winsock2_WSAAsyncSelect(Index2, WSAWindowHandle, WM_Sockets, Index3) = -1 Then
    winsock2_closesocket Index2
    ListenUDP = WSAGetLastError
    Exit Function
  End If
      
  ' Socket gebunden - listen anfangen
  SockAddr = String(WSA_Limit_SocketAddress, 0)
  SockAddrL = Len(SockAddr)
  If winsock2_getsockname(Index2, SockAddr, SockAddrL) = -1 Then
    winsock2_closesocket Index2
    ListenUDP = WSAGetLastError
    Exit Function
  End If
  
  ' Fertig
  SocketItem(vSocket).SocketDirection = SD_In
  SocketItem(vSocket).RemoteAddress = IIf(Index = AF_INET, AddressIPv4, AddressIPv6)
  SocketItem(vSocket).SocketStatus = SS_Listening
  SocketItem(vSocket).SocketType = SOCK_DGRAM
  SocketItem(vSocket).SocketNumber = Index2
  SocketItem(vSocket).LocalAddress = SockAddr
  ListenUDP = 0
End Function

Public Function WSAGetLocalHostName() As String
  Dim sName As String
  sName = String(256, 0)
  If winsock2_gethostname(sName, 256) Then
    sName = "Unknown"
  Else
    If InStr(sName, Chr(0)) Then
      sName = Left(sName, InStr(sName, Chr(0)) - 1)
    End If
  End If
  WSAGetLocalHostName = sName
End Function

'Converts an IP string (i.e. "127.0.0.1") to a socks compatible 4-byte string
Public Function MakeString(Text As String) As String
  Dim u As Long, Part As String, NewString As String
  For u = 1 To Len(Text)
    If Mid(Text, u, 1) = "." Then
      NewString = NewString + Chr(Val(Part))
      Part = ""
    Else
      Part = Part + Mid(Text, u, 1)
    End If
  Next u
  NewString = NewString + Chr(Val(Part))
  MakeString = NewString
End Function

