Attribute VB_Name = "Sockets_Base"
',-======================- ==-- -  -
'|   AnGeL - Sockets - Hauptmodul
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Sub Sockets_Load()
  Dim Dummy As Long
  ',-= Zähler zurücksetzen
  SocketCount = 0
  SocketUse = 0
  IPv6Caps = 1
  ResolveCount = 0
  
  ',-= Message registrieren
  WM_Sockets = RegisterWindowMessage("WinSockSockets")
  WM_Resolve = RegisterWindowMessage("WinSockResolve")

  ',-= Winsock StartUp
  If winsock2_WSAStartup(257, WSAData) = -1 Then
    Socket_Error -1, True, 10000, "Winsock zu alt"
    Exit Sub
  End If
  
  ',-= Erste Elemente erstellen
  ReDim SocketItem(SocketStepping)
  ReDim ResolveItem(5)
  
  ',-= Standard Addressen setzens
  AddressDefault = WSABuildSocketAddress("0.0.0.0")
  AddressIPv4 = AddressDefault
  On Local Error Resume Next
  AddressIPv6 = WSABuildSocketAddress("::0.0.0.0")
  If Err Then
    IPv6Caps = 0
    AddressIPv6 = ""
    Err.Clear
  End If
  
  ',-= Fenster erstellen
  WSAWindowHandle = user32_CreateWindowExA(0&, "static", "AnGeL_Socket_Window", WS_OVERLAPPEDWINDOW, 5&, 1&, 200&, 100&, 0&, 0&, App.hInstance, ByVal 0&)
  
  ',-= An Winsock hooken
  WSAOldWindowProc = user32_SetWindowLongA(WSAWindowHandle, GWL_WNDPROC, AddressOf WSAWindowMessage)
End Sub

Public Sub Sockets_Unload()
  ',-= Aus Winsock ausklinken
  Call user32_SetWindowLongA(WSAWindowHandle, GWL_WNDPROC, WSAOldWindowProc)
  
  ',-= Winsock 'aufräumen'
  If winsock2_WSAIsBlocking = -1 Then
    winsock2_WSACancelBlockingCall
  End If
  winsock2_WSACleanup
  
  ',-= Fenster zerstören
  user32_DestroyWindow WSAWindowHandle
End Sub
