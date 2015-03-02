Attribute VB_Name = "Sockets_Resolve"
',-======================- ==-- -  -
'|   AnGeL - Sockets - Resolver
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


',-======================- ==-- -  -
'|   Typen
'`-=====================-- -===- ==- -- -
Public Type ResolveType
  ResolveID As Long
  Host As String
  ASync As Long
End Type

Public Type CacheType
  Host As String
  IP As String
End Type


',-======================- ==-- -  -
'|   Konstanten
'`-=====================-- -===- ==- -- -
Public Const HostCacheSize As Long = 200
Public Const DNS_UNRESOLVE = "255.255.255.255"


',-======================- ==-- -  -
'|   Variablen
'`-=====================-- -===- ==- -- -
Public HostCachePos As Long
Public HostCache(HostCacheSize) As CacheType
Public ResolveItem() As ResolveType
Public ResolveCount As Long
Public ResolveBlocking As Byte
Public HostEntZero As HostEntAsync
Public HostEntBuffer As HostEntAsync
Public WM_Resolve As Long


',-======================- ==-- -  -
'|   Funktionen
'`-=====================-- -===- ==- -- -
Private Function ASyncResolve(ByVal Host As String) As Long
  Dim Index As Long
  If Host = "" Then Exit Function
  For Index = 1 To ResolveCount
    If LCase(ResolveItem(Index).Host) = LCase(Host) Then Exit Function
  Next Index
  ResolveCount = ResolveCount + 1
  If ResolveCount >= UBound(ResolveItem) Then ReDim Preserve ResolveItem(((ResolveCount / 5) + 1) * ResolveCount)
  ResolveItem(ResolveCount).Host = Host
  If ResolveBlocking = 0 Then
    ResolveItem(ResolveCount).ASync = WSAASyncGetHost(ResolveItem(ResolveCount).Host)
    ResolveBlocking = 1
  Else
    ResolveItem(ResolveCount).ASync = 0
  End If
  ASyncResolve = ResolveItem(ResolveCount).ASync
End Function

Public Sub AddCacheData(Host As String, IP As String)
  If HostCachePos > HostCacheSize Then HostCachePos = 0
  HostCache(HostCachePos).Host = Host
  HostCache(HostCachePos).IP = IP
  HostCachePos = HostCachePos + 1
End Sub

Public Function GetCacheIP(Host As String, Instant As Boolean) As String
  Dim Index As Long, Dummy As String
  
  For Index = 0 To HostCacheSize
    If LCase(Host) = HostCache(Index).Host Then
      GetCacheIP = HostCache(Index).IP
      If CBool(CInt(Rnd)) = True Then Exit For
    ElseIf HostCache(Index).Host = "" Then
      Exit For
    End If
  Next Index
  
  If GetCacheIP = "" Then
    If Instant = True Then
      Dummy = WSAGetHostByNameAlias(Host)
      GetCacheIP = Dummy
    Else
      ASyncResolve Host
    End If
  End If
End Function

Public Function GetCacheHost(IP As String, Instant As Boolean) As String
  Dim Index As Long, Dummy As String

  For Index = 0 To HostCacheSize
    If LCase(IP) = HostCache(Index).IP Then
      GetCacheHost = HostCache(Index).Host
      If CBool(Int(Rnd)) = True Then Exit For
    ElseIf HostCache(Index).Host = "" Then
      Exit For
    End If
  Next Index
  
  If GetCacheHost = "" Then
    If Instant = True Then
      Dummy = WSAGetHostByAddress(IP)
    Else
      ASyncResolve IP
    End If
  End If
End Function

Public Function GetLastLocalIP() As String
  GetLastLocalIP = WSAGetHostByNameAlias2(WSAGetLocalHostName)
End Function
