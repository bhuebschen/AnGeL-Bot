VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sockets_SocketItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
',-======================- ==-- -  -
'|   AnGeL - Sockets - SocketObjekt
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

',-= WinSock Informationen
Public SocketDirection As Byte
Public SocketStatus As Byte
Public SocketNumber As Long
Public SocketType As Byte
Public TrafficIn As Currency
Public Trafficout As Currency
Public RemoteAddress As String
Public LocalAddress As String
Public SendQTries As Byte
Public LastEvent As Date

',-= vSocket Informationen
Public Used As Boolean
Public IsInternalSocket As Boolean
Public NumOfServerEvents As Long
Public Hostmask As String
Public IRCNick As String
Public RegNick As String
Public OnBot As String
Public Flags As String
Public AwayMessage As String
Public UserNum As Long
Public FileAreaDir As String
Public FileName As String
Public FileNum As Integer
Public FileSize As Long
Public CurrentQuestion As String
Public SetupChan As String
Public InputBuffer As String
Public OrderSign As String
Public LinkStatus As String
Public PLChannel As Long
Public SockTag As String
Public BytesReceived As Long

',-= Interne Werte
Private mySendQ() As String
Private mySendQLines As Long
Private mySocketFlags As String
Private Const mySendQStepping As Long = 3


Public Property Get SendQ(Index As Long) As String
  If Index <= 0 Then Exit Sub
  If Index > SendQLines Then Exit Sub
  SendQ = mySendQ(Index)
End Property
Public Property Let SendQ(Index As Long, ByVal vNewValue As String)
  If Index <= 0 Then Exit Property
  If Index > SendQLines Then Exit Property
  mySendQ(Index) = vNewValue
End Property


Public Property Get SendQLines() As Long
  SendQLines = mySendQLines
End Property
Public Property Let SendQLines(ByVal vNewValue As Long)
  Dim Index As Integer
  If vNewValue < 0 Then vNewValue = 0
  If vNewValue < mySendQLines Then
    If UBound(mySendQ) - vNewValue >= mySendQStepping Then ReDim Preserve mySendQ(((vNewValue / mySendQStepping) + 1) * mySendQStepping)
    For Index = vNewValue + 1 To UBound(mySendQ)
      mySendQ(Index) = ""
    Next Index
    mySendQLines = vNewValue
  ElseIf vNewValue > SendQLines Then
    If vNewValue > UBound(mySendQ) Then ReDim Preserve mySendQ(((vNewValue / mySendQStepping) + 1) * mySendQStepping)
    mySendQLines = vNewValue
  End If
End Property


Public Property Get SocketFlag(Flag As Byte) As String
  If Flag <= 0 Then Exit Property
  If Flag >= 30 Then Exit Property
  If Len(mySocketFlags) < Flag Then
    SocketFlag = "0"
  Else
    SocketFlag = Mid(mySocketFlags, Flag, 1)
  End If
End Property
Public Property Let SocketFlag(Flag As Byte, ByVal vNewValue As String)
  If Flag <= 0 Then Exit Property
  If Flag >= 30 Then Exit Property
  If Len(mySocketFlags) < Flag Then
    mySocketFlags = mySocketFlags & String(Flag - Len(mySocketFlags), "0")
  End If
  Mid(mySocketFlags, Flag, 1) = vNewValue
End Property
Public Property Get SocketFlags() As String
  SocketFlags = mySocketFlags
End Property


Public Property Get LocalPort() As Long
  If SocketNumber = -1 Then
    LocalPort = 0
  Else
    LocalPort = GetPortFromSocketAddress(LocalAddress)
  End If
End Property

Public Property Get RemotePort() As Long
  If SocketNumber = -1 Then
    RemotePort = 0
  Else
    RemotePort = GetPortFromSocketAddress(RemoteAddress)
  End If
End Property

Public Property Get LocalHost() As String
  If SocketNumber = -1 Then
    LocalHost = ""
  Else
    LocalHost = WSAGetAscIP(LocalAddress)
  End If
End Property

Public Property Get RemoteHost() As String
  If SocketNumber = -1 Then
    RemoteHost = ""
  Else
    RemoteHost = WSAGetAscIP(RemoteAddress)
  End If
End Property


Private Sub Class_Initialize()
  ReDim mySendQ(mySendQStepping)
  LastEvent = Now
End Sub
