VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddresslist_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public pNick As String
Public pIdent As String
Public pHost As String
Public pIP As String
Public pUserlistNumber As Long


Private Sub Class_Initialize()
  pNick = ""
  pIdent = ""
  pHost = ""
  pIP = ""
  pUserlistNumber = 0
End Sub


Public Property Get pMemoryUsed() As Long
  pMemoryUsed = LenB(pNick) + LenB(pIdent) + LenB(pHost) + LenB(pIP) + LenB(pUserlistNumber)
End Property

