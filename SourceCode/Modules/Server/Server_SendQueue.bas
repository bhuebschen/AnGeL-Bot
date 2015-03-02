Attribute VB_Name = "Server_SendQueue"
',-======================- ==-- -  -
'|   AnGeL - Server - SendQueue
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit
  

Private Type SQueue
  LineBuffer() As String
  BufferedLines As Long
End Type


Public Buffer(1 To 3) As SQueue


Sub SendQueue_Load()
  ReDim Preserve Buffer(1).LineBuffer(5)
  ReDim Preserve Buffer(2).LineBuffer(5)
  ReDim Preserve Buffer(3).LineBuffer(5)
End Sub

Sub SendQueue_Unload()
'
End Sub

Public Sub SendIt(ByVal What As String)
  If ServerSocket > 0 Then SendTCP ServerSocket, What
End Sub

