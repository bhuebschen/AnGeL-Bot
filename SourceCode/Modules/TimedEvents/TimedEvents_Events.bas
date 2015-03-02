Attribute VB_Name = "TimedEvents_Events"
',-======================- ==-- -  -
'|   AnGeL - TimedEvents - Events
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Type TEvent
  DoThis As String
  AtTime As Currency
End Type

Public EventCount As Long
Public Events() As TEvent


Sub Events_Load()
  ReDim Preserve Events(5)
End Sub

Sub Events_Unload()
'
End Sub


Public Function IsTimed(DoThis As String) As Boolean
  Dim Index As Long
  For Index = 1 To EventCount
    If LCase(Events(Index).DoThis) = LCase(DoThis) Then
      IsTimed = True
      Exit Function
    End If
  Next Index
  IsTimed = False
End Function


Public Sub RemoveTimedEvent(SStart As String)
  Dim Index As Long
  For Index = 1 To EventCount
    If Param(Events(Index).DoThis, 1) = SStart Then
      Events(Index).DoThis = ""
    End If
  Next Index
End Sub


Public Sub RemoveTimedEventFromList(Num As Long)
  Dim Index As Long
  For Index = Num To EventCount - 1
    Events(Index) = Events(Index + 1)
  Next Index
  EventCount = EventCount - 1
  Index = ((EventCount \ 50) + 1) * 50
  If Index < UBound(Events()) Then ReDim Preserve Events(Index)
End Sub


Public Sub TimedEvent(DoThis As String, AtTime As Currency)
  EventCount = EventCount + 1
  If EventCount > UBound(Events()) Then ReDim Preserve Events(UBound(Events()) + 5)
  Events(EventCount).DoThis = DoThis
  Events(EventCount).AtTime = WinTickCount + AtTime * 1000
End Sub


Public Sub FixTimedEvent(DoThis As String, AtTime As Currency)
  EventCount = EventCount + 1
  If EventCount > UBound(Events()) Then ReDim Preserve Events(UBound(Events()) + 5)
  Events(EventCount).DoThis = DoThis
  Events(EventCount).AtTime = AtTime
End Sub
