Attribute VB_Name = "Plugin_Traffic"
',-======================- ==-- -  -
'|   AnGeL - Plugins - Traffic
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public GlobalBytesSent As Currency
Public GlobalBytesReceived As Currency


Private BytesInHistory(0 To 23) As Currency
Private BytesOutHistory(0 To 23) As Currency
Private BytesHour As Byte


Public SessionBytesSent As Currency
Public SessionBytesReceived As Currency


Sub AddBytesIn(Count As Long)
  CheckByteHour
  BytesInHistory(23) = BytesInHistory(23) + Count
  SessionBytesReceived = SessionBytesReceived + Count
  GlobalBytesReceived = GlobalBytesReceived + Count
End Sub

Sub AddBytesOut(Count As Long)
  CheckByteHour
  BytesOutHistory(23) = BytesOutHistory(23) + Count
  SessionBytesSent = SessionBytesSent + Count
  GlobalBytesSent = GlobalBytesSent + Count
End Sub

Private Sub CheckByteHour()
  Dim Index As Integer
  If Hour(Now) <> BytesHour Then
    For Index = 1 To 23
      BytesInHistory(Index - 1) = BytesInHistory(Index)
      BytesOutHistory(Index - 1) = BytesOutHistory(Index)
    Next Index
    BytesInHistory(23) = 0
    BytesOutHistory(23) = 0
    BytesHour = Hour(Now)
  End If
End Sub

Sub GUITrafficGraph(vSock As Long)
  Dim Index As Integer
  Dim Line As String
  Dim LineIn0(0 To 23) As String
  Dim LineOut0(0 To 23) As String
  Dim LineIn1(0 To 23) As String
  Dim LineOut1(0 To 23) As String
  Dim LineIn2(0 To 23) As String
  Dim LineOut2(0 To 23) As String
  Dim LineIn3(0 To 23) As String
  Dim LineOut3(0 To 23) As String
  Dim LineIn4(0 To 23) As String
  Dim LineOut4(0 To 23) As String
  
  Dim ScaleBytes As Long
  
  For Index = 0 To 23
    If BytesInHistory(Index) > ScaleBytes Then ScaleBytes = BytesInHistory(Index)
    If BytesOutHistory(Index) > ScaleBytes Then ScaleBytes = BytesOutHistory(Index)
  Next
  
  ScaleBytes = ScaleBytes - (ScaleBytes Mod 5)
  ScaleBytes = ScaleBytes / 5
  
  TU vSock, "2*** Traffic graph:"
  TU vSock, "0,1Bytes | In " & String(21, " ") & " | Out " & String(21, " ")
  
  For Index = 0 To 23
    If BytesInHistory(Index) > ScaleBytes * 5 Then
      LineIn0(Index) = "5,5#"
      LineIn1(Index) = "5,5#"
      LineIn2(Index) = "5,5#"
      LineIn3(Index) = "5,5#"
      LineIn4(Index) = "5,5#"
    ElseIf BytesInHistory(Index) > ScaleBytes * 4 Then
      LineIn0(Index) = "5_"
      LineIn1(Index) = "4,4#"
      LineIn2(Index) = "4,4#"
      LineIn3(Index) = "4,4#"
      LineIn4(Index) = "4,4#"
    ElseIf BytesInHistory(Index) > ScaleBytes * 3 Then
      LineIn0(Index) = " "
      LineIn1(Index) = "4_"
      LineIn2(Index) = "7,7#"
      LineIn3(Index) = "7,7#"
      LineIn4(Index) = "7,7#"
    ElseIf BytesInHistory(Index) > ScaleBytes * 2 Then
      LineIn0(Index) = " "
      LineIn1(Index) = " "
      LineIn2(Index) = "7_"
      LineIn3(Index) = "8,8#"
      LineIn4(Index) = "8,8#"
    ElseIf BytesInHistory(Index) > ScaleBytes Then
      LineIn0(Index) = " "
      LineIn1(Index) = " "
      LineIn2(Index) = " "
      LineIn3(Index) = "8_"
      LineIn4(Index) = "9,9#"
    ElseIf BytesInHistory(Index) <= ScaleBytes And BytesInHistory(Index) > 0 Then
      LineIn0(Index) = " "
      LineIn1(Index) = " "
      LineIn2(Index) = " "
      LineIn3(Index) = " "
      LineIn4(Index) = "9_"
    Else
      LineIn0(Index) = " "
      LineIn1(Index) = " "
      LineIn2(Index) = " "
      LineIn3(Index) = " "
      LineIn4(Index) = " "
    End If
    If BytesOutHistory(Index) > ScaleBytes * 5 Then
      LineOut0(Index) = "5,5#"
      LineOut1(Index) = "5,5#"
      LineOut2(Index) = "5,5#"
      LineOut3(Index) = "5,5#"
      LineOut4(Index) = "5,5#"
    ElseIf BytesOutHistory(Index) > ScaleBytes * 4 Then
      LineOut0(Index) = "5_"
      LineOut1(Index) = "4,4#"
      LineOut2(Index) = "4,4#"
      LineOut3(Index) = "4,4#"
      LineOut4(Index) = "4,4#"
    ElseIf BytesOutHistory(Index) > ScaleBytes * 3 Then
      LineOut0(Index) = " "
      LineOut1(Index) = "4_"
      LineOut2(Index) = "7,7#"
      LineOut3(Index) = "7,7#"
      LineOut4(Index) = "7,7#"
    ElseIf BytesOutHistory(Index) > ScaleBytes * 2 Then
      LineOut0(Index) = " "
      LineOut1(Index) = " "
      LineOut2(Index) = "7_"
      LineOut3(Index) = "8,8#"
      LineOut4(Index) = "8,8#"
    ElseIf BytesOutHistory(Index) > ScaleBytes Then
      LineOut0(Index) = " "
      LineOut1(Index) = " "
      LineOut2(Index) = " "
      LineOut3(Index) = "8_"
      LineOut4(Index) = "9,9#"
    ElseIf BytesOutHistory(Index) <= ScaleBytes And BytesOutHistory(Index) > 0 Then
      LineOut0(Index) = " "
      LineOut1(Index) = " "
      LineOut2(Index) = " "
      LineOut3(Index) = " "
      LineOut4(Index) = "9_"
    Else
      LineOut0(Index) = " "
      LineOut1(Index) = " "
      LineOut2(Index) = " "
      LineOut3(Index) = " "
      LineOut4(Index) = " "
    End If
  Next Index
  
  TU vSock, Spaces(5, SizeToString(ScaleBytes * 5)) & SizeToString(ScaleBytes * 5) & " | " & CStr(Join(LineIn0, "")) & " | " & CStr(Join(LineOut0, ""))
  TU vSock, Spaces(5, SizeToString(ScaleBytes * 4)) & SizeToString(ScaleBytes * 4) & " | " & CStr(Join(LineIn1, "")) & " | " & CStr(Join(LineOut1, ""))
  TU vSock, Spaces(5, SizeToString(ScaleBytes * 3)) & SizeToString(ScaleBytes * 3) & " | " & CStr(Join(LineIn2, "")) & " | " & CStr(Join(LineOut2, ""))
  TU vSock, Spaces(5, SizeToString(ScaleBytes * 2)) & SizeToString(ScaleBytes * 2) & " | " & CStr(Join(LineIn3, "")) & " | " & CStr(Join(LineOut3, ""))
  TU vSock, Spaces(5, SizeToString(ScaleBytes)) & SizeToString(ScaleBytes) & " | " & CStr(Join(LineIn4, "")) & " | " & CStr(Join(LineOut4, ""))
  TU vSock, String(60, "-")
  TU vSock, "Total | " & Spaces2(28, SizeToString(SessionBytesReceived) & " 14(" & SizeToString(GlobalBytesReceived) & ")") & " | " & Spaces2(27, SizeToString(SessionBytesSent) & " 14(" & SizeToString(GlobalBytesSent) & ")")
  TU vSock, " "
End Sub

