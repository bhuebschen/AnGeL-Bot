Attribute VB_Name = "TimedEvents_Functions"
Option Explicit

Public Function FindUnbanTime() As Currency
  On Local Error Resume Next
  Dim Num1 As Currency, Num2 As Currency
  Num1 = WinTickCount
  Num2 = CCur(Rnd * 240000) + 600000
  FindUnbanTime = CCur(CCur(Num1) + CCur(Num2))
  If Err.Number > 0 Then
    FindUnbanTime = 0
    Err.Clear
  End If
End Function

Public Function A_after_B(a As String, b As String) As Boolean
  Dim DateA As Date, DateB As Date
  If (a = "online!") And (b <> "online!") Then
    A_after_B = True
  ElseIf (a <> "online!") And (b = "online!") Then
    A_after_B = False
  Else
    DateA = GetDate(a)
    DateB = GetDate(b)
    A_after_B = (DateA > DateB)
  End If
End Function

Public Function TimeSpan(LastTime As String) As String
  Dim Last As Date, Days As Long, Hours As Long, Minutes As Long, Seconds As Long
  Dim StrDays As String, StrHours As String, StrMinutes As String, StrSeconds As String
  Dim Message As String, ErrLine As Long
  If IsDate(LastTime) Then
    Last = CDate(LastTime)
  Else
    On Local Error Resume Next
    Last = CDate(Val(Replace(LastTime, ",", ".")))
    If Err.Number > 0 Then
      TimeSpan = "some time ago."
      Err.Clear
      Exit Function
    End If
  End If
  Days = Int(CDate(Now) - CDate(Last))
  Hours = Hour(CDate(Now) - CDate(Last))
  Minutes = Minute(CDate(Now) - CDate(Last))
  If Days > 0 Then
    If Days = 1 Then
      StrDays = "1 day, "
    Else
      StrDays = Trim(Str(Days)) & " days, "
    End If
  End If
  If Hours > 0 Then
    If Hours = 1 Then
      StrHours = "1 hour, "
    Else
      StrHours = Trim(Str(Hours)) & " hours, "
    End If
  End If
  If Minutes > 0 Then
    If Minutes = 1 Then
      StrMinutes = "1 minute, "
    Else
      StrMinutes = Trim(Str(Minutes)) & " minutes, "
    End If
  End If
  Message = StrDays & StrHours & StrMinutes
  If Message <> "" Then
    Message = Left(Message, Len(Message) - 2) & " ago."
  Else
    Message = "just moments ago!!!"
  End If
  TimeSpan = Message
End Function

Public Function GetSecondsTillNow(FromTime As Date) As Long
  GetSecondsTillNow = DateDiff("s", FromTime, Now)
End Function

Public Function GetSecondsTillNow2(FromTime As Currency) As Currency
  GetSecondsTillNow2 = (WinTickCount - FromTime) \ 1000
End Function

Public Function GetSecondsTill(AtTime As Date) As Long
  GetSecondsTill = DateDiff("s", Now, AtTime)
End Function

Public Function TimeSpan2(Last As Date) As String
  Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long
  Dim StrDays As String, StrHours As String, StrMinutes As String, StrSeconds As String
  Dim Message As String, ErrLine As Long
  Days = Int(CDate(Now) - CDate(Last))
  Hours = Hour(CDate(Now) - CDate(Last))
  Minutes = Minute(CDate(Now) - CDate(Last))
  If Days > 0 Then
    If Days = 1 Then StrDays = "1d " Else StrDays = Trim(Str(Days)) & "d "
  Else
    StrDays = "   "
  End If
  If Hours > 0 Then
    If Hours = 1 Then StrHours = " 1h " Else StrHours = IIf(Len(Trim(Str(Hours))) = 1, " ", "") + Trim(Str(Hours)) & "h "
  Else
    If StrDays = "   " Then StrHours = "    " Else StrHours = " 0h "
  End If
  If Minutes > 0 Then
    If Minutes = 1 Then StrMinutes = " 1m " Else StrMinutes = IIf(Len(Trim(Str(Minutes))) = 1, " ", "") + Trim(Str(Minutes)) & "m "
  Else
    StrMinutes = " 0m "
  End If
  Message = StrDays + StrHours + StrMinutes
  If Message <> "" Then Message = Left(Message, Len(Message) - 1) Else Message = "0m"
  TimeSpan2 = Message
End Function

Public Function TimeSpan3(LastS As String) As String
  Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long
  Dim StrDays As String, StrHours As String, StrMinutes As String, StrSeconds As String
  Dim Message As String, ErrLine As Long, Last As Date
  Last = CDate(Val(LastS))
  Days = Int(CDate(Now) - CDate(Last))
  Hours = Hour(CDate(Now) - CDate(Last))
  Minutes = Minute(CDate(Now) - CDate(Last))
  If Days > 0 Then
    If Days = 1 Then StrDays = "1d " Else StrDays = Trim(Str(Days)) & "d "
  Else
    StrDays = ""
  End If
  If Hours > 0 Then
    StrHours = Trim(Str(Hours)) & "h "
  Else
    If StrDays = "" Then StrHours = "" Else StrHours = "0h "
  End If
  If Minutes > 0 Then
    StrMinutes = Trim(Str(Minutes)) & "m"
  Else
    StrMinutes = "0m"
  End If
  Message = StrDays + StrHours + StrMinutes
  If Message <> "" Then Message = Trim(Message) Else Message = "0m"
  TimeSpan3 = Message
End Function

Public Function TimeSince(LastTime As String) As String
  Dim Last As Date, Days As Long, Hours As Long, Minutes As Long, Seconds As Long
  Dim StrDays As String, StrHours As String, StrMinutes As String, StrSeconds As String
  Dim Message  As String
  Last = CDate(LastTime)
  Days = Int(CDate(Now) - CDate(Last))
  Hours = Hour(CDate(Now) - CDate(Last))
  Minutes = Minute(CDate(Now) - CDate(Last))
  If Days > 0 Then If Days = 1 Then StrDays = "1 day, " Else StrDays = Trim(Str(Days)) & " days, "
  If Hours > 0 Then If Hours = 1 Then StrHours = "1 hour, " Else StrHours = Trim(Str(Hours)) & " hours, "
  If Minutes > 0 Then If Minutes = 1 Then StrMinutes = "1 minute, " Else StrMinutes = Trim(Str(Minutes)) & " minutes, "
  Message = StrDays + StrHours + StrMinutes
  If Message <> "" Then Message = Left(Message, Len(Message) - 2) Else Message = "a few seconds..."
  TimeSince = Message
End Function


'Converts a date number or string to a date
Function GetDate(StrDate As String) As Date
  If IsDate(StrDate) Then
    GetDate = CDate(StrDate)
    Exit Function
  Else
    On Local Error Resume Next
    GetDate = CDate(Val(StrDate))
    If Err.Number > 0 Then
      Err.Clear
      GetDate = DateSerial(1998, 1, 1)
      Exit Function
    End If
  End If
End Function

