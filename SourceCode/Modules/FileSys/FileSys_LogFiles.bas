Attribute VB_Name = "FileSys_LogFiles"
Option Explicit


Public LogMaxAge As Integer


Function PutLog(Line As String)
  Dim FileNum As Integer
  On Local Error Resume Next
  FileNum = FreeFile
  Open FileAreaHome & "Logs\" & Format(Now, "yy-mm-dd.log") For Append As #FileNum
    If DebugMode Then Trace "[" & Time & "] " & Line
    Print #FileNum, "[" & Time & "] " & Line
  Close #FileNum
End Function


'Converts a log filename to a date
Public Function LogToDate(LogName As String) As Date
Dim Rest As String
  If Not (LCase(LogName) Like "*-*-*.log") Then LogToDate = DateSerial(1998, 1, 1): Exit Function
  Rest = DateSerial(Mid(LogName, 1, 2), Mid(LogName, 4, 2), Mid(LogName, 7, 2))
  If IsDate(Rest) Then LogToDate = CDate(Rest): Exit Function
  LogToDate = DateSerial(1998, 1, 1)
End Function

