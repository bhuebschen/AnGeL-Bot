Attribute VB_Name = "OpSys_WinInfo"
',-======================- ==-- -  -
'|   AnGeL - OpSys - WinInfo
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


' -= Informationen über Windows
Public WinIdentifier As String
Public WinVersionName As String
Public WinVersionMajor As Byte
Public WinVersionMinor As Byte
Public WinNTOS As Boolean


Sub GetWinInfo()
  Dim VersionInfo As OSVERSIONINFO
  Dim Result As Long
  Dim Index As Integer
  Dim Dummy As String
  
  VersionInfo.dwOSVersionInfoSize = 148
  Result = kernel32_GetVersionExA(VersionInfo)
  Index = InStr(1, VersionInfo.szCSDVersion, Chr(0), vbBinaryCompare)
  If Index = 0 Then
    Dummy = 0
  Else
    Dummy = Left(VersionInfo.szCSDVersion, Index - 1)
  End If
  
  If VersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    WinNTOS = False
    Select Case VersionInfo.dwMinorVersion
      Case 0
        WinIdentifier = "WIN95"
        WinVersionName = "Windows 95"
      Case 10
        WinIdentifier = "WIN98"
        WinVersionName = "Windows 98"
      Case 90
        WinIdentifier = "WINME"
        WinVersionName = "Windows Millenium Edition"
      Case Else
        WinIdentifier = "WIN"
        WinVersionName = "Windows"
    End Select
  ElseIf VersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
    WinNTOS = True
    If VersionInfo.dwMajorVersion = 5 And VersionInfo.dwMinorVersion = 0 Then
      WinIdentifier = "WIN2K"
      WinVersionName = "Windows 2000"
    ElseIf VersionInfo.dwMajorVersion = 5 And VersionInfo.dwMinorVersion = 1 Then
      WinIdentifier = "WINXP"
      WinVersionName = "Windows XP"
    Else
      WinIdentifier = "WINNT"
      WinVersionName = "Windows NT"
    End If
  Else
    WinNTOS = False
    WinIdentifier = "UNIX"
    WinVersionName = "UnknownOS"
  End If
  WinVersionMinor = VersionInfo.dwMinorVersion
  WinVersionMajor = VersionInfo.dwMajorVersion
End Sub

Public Function WinTickCount() As Variant
  Static OldTickCount As Currency, ResetCounter As Currency
  Dim LongValue As Long, NewTickCount As Currency
  
  On Local Error Resume Next
  LongValue = kernel32_GetTickCount
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
  
  Select Case LongValue
    Case Is < 0
      NewTickCount = CDec(LongValue - CDec(&H80000000) * 2)
    Case &H80000000
      NewTickCount = CDec(LongValue) * -1
    Case Else
      NewTickCount = CDec(LongValue)
  End Select
  If ResetCounter > 0 Then
    NewTickCount = NewTickCount + (4294967296# * ResetCounter)
  Else
  End If
  If NewTickCount < OldTickCount Then
    ResetCounter = ResetCounter + 1
    NewTickCount = NewTickCount + (4294967296#)
    OldTickCount = NewTickCount
    WinTickCount = OldTickCount
  Else
    OldTickCount = NewTickCount
    WinTickCount = OldTickCount
  End If
End Function
