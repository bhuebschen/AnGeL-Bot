Attribute VB_Name = "FileSys_FileInfo"
Option Explicit


Private Type FILETIME
  lLowDateTime As Long
  lHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * 260
  cAlternate As String * 16
End Type

Function GetFileDescription(FileName As String) As String ' : AddStack "FileInfo_GetFileDescription(" & FileName & ")"
  On Error GoTo ErrInFileDesc
  Dim lpAdress As Long, lpSize As Long
  Dim dwSize As Long, VersionData As String
  
  dwSize = version_GetFileVersionInfoSizeA(FileName, 0&)
  
  If dwSize Then
    VersionData = String(dwSize, 0)
    If version_GetFileVersionInfoA(FileName, ByVal 0&, dwSize, ByVal VersionData) Then
      
      'Sprachdaten auslesen
      If version_VerQueryValueA(ByVal VersionData, ByVal "\VarFileInfo\Translation", lpAdress, lpSize) Then
        Dim i As Integer, lang As String, s As String, h As String
        ReDim Languages(2, lpSize / 2) As Integer
        
        'SprachCodes in ein Array kopieren
        kernel32_RtlMoveMemory Languages(0, 0), ByVal lpAdress, lpSize
        
        'Daten für jede unterstützte Sprache auslesen
        For i = 0 To (lpSize / 4) - 1
          s = String(255, 0)
          s = Left(s, kernel32_VerLanguageNameA(Languages(0, i), s, Len(s)))
          
          lang = "\StringFileInfo\"
          h = Hex(Languages(0, i)): h = String(4 - Len(h), "0") & h
          lang = lang & h
          h = Hex(Languages(1, i)): h = String(4 - Len(h), "0") & h
          lang = lang & h & "\"
                    
          'FileDescription
          If version_VerQueryValueA(ByVal VersionData, ByVal lang & "FileDescription", lpAdress, lpSize) Then
            s = String(lpSize, 0)
            kernel32_RtlMoveMemory ByVal s, ByVal lpAdress, lpSize
            GetFileDescription = s
            Exit Function
          End If
        Next i
      End If
    End If
  End If
  Exit Function

ErrInFileDesc:
  Err.Clear
  'Stop
'  SpreadFlagMessage 0, "+m", "14[" & Time & "] *** ERROR in GetFileDescription! (" & Err.Number & ")"
End Function
