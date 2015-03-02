Attribute VB_Name = "FileSys_NTFS"
',-======================- ==-- -  -
'|   AnGeL - FileSys - NTFS
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Const GENERIC_ALL              As Long = &H10000000

Private Const COMPRESSION_FORMAT_NONE     As Long = 0&
Private Const COMPRESSION_FORMAT_DEFAULT  As Long = 1&
Private Const FILE_DEVICE_FILE_SYSTEM     As Long = &H9&
Private Const METHOD_BUFFERED             As Long = 0&
Private Const FILE_READ_DATA              As Long = &H1&
Private Const FILE_WRITE_DATA             As Long = &H2&
Private Const FILE_ANY_ACCESS             As Long = 0&
Private Const FILE_IS_ENCRYPTED        As Long = 1&
Private Const FILE_SHARE_READ          As Long = &H1&
Private Const FILE_SHARE_WRITE         As Long = &H2&
Private Const OPEN_EXISTING            As Long = 3&
Private Const FILE_FLAG_BACKUP_SEMANTICS  As Long = &H2000000
Private Const FILE_ATTRIBUTE_COMPRESSED   As Long = &H800




Private FSCTL_GET_COMPRESSION          As Long
Private FSCTL_SET_COMPRESSION          As Long


Sub NTFS_Load()
  FSCTL_GET_COMPRESSION = GetCtlCode(FILE_DEVICE_FILE_SYSTEM, 15, METHOD_BUFFERED, FILE_ANY_ACCESS)
  FSCTL_SET_COMPRESSION = GetCtlCode(FILE_DEVICE_FILE_SYSTEM, 16, METHOD_BUFFERED, FILE_READ_DATA Or FILE_WRITE_DATA)
End Sub


Sub NTFS_Unload()
'
End Sub


Private Function GetCtlCode(ByVal lngDeviceType As Long, ByVal lngFunction As Long, ByVal lngMethod As Long, ByVal lngAccess As Long) As Long
  GetCtlCode = (CLng(lngDeviceType) * (2 ^ 16)) Or (CLng(lngAccess) * (2 ^ 14)) Or (CLng(lngFunction) * (2 ^ 2)) Or lngMethod
End Function

Public Function NTFS_SetCompression(ByVal blnIsCompressed As Boolean, ByVal strFileName As String) As Long
  Dim p_lngRtn                        As Long
  Dim p_lngFileHwnd                   As Long
  Dim p_lngBytesRtn                   As Long

  p_lngFileHwnd = kernel32_CreateFileA(strFileName, GENERIC_ALL, FILE_SHARE_WRITE And FILE_SHARE_READ, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
  If p_lngFileHwnd = -1 Then
    NTFS_SetCompression = p_lngFileHwnd
    Exit Function
  End If

  If blnIsCompressed = False Then
     NTFS_SetCompression = kernel32_DeviceIoControl(p_lngFileHwnd, FSCTL_SET_COMPRESSION, COMPRESSION_FORMAT_DEFAULT, 2&, 0&, 0&, p_lngBytesRtn, 0&)
  ElseIf blnIsCompressed = True Then
     NTFS_SetCompression = kernel32_DeviceIoControl(p_lngFileHwnd, FSCTL_SET_COMPRESSION, COMPRESSION_FORMAT_NONE, 2&, 0&, 0&, p_lngBytesRtn, 0&)
  End If

  p_lngRtn = CloseHandle(p_lngFileHwnd)
End Function

Public Function NTFS_IsCompressed(ByVal strFullPath As String) As Boolean
   Dim p_lngRtn                        As Long
   p_lngRtn = kernel32_GetFileAttributesA(strFullPath)
   If p_lngRtn And FILE_ATTRIBUTE_COMPRESSED Then
      NTFS_IsCompressed = True
   Else
      NTFS_IsCompressed = False
   End If
End Function

Public Function NTFS_IsNTFS(ByVal strFilePath As String) As Boolean
   Dim p_strVolBuffer                  As String
   Dim p_strSystemName                 As String
   Dim p_strVol                        As String
   Dim p_lngSerialNum                  As Long
   Dim p_lngSystemFlags                As Long
   Dim p_lngComponentLen               As Long
   Dim p_lngRtn                        As Long
   
   p_strVolBuffer = String(256, 0)
   p_strSystemName = String(256, 0)
   p_strVol = UCase(Mid(strFilePath, 1, 3))
   p_lngRtn = kernel32_GetVolumeInformationA(p_strVol, p_strVolBuffer, Len(p_strVolBuffer) - 1, p_lngSerialNum, p_lngComponentLen, p_lngSystemFlags, p_strSystemName, Len(p_strSystemName) - 1)
   If p_lngRtn = 0 Then
      NTFS_IsNTFS = False
   Else
      If UCase(Mid(p_strSystemName, 1, 4)) = "NTFS" Then
         NTFS_IsNTFS = True
      Else
         NTFS_IsNTFS = False
      End If
   End If
End Function


Public Sub NTFS_CheckCompress(FilePath As String)
  Dim TheFile As String
  If NTFS_IsNTFS(FilePath) Then
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    'Check Dir
    TheFile = Left(FilePath, Len(FilePath) - 1)
    If NTFS_IsCompressed(TheFile) = False Then
      NTFS_SetCompression False, TheFile
    End If
    'Check Files in Dir
    TheFile = Dir(FilePath & "*.*")
    While TheFile <> ""
      If NTFS_IsCompressed(FilePath & "\" & TheFile) = False Then
        NTFS_SetCompression False, FilePath & "\" & TheFile
      End If
      TheFile = Dir
    Wend
  End If
End Sub

Public Function NTFS_IsEncrypted(ByVal strFilePath As String, Optional ByRef error As String) As Boolean
  Dim p_lngRtn                        As Long
  Dim p_lngStatus                     As Long
  Dim p_strErrMsg                     As String
   
  p_lngRtn = advapi32_FileEncryptionStatusA(strFilePath, p_lngStatus)
  
  If p_lngRtn = 0 Then NTFS_IsEncrypted = False: Exit Function
  
  If p_lngStatus = FILE_IS_ENCRYPTED Then
     NTFS_IsEncrypted = True
  End If
End Function

Public Function NTFS_Encrypt(ByVal strFilePath As String) As Long
   NTFS_Encrypt = advapi32_EncryptFileA(strFilePath)
End Function

Public Function NTFS_Decrypt(ByVal strFilePath As String) As Long
   NTFS_Decrypt = advapi32_DecryptFileA(strFilePath, 0&)
End Function

Function GetErrMSG(ErrorCode As Long) As String
  Dim Flags As Long, Puffer As String
  Dim Retval As Long, Sprache As Long
  
  Flags = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS
  Sprache = LANG_NEUTRAL Or (SUBLANG_DEFAULT * 1024)
  Puffer = Space(512)

  Retval = kernel32_FormatMessageA(Flags, 0&, ErrorCode, Sprache, Puffer, Len(Puffer), 0&)
  
  If Retval = 0 Then
    GetErrMSG = ""
  Else
    GetErrMSG = Left$(Puffer, Retval) & vbCrLf
  End If
End Function

