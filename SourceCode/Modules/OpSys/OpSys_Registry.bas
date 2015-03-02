Attribute VB_Name = "OpSys_Registry"
',-======================- ==-- -  -
'|   AnGeL - OpSys - Registry
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public Function GetRegString(ByVal HomeKey As HKEY_CONSTANTS, ByVal KeyName As String, ByVal ValueName As String) As String
  Dim hKey As Long
  Dim sData As String
  Dim lres As Long
  Dim lDataType As Long
  Dim lDlen As Long
  lres = advapi32_RegOpenKeyA(HomeKey, KeyName, hKey)
  If lres <> 0 Then GetRegString = vbNullString: Exit Function
  sData = String$(64, 32) & Chr$(0)
  lDlen = Len(sData)
  lres = advapi32_RegQueryValueExA(hKey, ValueName, 0, lDataType, sData, lDlen)
  If lres <> 0 Then GetRegString = vbNullString: Exit Function
  If lDataType = REG_SZ Then
    GetRegString = Left$(sData, lDlen - 1)
  Else
    GetRegString = vbNullString
  End If
  lres = advapi32_RegCloseKey(hKey)
End Function

