Attribute VB_Name = "Kernel_Functions"
Option Explicit

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Sub Wait(Count As Long)
  Dim cOld As Currency
  Dim cNew As Currency
  cOld = WinTickCount
  cNew = WinTickCount
  While cOld + Count > cNew
    cNew = WinTickCount
    DoEvents
  Wend
End Sub

Public Function Spaces(Count As Long, Tex As String) As String
  If Count - Len(Tex) >= 0 Then
    Spaces = String(Count - Len(Tex), " ")
  Else
    Spaces = " "
  End If
End Function

Public Function Spaces2(Count As Long, Tex As String) As String
  If Count - Len(Tex) >= 0 Then
    Spaces2 = Tex + String(Count - Len(Tex), " ")
  Else
    Spaces2 = Tex
  End If
End Function

Public Function SpacesC(Count As Long, ByVal Tex As String, Tex2 As String) As String
  If Len(Tex2) <= Count Then
    If Count - Len(Tex) - Len(Tex2) < 0 Then Tex = Left(Tex, Count - Len(Tex2))
    SpacesC = Tex & String(Count - Len(Tex) - Len(Tex2), " ") & Tex2
  Else
    SpacesC = Right(Tex2, Count)
  End If
End Function

Function ModifyList(List As String, Separator As String, SearchFor As String, NewEntry As String) As String
Dim u As Long, ThisEntry As String, NewList As String, FoundIt As Boolean
  If SearchFor <> "" Then
    For u = 1 To ParamXCount(List, Separator)
      ThisEntry = ParamX(List, Separator, u)
      If LCase(Left(ThisEntry, Len(SearchFor))) = LCase(SearchFor) Then
        If NewEntry <> "" Then
          If NewList <> "" Then NewList = NewList + Separator
          NewList = NewList + NewEntry: FoundIt = True
        End If
      Else
        If NewList <> "" Then NewList = NewList + Separator
        NewList = NewList + ThisEntry
      End If
    Next u
    If (NewEntry <> "") And (FoundIt = False) Then
      If NewList <> "" Then NewList = NewList + Separator
      NewList = NewList + NewEntry
    End If
  Else
    NewList = List
    If NewList <> "" Then NewList = NewList + Separator
    NewList = NewList + NewEntry
  End If
  ModifyList = NewList
End Function

Function GetListEntry(List As String, Separator As String, SearchFor As String) As String
Dim u As Long
  For u = 1 To ParamXCount(List, Separator)
    If LCase(Left(ParamX(List, Separator, u), Len(SearchFor))) = LCase(SearchFor) Then GetListEntry = ParamX(List, Separator, u): Exit Function
  Next u
  GetListEntry = ""
End Function

Function AddCheck(csum As Long, newstr As String) As Long
  Dim u As Long, newsum As Long
  newsum = csum
  For u = 1 To Len(newstr)
    newsum = (newsum + (Asc(Mid(newstr, u, 1)) * ((u Mod 50) + 1))) Mod 10000
  Next u
  AddCheck = newsum
End Function

Public Function GetPort(Serv As String, Default As Long) As Long
  Dim TempPort As Long, Dummy As String
  If Left(Serv, 1) = "[" Then
    TempPort = InStr(2, Serv, "]", vbBinaryCompare)
    If TempPort > 0 Then
      Dummy = "x:" & Mid(Serv, TempPort + 1)
    Else
      Dummy = "x:" & Default
    End If
  Else
    Dummy = Serv
  End If
  On Local Error Resume Next
  TempPort = CLng(ParamX(Dummy, ":", 2))
  If Err.Number <> 0 Or (TempPort = 0) Then
    Err.Clear
    TempPort = Default
  Else
    GetPort = TempPort
  End If
End Function

Public Function GetAddr(Serv As String) As String
  Dim Index As Long
  If Left(Serv, 1) = "[" Then
    Index = InStr(2, Serv, "]", vbBinaryCompare)
    If Index > 0 Then
      GetAddr = Mid(Serv, 2, Index - 2)
    Else
      GetAddr = ""
    End If
  Else
    GetAddr = ParamX(Serv, ":", 1)
  End If
End Function

Public Function GetServer(Serv As String) As String
  GetServer = Param(Serv, 1)
End Function

Public Function GetProxy(Serv As String) As String
Dim pa As Long, ps As Long
  pa = InStr(Serv, "|") + 1
  If pa > 1 Then
    ps = InStr(pa, Serv, " ")
    If ps > 0 Then
      GetProxy = Mid(Serv, pa, ps - pa)
    Else
      GetProxy = Mid(Serv, pa)
    End If
  Else
    GetProxy = ""
  End If
End Function

'Converts a string with underscores to a string with spaces (i.e. "This is a test.txt")
Function RemUnderscore(ByVal OldLine As String) As String
  RemUnderscore = Replace(OldLine, "_", " ")
End Function

'Converts a string with spaces to a string with underscores (i.e. "This_is_a_test.txt")
Function MakeUnderscore(ByVal OldLine As String) As String
  MakeUnderscore = Replace(OldLine, " ", "_")
End Function

'Converts a version number to a string
Public Function MakeVersionString(VNum As Long) As String
  Dim VString As String
  VString = Right(CStr(VNum), 8)
  VString = String(8 - Len(VString), "0") + VString
  MakeVersionString = CStr(Val(Mid(VString, 1, 2))) & "." & CStr(Val(Mid(VString, 3, 2))) & "." & CStr(Val(Mid(VString, 5, 2))) & "." & CStr(Val(Mid(VString, 7, 2)))
End Function

'Converts a long num to a dcc hex num
Function MakeHEXNum(Num As Long)
  Dim HexData As String, i As Long, Data(3) As Byte
  HexData = Hex(Num)
  HexData = String(8 - Len(HexData), "0") & HexData
  For i = 1 To Len(HexData) Step 2
    Data((i - 1) / 2) = CLng("&H" & Mid(HexData, i, 2))
  Next
  MakeHEXNum = Chr(Data(0)) + Chr(Data(1)) + Chr(Data(2)) + Chr(Data(3))
End Function

'Converts a dcc hex num to a long num
Function MakeLongNum(HEXNum As String)
  Dim u As Long, HexData As String, Part As String
  For u = 1 To 4
    Part = Hex$(Asc(Mid(HEXNum, u, 1))): If Len(Part) = 1 Then Part = "0" & Part
    HexData = HexData + Part
  Next u
  MakeLongNum = CLng("&H" & HexData)
End Function

'Converts decimal to base64
Public Function LongToBase64(Number As Long) As String
  Dim u As Long, HadNonZero As Boolean, Divi As Integer, LongNum As Long, Base64 As String
  If Number > 1073741823 Then LongToBase64 = "A": Exit Function
  LongNum = Number
  For u = 4 To 0 Step -1
    Divi = LongNum \ (64 ^ u): If Divi <> 0 Then HadNonZero = True
    If (Divi = 0 And HadNonZero) Or Divi <> 0 Or u = 0 Then Base64 = Base64 + IntToBase64Char(Divi)
    LongNum = LongNum - (Divi * (64 ^ u))
  Next u
  LongToBase64 = Base64
End Function

'Converts base64 to decimal
Public Function Base64ToLong(Base64 As String) As Long ' : AddStack "Conversions_Base64ToLong(" & Base64 & ")"
Dim u As Long, LongNum As Long
  If Len(Base64) > 5 Then Base64ToLong = 0: Exit Function
  For u = 1 To Len(Base64)
    LongNum = LongNum + (64 ^ (Len(Base64) - u)) * Base64CharToInt(Mid(Base64, u, 1))
  Next u
  Base64ToLong = LongNum
End Function

Public Function IntToBase64Char(Number As Integer) As String ' : AddStack "Conversions_IntToBase64Char(" & Number & ")"
  Select Case Number
    Case 0 To 25: IntToBase64Char = Chr(65 + Number)
    Case 26 To 51: IntToBase64Char = Chr(71 + Number)
    Case 52 To 61: IntToBase64Char = Chr(Number - 4)
    Case 62: IntToBase64Char = "["
    Case 63: IntToBase64Char = "]"
  End Select
End Function

Public Function Base64CharToInt(Char As String) As Integer ' : AddStack "Conversions_Base64CharToInt(" & Char & ")"
  Select Case Char
    Case "A" To "Z": Base64CharToInt = Asc(Char) - 65
    Case "a" To "z": Base64CharToInt = Asc(Char) - 71
    Case "0" To "9": Base64CharToInt = Asc(Char) + 4
    Case "[": Base64CharToInt = 62
    Case "]": Base64CharToInt = 63
  End Select
End Function

Public Function RemSpaces(Line As String) As String
  RemSpaces = Replace(Line, " ", "·")
End Function

Public Function AddSpaces(Line As String) As String
  AddSpaces = Replace(Line, "·", " ")
End Function

Public Function UnsignedToInteger(Value As Long) As Integer
  If Value < 0 Or Value >= OFFSET_2 Then UnsignedToInteger = 0: Exit Function
  If Value <= MAXINT_2 Then
    UnsignedToInteger = Value
  Else
    UnsignedToInteger = Value - OFFSET_2
  End If
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
  If Value < 0 Then
    IntegerToUnsigned = Value + OFFSET_2
  Else
    IntegerToUnsigned = Value
  End If
End Function

Function Param(Text As String, Num As Long) As String
  Param = ParamX(Text, " ", Num)
End Function

Function ParamCount(Text As String) As Long
  ParamCount = ParamXCount(Text, " ")
End Function

Function GetRest(Text As String, Num As Long) As String
  GetRest = GetRestX(Text, " ", Num)
End Function

Function ParamX(Line As String, Seperator As String, Number As Long) As String
  ParamX = ""
  
  If Number <= 0 Then Exit Function
  
  Dim Parms() As String
  Parms = ParamXArr(Line, Seperator)
  
  If Number > UBound(Parms) Then Exit Function
  
  ParamX = Parms(Number)
End Function

Function ParamXCount(Line As String, Seperator As String) As Long
  Dim Parms() As String
  Parms = ParamXArr(Line, Seperator)
  If Parms(1) = "" Then ParamXCount = 0 Else ParamXCount = UBound(Parms)
End Function

Function GetRestX(sText As String, sSepChar As String, iNum As Long) As String
  If iNum <= 0 Or iNum > ParamXCount(sText, sSepChar) Then GetRestX = "" Else If iNum = 1 Then GetRestX = sText Else GetRestX = Mid(Replace(sText, sSepChar, "", , iNum - 2, vbBinaryCompare), InStr(1, Replace(sText, sSepChar, "", , iNum - 2, vbBinaryCompare), sSepChar, vbBinaryCompare) + 1 + (Len(sSepChar) - 1))
End Function

Function ParamArr(Text As String) As Variant
  ParamArr = ParamXArr(Text, " ")
End Function

Function ParamXArr(Line As String, Seperator As String) As Variant
  Static LastLine As String, LastSeperator As String, LastParms() As String
  Dim Parms() As String, Dummy As String
  
  If Line = LastLine And Seperator = LastSeperator Then
    ParamXArr = LastParms
  Else
    Dummy = Line
    While InStr(Dummy, Seperator & Seperator) > 0
      Dummy = Replace(Dummy, Seperator & Seperator, Seperator)
    Wend
    If Left(Dummy, Len(Seperator)) = Seperator Then Dummy = Mid(Dummy, Len(Seperator) + 1)
    Parms = Split(Seperator & Dummy, Seperator)
    Parms(0) = Seperator
    ParamXArr = Parms
  End If
End Function

'Checks whether a command is "weak" -> too easy to guess
Function WeakPass(Line As String, Nick As String) As Boolean ' : AddStack "Routines_WeakPass(" & Line & ", " & Nick & ")"
  If (LCase(Line) = LCase(Nick)) Or (LCase(Line) = LCase(MyNick)) Or (LCase(Line) = LCase(BotNetNick)) Then WeakPass = True: Exit Function
  Select Case LCase(Line)
    Case "123456", "654321", "password", "passwort", "qwertz", "qwerty", "hallo!", "blabla", "abcdef", "asdfgh", "yxcvbn", "angelbot", "test123", "ichbins"
      WeakPass = True: Exit Function
  End Select
  WeakPass = False
End Function


'Encrypts a password
Function EncryptIt(Old As String) As String ' : AddStack "Routines_EncryptIt(" & Old & ")"
Dim u As Long, Crypt As String, Final As String, ChkSum As Single, AddLen As Single, Key As String
  For u = 1 To Len(Old)
    ChkSum = ChkSum + Sin(Asc(Mid(Old, u, 1)) * (u * 2.926))
  Next u
  For u = 1 To Len(Old)
    AddLen = AddLen + Cos(Asc(Mid(Old, u, 1)) * ((Len(Old) - u) * 61.251))
  Next u
  For u = 1 To Len(Old)
    Key = Key + Chr(Abs((Asc(Mid(Old, ((u - 1) * 6.29 Mod Len(Old)) + 1, 1)) ^ 1.4) Mod 255)) + Chr(Abs((Asc(Mid(Old, ((u - 1) * 1.88 Mod Len(Old)) + 1, 1)) * 4.7) Mod 255))
  Next u
  ChkSum = Abs(ChkSum) + Len(Old)
  For u = 1 To Len(Old) + Int(Abs(AddLen) + Abs(ChkSum))
    Crypt = Crypt + Chr(33 + Abs(Cos(Asc(Mid(Old, ((u - 1) Mod Len(Old)) + 1, 1))) * 59.4 - ChkSum * 3.2 + IIf(u Mod 2 = 0, -1, 1) * (Asc(Mid(Old, ((u - 1) Mod Len(Old)) + 1, 1)) + 17.3) * (u * 4.53)) Mod 94)
  Next u
  If Len(Old) + Int(Abs(AddLen)) > 60 Then
    Final = Space(Len(Old) \ 6 + Int(Abs(AddLen)) + 6)
  ElseIf Len(Old) + Int(Abs(AddLen)) > 25 Then
    Final = Space(Len(Old) + Int(Abs(AddLen)) - 12)
  ElseIf Len(Old) + Int(Abs(AddLen)) < 8 Then
    Final = Space(Len(Old) + Int(Abs(AddLen)) + 5)
  Else
    Final = Space(Len(Old) + Int(Abs(AddLen)))
  End If
  For u = 1 To Len(Crypt)
    Mid(Final, ((u - 1) Mod Len(Final)) + 1, 1) = Chr(33 + Abs((Asc(Mid(Crypt, u, 1)) + Asc(Mid(Key, ((u - 1) Mod Len(Key)) + 1, 1))) Mod 94))
  Next u
  EncryptIt = "¤" & Replace(Replace(MakeININick(Final), "=", "*"), ";", "_")
End Function

'Encrypts a password
Function EncryptIt2(Old As String) As String ' : AddStack "Routines_EncryptIt(" & Old & ")"
  Dim u As Long, Crypt As String, Final As String, ChkSum As Single, AddLen As Single, Key As String
  For u = 1 To Len(Old)
    ChkSum = ChkSum + Sin(Asc(Mid(Old, u, 1)) * (u * 2.926))
  Next u
  For u = 1 To Len(Old)
    AddLen = AddLen + Cos(Asc(Mid(Old, u, 1)) * ((Len(Old) - u) * 61.251))
  Next u
  For u = 1 To Len(Old)
    Key = Key + Chr(Abs((Asc(Mid(Old, ((u - 1) * 6.29 Mod Len(Old)) + 1, 1)) ^ 1.4) Mod 255)) + Chr(Abs((Asc(Mid(Old, ((u - 1) * 1.88 Mod Len(Old)) + 1, 1)) * 4.7) Mod 255))
  Next u
  ChkSum = Abs(ChkSum) + Len(Old)
  For u = 1 To Len(Old) + Int(Abs(AddLen) + Abs(ChkSum))
    Crypt = Crypt + Chr(33 + Abs(Cos(Asc(Mid(Old, ((u - 1) Mod Len(Old)) + 1, 1))) * 59.4 - ChkSum * 3.2 + IIf(u Mod 2 = 0, -1, 1) * (Asc(Mid(Old, ((u - 1) Mod Len(Old)) + 1, 1)) + 17.3) * (u * 4.53)) Mod 94)
  Next u
  If Len(Old) + Int(Abs(AddLen)) > 60 Then
    Final = Space(Len(Old) \ 6 + Int(Abs(AddLen)) + 6)
  ElseIf Len(Old) + Int(Abs(AddLen)) > 25 Then
    Final = Space(Len(Old) + Int(Abs(AddLen)) - 12)
  ElseIf Len(Old) + Int(Abs(AddLen)) < 8 Then
    Final = Space(Len(Old) + Int(Abs(AddLen)) + 5)
  Else
    Final = Space(Len(Old) + Int(Abs(AddLen)))
  End If
  For u = 1 To Len(Crypt)
    Mid(Final, ((u - 1) Mod Len(Final)) + 1, 1) = Chr(33 + Abs((Asc(Mid(Crypt, u, 1)) + Asc(Mid(Key, ((u - 1) Mod Len(Key)) + 1, 1))) Mod 94))
  Next u
  EncryptIt2 = "¤" & Replace(Replace(Final, "=", "*"), ";", "_")
End Function

Public Function CLongToULong(LongValue As Long) As Variant
    Select Case LongValue
        Case Is < 0
            CLongToULong = CDec(LongValue - CDec(&H80000000) * 2)
        Case &H80000000
            CLongToULong = CDec(LongValue) * -1
        Case Else
            CLongToULong = CDec(LongValue)
    End Select
End Function

Function MakeSettingText(Text As String, Length As Long) As String '' : AddStack "Setups_MakeSettingText(" & Length & ", " & Text & ")"
  Dim i As Long
  If Len(Text) >= Length Then
    MakeSettingText = Text
    Exit Function
  Else
    MakeSettingText = Text & "14"
    For i = Len(Text) To Length
      MakeSettingText = MakeSettingText & "."
    Next
    MakeSettingText = MakeSettingText & ": "
  End If
End Function
Function MakeLength(Text As String, Length As Long) As String '' : AddStack "Setups_MakeSettingText(" & Length & ", " & Text & ")"
  Dim i As Long
  If Len(Text) > Length Then
    MakeLength = Mid(Text, 1, Len(Text) - 3) & "..."
    Exit Function
  ElseIf Len(Text) = Length Then
    MakeLength = Text
  Else
    MakeLength = Text
    For i = Len(Text) To Length
      MakeLength = MakeLength & " "
    Next
    MakeLength = Mid(MakeLength, 1, Len(MakeLength) - 1)
  End If
End Function

Function Switch(Text As String) As Boolean
  Select Case LCase(Text)
    Case "0", "no", "off", "false", "nein", "aus", "falsch"
      Switch = False
    Case "1", "yes", "on", "true", "ja", "an", "wahr"
      Switch = True
  End Select
End Function

Function hexdump(Data As String) As String
  Dim Index As Integer
  Dim Dummy As String
  Dim dummy2 As String
  For Index = 1 To Len(Data)
    Dummy = Hex(Asc(Mid(Data, Index, 1)))
    If Len(Dummy) = 1 Then Dummy = "0" & Dummy
    dummy2 = dummy2 & Dummy & " "
  Next Index
  hexdump = dummy2
End Function

Public Function EncryptString(OString As String, Key As String) As String
Dim u As Long, QSum As Long, SVal As Long, KVal As Long, KeyPos As Long
Dim NVal As Long, TmpStr As String, HStr As String, Line As String
  If Key = "" Then EncryptString = OString: Exit Function
  Line = OString
  QSum = Len(Key)
  For u = 1 To Len(Key)
    QSum = QSum + Asc(Mid(Key, u, 1))
  Next u
  KeyPos = QSum Mod Len(Key)
  For u = 1 To Len(Line)
    SVal = Asc(Mid(Line, u, 1))
    KeyPos = KeyPos + 1: If KeyPos > Len(Key) Then KeyPos = 1
    KVal = Asc(Mid(Key, KeyPos, 1))
    NVal = (SVal + (KVal * Int(u * QSum / 5))) Mod 256
    NVal = NVal Xor Asc(Mid(Key, Len(Key) - KeyPos + 1, 1))
    TmpStr = TmpStr + Chr(NVal)
  Next u
  Line = TmpStr: TmpStr = ""
  KeyPos = 0
  For u = 1 To Len(Line)
    SVal = Asc(Mid(Line, u, 1))
    KeyPos = KeyPos + 1: If KeyPos > Len(Key) Then KeyPos = 1
    KVal = Asc(Mid(Key, KeyPos, 1))
    NVal = (SVal + (KVal * Int(u * QSum))) Mod 256
    NVal = NVal Xor Asc(Mid(Key, Len(Key) - KeyPos + 1, 1))
    HStr = Hex(NVal): If Len(HStr) < 2 Then HStr = "0" & HStr
    TmpStr = TmpStr + HStr
  Next u
  EncryptString = TmpStr
End Function

Public Function DecryptString(OString As String, Key As String) As String
Dim u As Long, QSum As Long, SVal As Long, KVal As Long, KeyPos As Long
Dim NVal As Long, TmpStr As String, HStr As String, Line As String
  If Key = "" Then DecryptString = OString: Exit Function
  QSum = Len(Key)
  For u = 1 To Len(Key)
    QSum = QSum + Asc(Mid(Key, u, 1))
  Next u
  For u = 1 To Len(OString) Step 2
    If IsNumeric("&H" & Mid(OString, u, 2)) Then TmpStr = TmpStr + Chr(CLng("&H" & Mid(OString, u, 2)))
  Next u
  Line = TmpStr: TmpStr = ""
  KeyPos = 0
  For u = 1 To Len(Line)
    SVal = Asc(Mid(Line, u, 1))
    KeyPos = KeyPos + 1: If KeyPos > Len(Key) Then KeyPos = 1
    KVal = Asc(Mid(Key, KeyPos, 1))
    SVal = SVal Xor Asc(Mid(Key, Len(Key) - KeyPos + 1, 1))
    NVal = (SVal - (KVal * Int(u * QSum))) Mod 256
    If NVal < 0 Then NVal = 256 - Abs(NVal)
    TmpStr = TmpStr + Chr(NVal)
  Next u
  Line = TmpStr: TmpStr = ""
  KeyPos = QSum Mod Len(Key)
  For u = 1 To Len(Line)
    SVal = Asc(Mid(Line, u, 1))
    KeyPos = KeyPos + 1: If KeyPos > Len(Key) Then KeyPos = 1
    KVal = Asc(Mid(Key, KeyPos, 1))
    SVal = SVal Xor Asc(Mid(Key, Len(Key) - KeyPos + 1, 1))
    NVal = (SVal - (KVal * Int(u * QSum / 5))) Mod 256
    If NVal < 0 Then NVal = 256 - Abs(NVal)
    TmpStr = TmpStr + Chr(NVal)
  Next u
  DecryptString = TmpStr
End Function

