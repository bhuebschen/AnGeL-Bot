Attribute VB_Name = "Scripting_AddinSecurity"
',-——————————————————————- ——-- -  -
'|   AnGeL AddIn Security Module
'|   © 2003 by Benedikt Hübschen
'|-—————————————- -——- ——- -- -
'|
'| 19.02.2003 - Benedikt Hübschen
'|   SHA256 Hashing
'|
'`-—————————————————————-- -———- ——- -- -

Private m_lOnBits(30) As Long
Private m_l2Power(30) As Long
Private K(63) As Long

Private Const BITS_TO_A_BYTE As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE

Sub InitHashtable()
  m_lOnBits(0) = 1:  m_lOnBits(1) = 3:  m_lOnBits(2) = 7:  m_lOnBits(3) = 15:  m_lOnBits(4) = 31
  m_lOnBits(5) = 63:  m_lOnBits(6) = 127:  m_lOnBits(7) = 255:  m_lOnBits(8) = 511:  m_lOnBits(9) = 1023
  m_lOnBits(10) = 2047:  m_lOnBits(11) = 4095:  m_lOnBits(12) = 8191:  m_lOnBits(13) = 16383:  m_lOnBits(14) = 32767
  m_lOnBits(15) = 65535:  m_lOnBits(16) = 131071:  m_lOnBits(17) = 262143:  m_lOnBits(18) = 524287:  m_lOnBits(19) = 1048575
  m_lOnBits(20) = 2097151:  m_lOnBits(21) = 4194303:  m_lOnBits(22) = 8388607:  m_lOnBits(23) = 16777215:  m_lOnBits(24) = 33554431
  m_lOnBits(25) = 67108863:  m_lOnBits(26) = 134217727:  m_lOnBits(27) = 268435455:  m_lOnBits(28) = 536870911:  m_lOnBits(29) = 1073741823
  m_lOnBits(30) = 2147483647
  m_l2Power(0) = 1:  m_l2Power(1) = 2:  m_l2Power(2) = 4:  m_l2Power(3) = 8:  m_l2Power(4) = 16
  m_l2Power(5) = 32:  m_l2Power(6) = 64:  m_l2Power(7) = 128:  m_l2Power(8) = 256:  m_l2Power(9) = 512
  m_l2Power(10) = 1024:  m_l2Power(11) = 2048:  m_l2Power(12) = 4096:  m_l2Power(13) = 8192:  m_l2Power(14) = 16384
  m_l2Power(15) = 32768:  m_l2Power(16) = 65536:  m_l2Power(17) = 131072:  m_l2Power(18) = 262144:  m_l2Power(19) = 524288
  m_l2Power(20) = 1048576:  m_l2Power(21) = 2097152:  m_l2Power(22) = 4194304:  m_l2Power(23) = 8388608:  m_l2Power(24) = 16777216
  m_l2Power(25) = 33554432:  m_l2Power(26) = 67108864:  m_l2Power(27) = 134217728:  m_l2Power(28) = 268435456:  m_l2Power(29) = 536870912
  m_l2Power(30) = 1073741824
  K(0) = &H428A2F98:  K(1) = &H71374491:  K(2) = &HB5C0FBCF:  K(3) = &HE9B5DBA5:  K(4) = &H3956C25B
  K(5) = &H59F111F1:  K(6) = &H923F82A4:  K(7) = &HAB1C5ED5:  K(8) = &HD807AA98:  K(9) = &H12835B01
  K(10) = &H243185BE:  K(11) = &H550C7DC3:  K(12) = &H72BE5D74:  K(13) = &H80DEB1FE:  K(14) = &H9BDC06A7
  K(15) = &HC19BF174:  K(16) = &HE49B69C1:  K(17) = &HEFBE4786:  K(18) = &HFC19DC6:  K(19) = &H240CA1CC
  K(20) = &H2DE92C6F:  K(21) = &H4A7484AA:  K(22) = &H5CB0A9DC:  K(23) = &H76F988DA:  K(24) = &H983E5152
  K(25) = &HA831C66D:  K(26) = &HB00327C8:  K(27) = &HBF597FC7:  K(28) = &HC6E00BF3:  K(29) = &HD5A79147
  K(30) = &H6CA6351:  K(31) = &H14292967:  K(32) = &H27B70A85:  K(33) = &H2E1B2138:  K(34) = &H4D2C6DFC
  K(35) = &H53380D13:  K(36) = &H650A7354:  K(37) = &H766A0ABB:  K(38) = &H81C2C92E:  K(39) = &H92722C85
  K(40) = &HA2BFE8A1:  K(41) = &HA81A664B:  K(42) = &HC24B8B70:  K(43) = &HC76C51A3:  K(44) = &HD192E819
  K(45) = &HD6990624:  K(46) = &HF40E3585:  K(47) = &H106AA070:  K(48) = &H19A4C116:  K(49) = &H1E376C08
  K(50) = &H2748774C:  K(51) = &H34B0BCB5:  K(52) = &H391C0CB3:  K(53) = &H4ED8AA4A:  K(54) = &H5B9CCA4F
  K(55) = &H682E6FF3:  K(56) = &H748F82EE:  K(57) = &H78A5636F:  K(58) = &H84C87814:  K(59) = &H8CC70208
  K(60) = &H90BEFFFA:  K(61) = &HA4506CEB:  K(62) = &HBEF9A3F7:  K(63) = &HC67178F2
End Sub

Public Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
  Dim lX4 As Long, lY4 As Long, lY8 As Long, lX8 As Long, lResult As Long
  lX8 = lX And &H80000000
  lY8 = lY And &H80000000
  lX4 = lX And &H40000000
  lY4 = lY And &H40000000
  lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
  If lX4 And lY4 Then
    lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
  ElseIf lX4 Or lY4 Then
    If lResult And &H40000000 Then
      lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
    Else
      lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
    End If
  Else
    lResult = lResult Xor lX8 Xor lY8
  End If
  AddUnsigned = lResult
End Function
Public Function UnsignedDel(Data1 As Long, Data2 As Long) As Long

  Dim x1(0 To 3) As Byte
  Dim x2(0 To 3) As Byte
  Dim xx(0 To 3) As Byte
  Dim Rest As Long
  Dim Value As Long
  Dim a As Long
  
  Call kernel32_RtlMoveMemory(x1(0), Data1, 4)
  Call kernel32_RtlMoveMemory(x2(0), Data2, 4)
  Call kernel32_RtlMoveMemory(xx(0), UnsignedDel, 4)
  
  For a = 0 To 3
    Value = CLng(x1(a)) - CLng(x2(a)) - Rest
    If (Value < 0) Then
      Value = Value + 256
      Rest = 1
    Else
      Rest = 0
    End If
    xx(a) = Value
  Next
  
  Call kernel32_RtlMoveMemory(UnsignedDel, xx(0), 4)

End Function


Private Function ch(ByVal x As Long, ByVal Y As Long, ByVal Z As Long) As Long
   ch = ((x And Y) Xor ((Not x) And Z))
End Function

Private Function ConvertToWordArray(sMessage As String) As Long()
  Dim lMessageLength, lNumberOfWords, lBytePosition, lByteCount, lWordCount, lByte As Long
  Dim lWordArray() As Long
  Const MODULUS_BITS As Long = 512
  Const CONGRUENT_BITS As Long = 448
  lMessageLength = Len(sMessage)
  lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
  ReDim lWordArray(lNumberOfWords - 1)
  lBytePosition = 0
  lByteCount = 0
  Do Until lByteCount >= lMessageLength
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
    lByte = AscB(Mid$(sMessage, lByteCount + 1, 1))
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
    lByteCount = lByteCount + 1
  Loop
  lWordCount = lByteCount \ BYTES_TO_A_WORD
  lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
  lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
  lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
  lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
  ConvertToWordArray = lWordArray
End Function

Private Function Gamma0(ByVal x As Long) As Long
  Gamma0 = (s(x, 7) Xor s(x, 18) Xor r(x, 3))
End Function

Private Function Gamma1(ByVal x As Long) As Long
  Gamma1 = (s(x, 17) Xor s(x, 19) Xor r(x, 10))
End Function

Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
  If iShiftBits = 0 Then
    LShift = lValue
    Exit Function
  ElseIf iShiftBits = 31 Then
    If lValue And 1 Then
      LShift = &H80000000
    Else
      LShift = 0
    End If
    Exit Function
  ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
    LShift = 0 'Err.Raise 6
  End If
  If (lValue And m_l2Power(31 - iShiftBits)) Then
    LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
  Else
    LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
  End If
End Function

Private Function Maj(ByVal x As Long, ByVal Y As Long, ByVal Z As Long) As Long
  Maj = ((x And Y) Xor (x And Z) Xor (Y And Z))
End Function

Private Function r(ByVal x As Long, ByVal n As Long) As Long
  r = RShift(x, CInt(n And m_lOnBits(4)))
End Function

Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
  If iShiftBits = 0 Then
    RShift = lValue
    Exit Function
  ElseIf iShiftBits = 31 Then
    If lValue And &H80000000 Then
        RShift = 1
      Else
        RShift = 0
    End If
    Exit Function
  ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
    RShift = 0 'Err.Raise 6
  End If
  RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
  If (lValue And &H80000000) Then
    RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
  End If
End Function

Private Function s(ByVal x As Long, ByVal n As Long) As Long
  s = (RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4)))))
End Function

Public Function SHA256(sMessage As String) As String
  Dim HASH(7) As Long
  Dim M() As Long
  Dim W(63) As Long
  Dim a, b, c, d, E, F, g, h, i, j, T1, T2 As Long
  HASH(0) = &H6A09E667:  HASH(1) = &HBB67AE85:  HASH(2) = &H3C6EF372
  HASH(3) = &HA54FF53A:  HASH(4) = &H510E527F:  HASH(5) = &H9B05688C
  HASH(6) = &H1F83D9AB:  HASH(7) = &H5BE0CD19
  M = ConvertToWordArray(sMessage)
  For i = 0 To UBound(M) Step 16
    a = HASH(0):    b = HASH(1):    c = HASH(2)
    d = HASH(3):    E = HASH(4):    F = HASH(5)
    g = HASH(6):    h = HASH(7)
    For j = 0 To 63
      If j < 16 Then
        W(j) = M(j + i)
      Else
        W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
      End If
      T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(E)), ch(E, F, g)), K(j)), W(j))
      T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
      h = g:      g = F:      F = E
      E = AddUnsigned(d, T1)
      d = c:      c = b:      b = a
      a = AddUnsigned(T1, T2)
    Next j
    HASH(0) = AddUnsigned(a, HASH(0))
    HASH(1) = AddUnsigned(b, HASH(1))
    HASH(2) = AddUnsigned(c, HASH(2))
    HASH(3) = AddUnsigned(d, HASH(3))
    HASH(4) = AddUnsigned(E, HASH(4))
    HASH(5) = AddUnsigned(F, HASH(5))
    HASH(6) = AddUnsigned(g, HASH(6))
    HASH(7) = AddUnsigned(h, HASH(7))
    'DoEvents
  Next i
  SHA256 = UCase$(Right$("00000000" & Hex$(HASH(0)), 8) & Right$("00000000" & Hex$(HASH(1)), 8) & Right$("00000000" & Hex$(HASH(2)), 8) & Right$("00000000" & Hex$(HASH(3)), 8) & Right$("00000000" & Hex$(HASH(4)), 8) & Right$("00000000" & Hex$(HASH(5)), 8) & Right$("00000000" & Hex$(HASH(6)), 8) & Right$("00000000" & Hex$(HASH(7)), 8))
  For i = 1 To Len(SHA256) Step 2
    Retval = Retval & Chr(Int("&H" & Mid(SHA256, i, 2)))
  Next i
  SHA256 = Retval
End Function

Private Function Sigma0(ByVal x As Long) As Long
  Sigma0 = (s(x, 2) Xor s(x, 13) Xor s(x, 22))
End Function

Private Function Sigma1(ByVal x As Long) As Long
  Sigma1 = (s(x, 6) Xor s(x, 11) Xor s(x, 25))
End Function
Sub SignIt()
  Dim dll As String, F As Integer, dllContent$
  Dim signatur As String * 144
  dll = GetRegString(HKEY_CLASSES_ROOT, "CLSID\" & GetRegString(HKEY_CLASSES_ROOT, "AnGeL.XML\CLSID", "") & "\InprocServer32", "")
  F = FreeFile: Open dll For Binary As #F
  dllcontentxx$ = String(LOF(F), Chr(0))
  Get #F, , dllcontentxx$
  Close #F
  Open dll For Binary As #F
  'Seek #F, LOF(F)
  InitHashtable
  MMM$ = SHA256(dllcontentxx$)
  MMM$ = MMM$ & SHA256("Benedikt Hübschen")
  MMM$ = MMM$ & SHA256("20.02.2003")
  Set a = New clsBlowfish
  MMM$ = a.EncryptString(MMM$, DecryptString("1C2660C2D6", "AnGeL"))
  MMX$ = SHA256(MMM$)
  signatur = MMM$ & MMX$
  Seek #F, LOF(F) + 1
  Put #F, , signatur
  Close F
End Sub
