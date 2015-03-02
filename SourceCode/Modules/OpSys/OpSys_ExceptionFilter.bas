Attribute VB_Name = "OpSys_ExceptionFilter"
',-======================- ==-- -  -
'|   AnGeL - OpSys - ExceptionFilter
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIVILEGED_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE_VIOLATION = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const EXCEPTION_CONTROL_C_EXIT = &HC000013A

Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_EXECUTE_HANDLER = 1

Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(0 To 14) As Long
End Type

Private Type CONTEXT
  FltF0 As Double
  FltF1 As Double
  FltF2 As Double
  FltF3 As Double
  FltF4 As Double
  FltF5 As Double
  FltF6 As Double
  FltF7 As Double
  FltF8 As Double
  FltF9 As Double
  FltF10 As Double
  FltF11 As Double
  FltF12 As Double
  FltF13 As Double
  FltF14 As Double
  FltF15 As Double
  FltF16 As Double
  FltF17 As Double
  FltF18 As Double
  FltF19 As Double
  FltF20 As Double
  FltF21 As Double
  FltF22 As Double
  FltF23 As Double
  FltF24 As Double
  FltF25 As Double
  FltF26 As Double
  FltF27 As Double
  FltF28 As Double
  FltF29 As Double
  FltF30 As Double
  FltF31 As Double

  IntV0 As Double
  IntT0 As Double
  IntT1 As Double
  IntT2 As Double
  IntT3 As Double
  IntT4 As Double
  IntT5 As Double
  IntT6 As Double
  IntT7 As Double
  IntS0 As Double
  IntS1 As Double
  IntS2 As Double
  IntS3 As Double
  IntS4 As Double
  IntS5 As Double
  IntFp As Double
  IntA0 As Double
  IntA1 As Double
  IntA2 As Double
  IntA3 As Double
  IntA4 As Double
  IntA5 As Double
  IntT8 As Double
  IntT9 As Double
  IntT10 As Double
  IntT11 As Double
  IntRa As Double
  IntT12 As Double
  IntAt As Double
  IntGp As Double
  IntSp As Double
  IntZero As Double

  Fpcr As Double
  SoftFpcr As Double

  Fir As Double
  Psr As Long

  ContextFlags As Long
  Fill(4) As Long
End Type

Private Type EXCEPTION_POINTERS
  pExceptionRecord As EXCEPTION_RECORD
  ContextRecord As CONTEXT
End Type

Public Function ExceptionFilter(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
  Dim Rec As EXCEPTION_RECORD
  Dim i As Long, CTemp As String
  Dim strException As String
  Dim SubErrorCode As Long
  Dim Message As String
  Static iExceptionCount
  Static TerminationInProgess As Boolean, LastExceptionCode As Long
  If TerminationInProgess Then
    ' -= Unser "Exception Filter" wird beendet
    SubErrorCode = kernel32_SetUnhandledExceptionFilter(0&)
    ' -= Der Fehler wird noch einmal erzeugt,
    Call kernel32_RaiseException(LastExceptionCode, 0&, 0&, 0&)
  Else
    'Get current exception record.
    Rec = ExceptionPtrs.pExceptionRecord
    
    'If Rec.pExceptionRecord is not zero, then it is a nested exception and
    'Rec.pExceptionRecord points to another EXCEPTION_RECORD structure.  Follow
    'the pointers back to the original exception.
    Do Until Rec.pExceptionRecord = 0
      kernel32_RtlMoveMemory Rec, ByVal Rec.pExceptionRecord, Len(Rec)
    Loop
    
    'Translate the exception code into a user-friendly string.
    strException = GetExceptionText(Rec.ExceptionCode, SubErrorCode)
    
    'Raise an error to return control to the calling procedure.
    PutLog "<<==----E----X----C----E----P----T----I----O----N----==>>"
    PutLog "*** EXCEPTION: " & CStr(10000 + SubErrorCode) & " (" & strException & ")"
    PutLog ">>==----E----X----C----E----P----T----I----O----N----==<<"
  End If
  ExceptionFilter = EXCEPTION_CONTINUE_EXECUTION
  Exit Function
End Function


Private Function GetExceptionText(ByVal ExceptionCode As Long, SubErrorCode As Long) As String
  Dim strExceptionString As String
  Select Case ExceptionCode
    Case EXCEPTION_ACCESS_VIOLATION
      strExceptionString = "Access Violation"
      SubErrorCode = 1
    Case EXCEPTION_DATATYPE_MISALIGNMENT
      strExceptionString = "Data Type Misalignment"
      SubErrorCode = 2
    Case EXCEPTION_BREAKPOINT
      strExceptionString = "Breakpoint"
      SubErrorCode = 3
    Case EXCEPTION_SINGLE_STEP
      strExceptionString = "Single Step"
      SubErrorCode = 4
    Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
      strExceptionString = "Array Bounds Exceeded"
      SubErrorCode = 5
    Case EXCEPTION_FLT_DENORMAL_OPERAND
      strExceptionString = "Float Denormal Operand"
      SubErrorCode = 6
    Case EXCEPTION_FLT_DIVIDE_BY_ZERO
      strExceptionString = "Divide By Zero"
      SubErrorCode = 7
    Case EXCEPTION_FLT_INEXACT_RESULT
      strExceptionString = "Floating Point Inexact Result"
      SubErrorCode = 8
    Case EXCEPTION_FLT_INVALID_OPERATION
      strExceptionString = "Invalid Operation"
      SubErrorCode = 9
    Case EXCEPTION_FLT_OVERFLOW
      strExceptionString = "Float Overflow"
      SubErrorCode = 10
    Case EXCEPTION_FLT_STACK_CHECK
      strExceptionString = "Float Stack Check"
      SubErrorCode = 11
    Case EXCEPTION_FLT_UNDERFLOW
      strExceptionString = "Float Underflow"
      SubErrorCode = 12
    Case EXCEPTION_INT_DIVIDE_BY_ZERO
      strExceptionString = "Integer Divide By Zero"
      SubErrorCode = 13
    Case EXCEPTION_INT_OVERFLOW
      strExceptionString = "Integer Overflow"
      SubErrorCode = 14
    Case EXCEPTION_PRIVILEGED_INSTRUCTION
      strExceptionString = "Privileged Instruction"
      SubErrorCode = 15
    Case EXCEPTION_IN_PAGE_ERROR
      strExceptionString = "In Page Error"
      SubErrorCode = 16
    Case EXCEPTION_ILLEGAL_INSTRUCTION
      strExceptionString = "Illegal Instruction"
      SubErrorCode = 17
    Case EXCEPTION_NONCONTINUABLE_EXCEPTION
      strExceptionString = "Non Continuable Exception"
      SubErrorCode = 18
    Case EXCEPTION_STACK_OVERFLOW
      strExceptionString = "Stack Overflow"
      SubErrorCode = 19
    Case EXCEPTION_INVALID_DISPOSITION
      strExceptionString = "Invalid Disposition"
      SubErrorCode = 20
    Case EXCEPTION_GUARD_PAGE_VIOLATION
      strExceptionString = "Guard Page Violation"
      SubErrorCode = 21
    Case EXCEPTION_INVALID_HANDLE
      strExceptionString = "Invalid Handle"
      SubErrorCode = 22
    Case EXCEPTION_CONTROL_C_EXIT
      strExceptionString = "Control-C Exit"
      SubErrorCode = 23
    Case Else
      strExceptionString = "Unknown (&H" & Right("00000000" & Hex(ExceptionCode), 8) & ")"
      SubErrorCode = 24
  End Select
  GetExceptionText = strExceptionString
End Function

