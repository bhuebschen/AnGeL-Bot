Attribute VB_Name = "OpSys_NTService"
',-======================- ==-- -  -
'|   AnGeL - OpSys - NTService
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Public NTServiceName As String
Public IsNTService As Boolean

Private hnd As Long
Private h(0 To 1) As Long
Private hStopEvent As Long
Private hStartEvent As Long
Private hStopPendingEvent As Long
Private hServiceStatus As Long
Private ServiceName() As Byte
Private ServiceNamePtr As Long
Private ServiceStatus As SERVICE_STATUS


Private Const INFINITE As Long = -1
Private Const WAIT_TIMEOUT As Long = 258
Private Const SERVICE_ALWAYS_START As Long = &H2
Private Const SERVICE_DEMAND_START As Long = &H3
Private Const SERVICE_ERROR_NORMAL As Long = &H1
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)




Function StartNTService() As Boolean
  ' -= Prüfen ob WindowsNT
  If WinNTOS Then
    ' -= Events erzeugen
    hStopEvent = CreateEvent(0, 1, 0, vbNullString)
    hStopPendingEvent = CreateEvent(0, 1, 0, vbNullString)
    hStartEvent = CreateEvent(0, 1, 0, vbNullString)
    
    ' -= Service Namen konvertieren (zu ANSI)
    ServiceName = StrConv(NTServiceName, vbFromUnicode)
    ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))
    
    hnd = StartAsService
    h(0) = hnd
    h(1) = hStartEvent
    
    If Not WinNTOSService Then
      CloseHandle hnd
      StartNTService = False
    Else
      SetServiceState SERVICE_RUNNING
      StartNTService = True
    End If
  Else
    StartNTService = False
  End If
End Function


Private Function FunctionPointer(ByVal Address As Long) As Long
  FunctionPointer = Address
End Function


Private Function StartAsService() As Long
  Dim ThreadId As Long
  StartAsService = kernel32_CreateThread(0&, 0&, AddressOf ServiceThread, 0&, 0&, ThreadId)
End Function


Private Sub ServiceThread(ByVal Dummy As Long)
  Dim ServiceTableEntry As SERVICE_TABLE
  ServiceTableEntry.lpServiceName = ServiceNamePtr
  ServiceTableEntry.lpServiceProc = FunctionPointer(AddressOf ServiceMain)
  StartServiceCtrlDispatcher ServiceTableEntry
End Sub


Private Function WinNTOSService() As Boolean
  WinNTOSService = (WaitForMultipleObjects(2&, h(0), 0&, INFINITE) = 1&)
End Function


Private Sub ServiceMain(ByVal dwArgc As Long, ByVal lpszArgv As Long)
  ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS
  ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_SHUTDOWN
  ServiceStatus.dwWin32ExitCode = 0&
  ServiceStatus.dwServiceSpecificExitCode = 0&
  ServiceStatus.dwCheckPoint = 0&
  ServiceStatus.dwWaitHint = 0&
  hServiceStatus = RegisterServiceCtrlHandler(NTServiceName, AddressOf Handler)
  SetServiceState SERVICE_START_PENDING
  SetEvent hStartEvent
  WaitForSingleObject hStopEvent, INFINITE
End Sub


Private Sub Handler(ByVal fdwControl As Long)
  Select Case fdwControl
    Case SERVICE_ACCEPT_SHUTDOWN, SERVICE_CONTROL_STOP
      SetServiceState SERVICE_STOP_PENDING
      SetEvent hStopPendingEvent
    Case Else
      SetServiceState
  End Select
End Sub


Private Sub SetServiceState(Optional ByVal NewState As SERVICE_STATE = 0&)
  If NewState <> 0& Then ServiceStatus.dwCurrentState = NewState
  SetServiceStatus hServiceStatus, ServiceStatus
End Sub


Public Sub StopService()
    SetServiceState SERVICE_STOPPED
    SetEvent hStopEvent
    WaitForSingleObject hnd, INFINITE
    CloseHandle hnd
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
End Sub

Public Function ContinueService() As Boolean
  ContinueService = (WaitForSingleObject(hStopPendingEvent, 25) = WAIT_TIMEOUT)
End Function


Public Sub InstallService()
  Dim ServiceManager As Long, ServiceHandle As Long
  ServiceManager = advapi32_OpenSCManagerA(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
  ServiceHandle = advapi32_CreateServiceA(ServiceManager, NTServiceName, NTServiceName & " - " & BotNetNick, SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS Or SERVICE_USER_DEFINED_CONTROL, SERVICE_ALWAYS_START, SERVICE_ALWAYS_START, App.Path & "\AnGeL.exe serv", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
  CloseHandle ServiceHandle
  CloseHandle ServiceManager
End Sub


Public Sub UninstallService()
  Dim ServiceManager As Long, ServiceHandle As Long
  ServiceManager = advapi32_OpenSCManagerA(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
  ServiceHandle = advapi32_OpenServiceA(ServiceManager, NTServiceName, SERVICE_ALL_ACCESS)
  advapi32_DeleteService ServiceHandle
  CloseHandle ServiceHandle
  CloseHandle ServiceManager
End Sub

