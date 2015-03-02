Attribute VB_Name = "OpSys_API"
Option Explicit

Declare Function version_VerQueryValueA Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Declare Function version_GetFileVersionInfoSizeA Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function version_GetFileVersionInfoA Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

Declare Function kernel32_FindClose Lib "kernel32.dll" Alias "FindClose" (ByVal hFindFile As Long) As Long
Declare Function kernel32_FindFirstFileA Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function kernel32_FindNextFileA Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function kernel32_GetPrivateProfileStringA Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function kernel32_WritePrivateProfileStringA Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function kernel32_GetPrivateProfileSectionA Lib "kernel32.dll" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function kernel32_WritePrivateProfileSectionA Lib "kernel32.dll" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function kernel32_CreateFileA Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function kernel32_DeviceIoControl Lib "kernel32.dll" Alias "DeviceIoControl" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Integer, ByVal nInBufferSize As Integer, lpOutBuffer As Long, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Any) As Long
Declare Function kernel32_GetFileAttributesA Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function kernel32_GetVolumeInformationA Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function kernel32_OpenProcess Lib "kernel32.dll" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function kernel32_RegisterServiceProcess Lib "kernel32.dll" Alias "RegisterServiceProcess" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Declare Function kernel32_CreateThread Lib "kernel32.dll" Alias "CreateThread" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function kernel32_GetVersionExA Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Declare Function kernel32_GetTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
Declare Function kernel32_SetPriorityClass Lib "kernel32.dll" Alias "SetPriorityClass" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function kernel32_VerLanguageNameA Lib "kernel32.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Declare Function kernel32_GetCurrentProcessId Lib "kernel32.dll" Alias "GetCurrentProcessId" () As Long
Declare Function kernel32_SetUnhandledExceptionFilter Lib "kernel32.dll" Alias "SetUnhandledExceptionFilter" (ByVal lpTopLevelExceptionFilter As Long) As Long
Declare Function kernel32_FormatMessageA Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function kernel32_GetSystemDirectoryA Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Integer
Declare Function kernel32_lstrlenA Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Declare Sub kernel32_RtlMoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpDestination As Any, lpSource As Any, ByVal dwLength As Long)
Declare Sub kernel32_RaiseException Lib "kernel32.dll" Alias "RaiseException" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Declare Function advapi32_FileEncryptionStatusA Lib "advapi32.dll" Alias "FileEncryptionStatusA" (ByVal lpFileName As String, lpStatus As Long) As Long
Declare Function advapi32_EncryptFileA Lib "advapi32.dll" Alias "EncryptFileA" (ByVal lpFileName As String) As Long
Declare Function advapi32_DecryptFileA Lib "advapi32.dll" Alias "DecryptFileA" (ByVal lpFileName As String, ByVal dwReserved As Long) As Long
Declare Function advapi32_RegQueryValueExA Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function advapi32_OpenSCManagerA Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function advapi32_CreateServiceA Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Declare Function advapi32_DeleteService Lib "advapi32.dll" Alias "DeleteService" (ByVal hService As Long) As Long
Declare Function advapi32_RegOpenKeyA Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As HKEY_CONSTANTS, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function advapi32_RegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As HKEY_CONSTANTS) As Long
Declare Function advapi32_OpenServiceA Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long

Declare Function user32_SetWindowLongA Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal NIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function user32_SetWindowPos Lib "user32.dll" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal WFlags As Long) As Long
Declare Function user32_FindWindowA Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function user32_CallWindowProcA Lib "user32.dll" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function user32_CreateWindowExA Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function user32_DestroyWindow Lib "user32.dll" Alias "DestroyWindow" (ByVal hwnd As Long) As Long

Declare Function iphlpapi_GetTcpTable Lib "iphlpapi.dll" Alias "GetTcpTable" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Declare Function winsock2_WSAStartup Lib "ws2_32.dll" Alias "WSAStartup" (ByVal wVersionRequested As Long, lpWSADataType As LPWSADATA) As Long
Declare Function winsock2_WSACleanup Lib "ws2_32.dll" Alias "WSACleanup" () As Long
Declare Function winsock2_WSAIsBlocking Lib "ws2_32.dll" Alias "WSAIsBlocking" () As Long
Declare Function winsock2_WSACancelBlockingCall Lib "ws2_32.dll" Alias "WSACancelBlockingCall" () As Long
Declare Function winsock2_WSAStringToAddressA Lib "ws2_32.dll" Alias "WSAStringToAddressA" (ByVal AddressString As String, ByVal AddressFamily As Long, ByVal ProtocolInfo As Long, ByVal Address As String, AddressLenght As Long) As Long
Declare Function winsock2_WSAAddressToStringA Lib "ws2_32.dll" Alias "WSAAddressToStringA" (ByVal Address As String, ByVal AddressLenght As Long, ByVal ProtocolInfo As Long, ByVal AddressString As String, AddressStringLength As Long) As Long
Declare Function winsock2_WSAAsyncGetHostByName Lib "ws2_32.dll" Alias "WSAAsyncGetHostByName" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal HostName As String, Buffer As Any, ByVal BufferLen As Long) As Long
Declare Function winsock2_WSAAsyncGetHostByAddr Lib "ws2_32.dll" Alias "WSAAsyncGetHostByAddr" (ByVal hwnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, Buf As Any, ByVal buflen As Long) As Long
Declare Function winsock2_WSACancelAsyncRequest Lib "ws2_32.dll" Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As Long
Declare Function winsock2_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Declare Function winsock2_gethostname Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Declare Function winsock2_WSAAsyncSelect Lib "ws2_32.dll" Alias "WSAAsyncSelect" (ByVal SocketNumber As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Declare Function winsock2_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal CP As String) As Long
Declare Function winsock2_inet_ntoa Lib "ws2_32.dll" Alias "inet_ntoa" (ByVal Inn As Long) As Long
Declare Function winsock2_accept Lib "ws2_32.dll" Alias "accept" (ByVal SocketNumber As Long, ByVal Address As String, ByRef AddressLength As Long) As Long
Declare Function winsock2_recv Lib "ws2_32.dll" Alias "recv" (ByVal SocketNumber As Long, ByVal Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Declare Function winsock2_recvfrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal SocketNumber As Long, ByVal Buf As Any, ByVal buflen As Long, ByVal Flags As Long, ByVal From As String, fromlen As Long) As Long
Declare Function winsock2_send Lib "ws2_32.dll" Alias "send" (ByVal SocketNumber As Long, Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Declare Function winsock2_sendto Lib "ws2_32.dll" Alias "sendto" (ByVal SocketNumber As Long, Buf As Any, ByVal buflen As Long, ByVal Flags As Long, ByVal Address As String, AddressLength As Long) As Long
Declare Function winsock2_listen Lib "ws2_32.dll" Alias "listen" (ByVal SocketNumber As Long, ByVal backlog As Long) As Long
Declare Function winsock2_connect Lib "ws2_32.dll" Alias "connect" (ByVal SocketNumber As Long, ByVal Address As String, ByVal AddressLength As Long) As Long
Declare Function winsock2_bind Lib "ws2_32.dll" Alias "bind" (ByVal SocketNumber As Long, ByVal Address As String, ByVal AddressLength As Long) As Long
Declare Function winsock2_socket Lib "ws2_32.dll" Alias "socket" (ByVal AddressFamily As Long, ByVal SocketType As Long, ByVal Protocol As Long) As Long
Declare Function winsock2_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal SocketNumber As Long) As Long
Declare Function winsock2_shutdown Lib "ws2_32.dll" Alias "shutdown" (ByVal SocketNumber As Long, ByVal how As Long) As Long
Declare Function winsock2_getpeername Lib "ws2_32.dll" Alias "getpeername" (ByVal SocketNumber As Long, ByVal Address As String, AddressLength As Long) As Long
Declare Function winsock2_getsockname Lib "ws2_32.dll" Alias "getsockname" (ByVal SocketNumber As Long, ByVal Address As String, AddressLength As Long) As Long
Declare Function winsock2_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal SocketNumber As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Declare Function winsock2_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal SocketNumber As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Declare Function winsock2_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal SocketNumber As Long, ByVal cmd As Long, argp As Long) As Long
Declare Function winsock2_gethostbyaddr Lib "ws2_32.dll" Alias "gethostbyaddr" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const LANG_NEUTRAL = &H0
Public Const SUBLANG_DEFAULT = &H1
Public Const GWL_WNDPROC = (-4)
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const REG_SZ = 1
Public Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_SET_INFORMATION = &H200
Public Const VER_PLATFORM_WIN32_NT& = 2
Public Const VER_PLATFORM_WIN32_WINDOWS& = 1
Public Const WSAEINTR = 10004
Public Const WSAEBADF = 10009
Public Const WSAEACCES = 10013
Public Const WSAEFAULT = 10014
Public Const WSAEINVAL = 10022
Public Const WSAEMFILE = 10024
Public Const WSAEWOULDBLOCK = 10035
Public Const WSAEINPROGRESS = 10036
Public Const WSAEALREADY = 10037
Public Const WSAENOTSOCK = 10038
Public Const WSAEDESTADDRREQ = 10039
Public Const WSAEMSGSIZE = 10040
Public Const WSAEPROTOTYPE = 10041
Public Const WSAENOPROTOOPT = 10042
Public Const WSAEPROTONOSUPPORT = 10043
Public Const WSAESOCKTNOSUPPORT = 10044
Public Const WSAEOPNOTSUPP = 10045
Public Const WSAEPFNOSUPPORT = 10046
Public Const WSAEAFNOSUPPORT = 10047
Public Const WSAEADDRINUSE = 10048
Public Const WSAEADDRNOTAVAIL = 10049
Public Const WSAENETDOWN = 10050
Public Const WSAENETUNREACH = 10051
Public Const WSAENETRESET = 10052
Public Const WSAECONNABORTED = 10053
Public Const WSAECONNRESET = 10054
Public Const WSAENOBUFS = 10055
Public Const WSAEISCONN = 10056
Public Const WSAENOTCONN = 10057
Public Const WSAESHUTDOWN = 10058
Public Const WSAETOOMANYREFS = 10059
Public Const WSAETIMEDOUT = 10060
Public Const WSAECONNREFUSED = 10061
Public Const WSAELOOP = 10062
Public Const WSAENAMETOOLONG = 10063
Public Const WSAEHOSTDOWN = 10064
Public Const WSAEHOSTUNREACH = 10065
Public Const WSAENOTEMPTY = 10066
Public Const WSAEPROCLIM = 10067
Public Const WSAEUSERS = 10068
Public Const WSAEDQUOT = 10069
Public Const WSAESTALE = 10070
Public Const WSAEREMOTE = 10071
Public Const WSASYSNOTREADY = 10091
Public Const WSAVERNOTSUPPORTED = 10092
Public Const WSANOTINITIALISED = 10093
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004
Public Const WSANO_ADDRESS = 11004
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CAPTION = &HC00000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const AF_INET = 2
Public Const AF_INET6 = 23
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17
Public Const WSA_Limit_SocketAddress = 64
Public Const MAXGETHOSTSTRUCT = 1024
Public Const hostentasync_size = MAXGETHOSTSTRUCT + 16
Public Const hostent_size = 16
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SO_LINGER = &H80&
Public Const SOL_Socket = &HFFFF&
Public Const SO_BROADCAST = &H20
Public Const FD_READ = &H1&
Public Const FD_WRITE = &H2&
Public Const FD_OOB = &H4&
Public Const FD_winsock2_accept = &H8&
Public Const FD_connect = &H10&
Public Const FD_CLOSE = &H20&
Public Const WM_USER = 1024

Public Enum HKEY_CONSTANTS
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_CURRENT_USER = &H80000001
  HKEY_DYN_DATA = &H80000006
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_USERS = &H80000003
End Enum

Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(1 To 128) As Byte
End Type

Type HOSTENT
  h_name As Long
  h_aliases As Long
  h_addrtype As Integer
  h_length As Integer
  h_addr_list As Long
End Type

Type HostEntAsync
  h_name As Long
  h_aliases As Long
  h_addrtype As Integer
  h_length As Integer
  h_addr_list As Long
  h_asyncbuffer(hostentasync_size) As Byte
End Type

Type LPWSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription As String * 257
  szSystemStatus As String * 129
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpVendorInfo As Long
End Type

Type LingerType
  l_onoff As Integer
  l_linger As Integer
End Type

