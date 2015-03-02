Attribute VB_Name = "Server_Connection"
Option Explicit

Public MyNick As String
Public PrimaryNick As String
Public SecondaryNick As String
Public RealName As String
Public MyHostmask As String
Public MyIPmask As String
Public ConnectTime As Date
Public GotServerPong As Boolean
Public CountServerPing As Long
Public BytesSent As Long
Public MaxBytesToServer As Long
Public BlockSends As Boolean
Public LinesSent As Long
Public LineToComplete As String
Public GetNextPartRest As String
Public WaitThisLine As Boolean
Public CommandPrefix As String
Public RestrictedIndex As Long
Public Const EmptyLine As String = " "
Public VersionReply As String, FakeIDKick As Boolean, HideBot As Boolean
Public PumpDCC As Long, Invisible As Boolean
Public DontConnect As Boolean, ConnectTryCounter As Integer
Public Initializing As Boolean, AutoUpdate As Boolean, Exitting As Boolean
Public MassModeOps As String, MassModeDeops As String, MassModeBans As String, MassModeUnbans As String, MassModeUnExcept As String, MassModeUnInvite As String, MassModeExcept As String, MassModeInvite As String, MassModeVoice As String, MassModeHOps As String, MassModeDevoice As String, MassModeDehops As String
Public MassModeOps2 As String, MassModeDeops2 As String, MassModeBans2 As String, MassModeUnbans2 As String, MassModeUnExcept2 As String, MassModeUnInvite2 As String, MassModeExcept2 As String, MassModeInvite2 As String, MassModeVoice2 As String, MassModeHOps2 As String, MassModeDevoice2 As String, MassModeDehops2 As String
Public IdentCommand As String, LastDay As Long, HubBot As Boolean, FormPassword As String
Public ServerName As String
Public StartUpTime As Date, BotNetNick As String, BotnetSocket As Long
Public TelnetSocket As Long, UnbanCount As Long, FileAreaEnabled As Boolean
Public UnbanTime As Currency, BotRestart As Boolean, LastEvent As Date, LastSendTime As Date
Public ProxyToConnect As String, ServerToConnect As String
Public SentLogin As Boolean, JustJumped As Boolean
Public Const MaxPrivEvents = 15, MaxUserEvents = 8
Public TelnetPort As Long, BotnetPort As Long, UsedServer As Long
Public LastWhoisOutput As String
Public CountCTCPs As Long, UnIgnoreTimes As Long, DontSeen As Boolean, BanLimit As Long, ExceptLimit As Long, InviteLimit As Long
Public ServBuff As String
Public DefaultConnectDelay As Currency, FailureConnectDelay As Currency, JumpConnectDelay As Currency, CurrentConnectDelay As Currency, FirstTimeKeyword As String
Public ListenSock As Long, ServerSocket As Long, OldWndProc As Long
Public IdentSocket As Long
Public ResolveIP As String, MyIP As String
Public ServerPort As Long
Public StrictHost As Boolean, AllowCreateObject As Boolean, AllowRunToS As Boolean
Public ExtReply As String
Public LastTick As Currency
Public RestrictCycle As Boolean
Public UseIDENTD As Boolean
Public RouterWorkAround As Boolean
