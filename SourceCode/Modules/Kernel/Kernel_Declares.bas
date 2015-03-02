Attribute VB_Name = "Kernel_Declares"
Option Explicit

Public AnGeLFiles As New clsFileTypes

'AddUser replies
Public Const AU_Success As Byte = 0
Public Const AU_TooLong As Byte = 1
Public Const AU_UserExists As Byte = 2
Public Const AU_InvalidNick As Byte = 3

'RemUser replies
Public Const RU_Success As Byte = 0
Public Const RU_UserNotFound As Byte = 1

'Chattr replies
Public Const CH_Success As Byte = 0
Public Const CH_NoChanges As Byte = 1
Public Const CH_NoChanFlag As Byte = 2

'AddHost replies
Public Const AH_Success As Byte = 0
Public Const AH_UserNotFound As Byte = 1
Public Const AH_InvalidHost As Byte = 2
Public Const AH_AlreadyThere As Byte = 3
Public Const AH_MatchingUser As Byte = 4
Public Const AH_TooManyHosts As Byte = 5
Public Const AH_DENIED As Byte = 6

'RemHost replies
Public Const RH_Success As Byte = 0
Public Const RH_UserNotFound As Byte = 1
Public Const RH_HostNotFound As Byte = 2

Public Const ChanStat_OK As String = "."
Public Const ChanStat_NotOn As String = "!"
Public Const ChanStat_NeedInvite As String = "i"
Public Const ChanStat_NeedKey As String = "k"
Public Const ChanStat_BadLimit As String = "l"
Public Const ChanStat_ImBanned As String = "b"
Public Const ChanStat_Left As String = "x"
Public Const ChanStat_Duped As String = "d"
Public Const ChanStat_OutLimits As String = "m"
Public Const ChanStat_Unsup As String = "u"
Public Const ChanStat_RegisteredOnly As String = "r"

Public Const T_GA = "ù"   '249
Public Const T_WILL = "û" '251
Public Const T_WONT = "ü" '252
Public Const T_DO = "ý"   '253
Public Const T_DONT = "þ" '254
Public Const T_IAC = "ÿ"  '255
Public Const T_ECHO = "" '1
Public Const T_SPGA = "" '3 (Suppress Go Ahead)

' - Befehle - '
Public Const T_EOF = "ì"                                                                                                                                                                                                                                                                                                                   '236 (EOF)
Public Const T_AYT = "ö"                                                                                                                                                                                                                                                                                                                   '246 (Are you there?)
Public Const T_BRK = "ó"                                                                                                                                                                                                                                                                                                                   '243 (Break)
Public Const T_EC = "÷"                                                                                                                                                                                                                                                                                                                    '247 (Erase Character)
Public Const T_EL = "ø"                                                                                                                                                                                                                                                                                                                    '248 (Erase Line)
Public Const T_NOOP = "ñ"                                                                                                                                                                                                                                                                                                                  '241 (NOOP)
Public Const T_AP = "î"                                                                                                                                                                                                                                                                                                                    '238 (Abort Process)
Public Const T_AO = "õ"                                                                                                                                                                                                                                                                                                                    '245 (Abort Output)
Public Const T_IP = "ô"                                                                                                                                                                                                                                                                                                                    '244 (Interupt Process)
Public Const T_SP = "í"                                                                                                                                                                                                                                                                                                                    '237 (Supsend Process)
Public Const T_EOR = "ï"                                                                                                                                                                                                                                                                                                                   '239 (End of Record)
' - Befehle - '

Public Const SF_Colors As Byte = 1
Public Const SF_Echo As Byte = 2
Public Const SF_Status As Byte = 3
Public Const SF_DCC As Byte = 4
Public Const SF_AutoWHO As Byte = 5
Public Const SF_Silent As Byte = 6
Public Const SF_LoggedIn As Byte = 7
Public Const SF_LF_ONLY As Byte = 8
Public Const SF_LocalVisibleUser As Byte = 9
'Saved Socket flags -> change in GiveSockFlags too!
  Public Const SavedStart As Byte = 20
  Public Const SavedEnd As Byte = 28
  Public Const SF_ExtraHelp As Byte = 20
  Public Const SF_Local_JP As Byte = 21
  Public Const SF_Local_Talk As Byte = 22
  Public Const SF_Local_Bot As Byte = 23
  Public Const SF_Botnet_JP As Byte = 24
  Public Const SF_Botnet_Talk As Byte = 25
  Public Const SF_Botnet_Bot As Byte = 26
  Public Const SF_PrivToBot As Byte = 27      ' Only for masters
  Public Const SF_UserCommands As Byte = 28   '        "
'---
  Public Const SF_Telnet As Byte = 29
Public Const SF_NO As String = "0"
Public Const SF_YES As String = "1"
Public Const SF_Status_Dead As String = SF_NO
Public Const SF_Status_DCCWaiting As String = "d"
Public Const SF_Status_DCCInit As String = "i"
Public Const SF_Status_Party As String = "P"
Public Const SF_Status_FileArea As String = "x"
Public Const SF_Status_PersonalSetup As String = "?"
Public Const SF_Status_POLSetup As String = "p"
Public Const SF_Status_ChanSetup As String = "s"
Public Const SF_Status_BotSetup As String = "S"
Public Const SF_Status_AUTHSetup As String = "A"
Public Const SF_Status_NETSetup As String = "n"
Public Const SF_Status_KISetup As String = "K"
Public Const SF_Status_RelaySrv As String = "R"
Public Const SF_Status_RelayCli As String = "r"
Public Const SF_Status_SharingSetup As String = "+"
Public Const SF_Status_BotNetParty As String = "N"
Public Const SF_Status_FileWaiting As String = "f"
Public Const SF_Status_File As String = "F"
Public Const SF_Status_SendFileWaiting As String = "q"
Public Const SF_Status_SendFile As String = "Q"
Public Const SF_Status_UserGetName As String = "*"
Public Const SF_Status_UserGetPass As String = "^"
Public Const SF_Status_UserChoosePass As String = "°"
Public Const SF_Status_UserChooseColors As String = "`"
Public Const SF_Status_InitBotLink As String = "b"
Public Const SF_Status_BotGetName As String = "."
Public Const SF_Status_BotGetPass As String = ","
Public Const SF_Status_BotLinking As String = "-"
Public Const SF_Status_BotPreCache As String = "/"
Public Const SF_Status_Bot As String = "B"
Public Const SF_Status_Ident As String = "I"
Public Const SF_Status_ScriptSocket As String = "c"
Public Const SF_Status_ScriptUser As String = "§"
Public Const SF_Status_IdentListen As String = "l"
Public Const SF_Status_BotnetListen As String = "L"
Public Const SF_Status_TelnetListen As String = "X"
Public Const SF_Status_Server As String = "Y"

