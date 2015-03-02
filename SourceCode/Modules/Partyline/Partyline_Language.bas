Attribute VB_Name = "Partyline_Language"
',-======================- ==-- -  -
'|   AnGeL - Partyline - Language
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

'| Normal messages
' -—————————- -- -  -
Public MSG_EnterPWD         As String
Public MSG_ChoosePWD        As String
Public MSG_ThanksPWD        As String
Public MSG_CanYouSeeColors  As String
Public MSG_CanYouSeeColTwo  As String
Public MSG_ConnectedTo      As String
Public MSG_FirstLogin       As String
Public MSG_SecurityInfo     As String
Public MSG_ThankYou         As String
Public MSG_ShortIntro       As String
Public MSG_PLJoin           As String
Public MSG_PLJoinChan       As String
Public MSG_PLBotNetJoin     As String
Public MSG_PLLeave          As String
Public MSG_PLLeaveMsg       As String
Public MSG_PLBotNetLeave    As String
Public MSG_PLBotNetLeaveMsg As String
Public MSG_PLTalk           As String
Public MSG_PLBotTalk        As String
Public MSG_PLBotNetTalk     As String
Public MSG_PLAct            As String
Public MSG_PLBotNetAct      As String
Public MSG_PLNick           As String
Public MSG_PLAway           As String
Public MSG_PLBotNetAway     As String
Public MSG_PLBack           As String
Public MSG_PLBotNetBack     As String
Public MSG_PLNote           As String
Public MSG_PLNotes          As String
Public MSG_PLHelpIntro      As String
Public MSG_PLHelpFlagW      As String
Public MSG_PLHelpFlagO      As String
Public MSG_PLHelpFlagT      As String
Public MSG_PLHelpFlagCM     As String
Public MSG_PLHelpFlagM      As String
Public MSG_PLHelpFlagCN     As String
Public MSG_PLHelpFlagN      As String
Public MSG_PLHelpFlagS      As String
Public MSG_PLHelpOutro      As String
Public MSG_PLDCCIncoming    As String
Public MSG_PLDCCOpened      As String
Public MSG_PLDCCClosed      As String
Public MSG_PLDCCFirst       As String
Public MSG_PLDCCWrongPW     As String
Public MSG_PLDCCFailed      As String
Public MSG_PLDCCRefused     As String
Public MSG_PLDCCNoAcc       As String
Public MSG_PLTelNetIncoming As String
Public MSG_PLTelNetOpened   As String
Public MSG_PLTelNetClosed   As String
Public MSG_PLTelNetFirst    As String
Public MSG_PLTelNetNoUser   As String
Public MSG_PLTelNetWrongPW  As String
Public MSG_PLTelNetNoAcc    As String
Public MSG_PLTelNetBot      As String
Public MSG_PLBotNetIncoming As String
Public MSG_PLBotNetLinking  As String
Public MSG_PLBotNetLinkFrom As String
Public MSG_PLBotNetError    As String
Public MSG_PLBotNetLost     As String
Public MSG_PLBotNetNoLink   As String
Public MSG_PLBotNetConnect  As String
Public MSG_PLIdentRequest   As String
Public MSG_PLFloodStart     As String
Public MSG_PLFloodEnd       As String
Public MSG_PLNickDid        As String
Public MSG_PLNickTried      As String
Public MSG_PLNickFailed     As String
Public MSG_PLWrongBot       As String
Public MSG_PLYouWereBooted  As String
Public MSG_PLYouWereBooted2 As String
Public MSG_PLNickWasBooted  As String
Public MSG_PLNickWasBooted2 As String
  
'| Botnet messages
' -—————————- -- -  -
Public MSG_BNBooted         As String
Public MSG_BNBooted2        As String
Public MSG_BNSendUser       As String
Public MSG_BNSendUserPass   As String
Public MSG_BNPingTimeout    As String
Public MSG_BNNoPass         As String
Public MSG_BNBadPass        As String
Public MSG_BNNoAccess       As String
Public MSG_BNLoop           As String
Public MSG_BNLeafLink       As String
Public MSG_BNLeafLinks      As String
Public MSG_BNBogusLink      As String
Public MSG_BNRestructure    As String
Public MSG_BNConnect        As String
Public MSG_BNDisconnect     As String
Public MSG_BNLostBot        As String
Public MSG_BNWrongBot       As String

'| Timed messages
' -—————————- -- -  -
Public MSG_PLSwitchLogs1    As String
Public MSG_PLSwitchLogs2    As String
Public MSG_PLRequestedGo    As String
Public MSG_PLTelNetNickTO   As String
Public MSG_PLTelNetPassTO   As String
Public MSG_PLBotNetLinkTO   As String
Public MSG_PLBotNetConnTO   As String
Public MSG_PLBotNetNickTO   As String
Public MSG_PLBotNetPassTO   As String
Public MSG_PLDCCGetTO       As String
Public MSG_PLDCCSendTO      As String
Public MSG_PLDCCConnTO      As String

'| IRC messages
' -—————————- -- -  -
Public IRC_NoRepeats        As String
Public IRC_Seen_NoNick      As String
Public IRC_Seen_ErrNick     As String
Public IRC_Whois_YUnknown   As String
Public IRC_Whois_UUnknown   As String
Public IRC_Whois_NoInfo     As String
Public IRC_Whatis_NotFound  As String

'| Error messages
' -—————————- -- -  -
Public ERR_Login_WrongPass  As String
Public ERR_Login_Unknown    As String
Public ERR_Login_NoChat     As String
Public ERR_Pass_TooShort    As String
Public ERR_Pass_NoSpaces    As String
Public ERR_Pass_TooWeak     As String
Public ERR_Nick_TooLong     As String
Public ERR_Nick_Erroneous   As String
Public ERR_Nick_InUse       As String
Public ERR_NotOnLocalChans  As String
Public ERR_UserNotFound     As String
Public ERR_BotNotFound      As String
Public ERR_ServerLost       As String
Public ERR_ServerFailed     As String
Public ERR_NotAllowed       As String
Public ERR_CommandUsage     As String
Public MSG_PL_LookHelp      As String
Public MSG_FA_LookHelp      As String

Public Function InitLanguage() As Boolean ' : AddStack "LanguageFile_InitLanguage()"
Dim LanguageFile As String
  
  '| Normal messages
  ' -—————————- -- -  -
  MSG_EnterPWD = "Enter your password."
  MSG_ChoosePWD = "Welcome! Please choose a password (at least 6 characters long):"
  MSG_ThanksPWD = "Thank you. Your password is now ""#1#""."
  MSG_CanYouSeeColors = "Can you see 10colors in your IRC Client? Type ""yes"" or ""no""."
  MSG_CanYouSeeColTwo = "Please answer my question! Type ""yes"" or ""no""."
  MSG_ConnectedTo = "Connected to #1#, running AnGeL"
  MSG_FirstLogin = "2-> Hiya! This is your first login. Please write down your##2-> password, you'll need it every time you log on."
  MSG_SecurityInfo = "2-> ##2-> #CLCritical#Security info: The IDENT command of this bot has##2-> been changed to '/msg <bot> #1# <pass>'."
  MSG_ThankYou = "2*** Thank you for using an AnGeL Bot! C-Ya!"
  MSG_ShortIntro = "Commands start with '.' (like '.quit' or '.help')##Everything else goes out to the party line."
  MSG_PLJoin = "3*** #1# joined the party line."
  MSG_PLJoinChan = "3*** #1# joined the the channel #2#."
  MSG_PLBotNetJoin = "3*** 14(#1#)3 #2# joined the party line."
  MSG_PLLeave = "3*** #1# left the party line."
  MSG_PLLeaveMsg = "3*** #1# left the party line (#2#3)"
  MSG_PLBotNetLeave = "3*** 14(#1#)3 #2# left the party line."
  MSG_PLBotNetLeaveMsg = "3*** 14(#1#)3 #2# left the party line (#3#3)"
  MSG_PLTalk = "<#1#> #2#"
  MSG_PLBotTalk = "3*** 14(#1#)3 #2#"
  MSG_PLBotNetTalk = "14[#1#] <#2#> #3#"
  MSG_PLAct = "6* #1# #2#"
  MSG_PLBotNetAct = "14[#1#]6 * #2# #3#"
  MSG_PLNick = "3*** #1# is now known as #2#"
  MSG_PLAway = "3*** #1# is now away: #2#"
  MSG_PLBotNetAway = "3*** 14(#1#)3 #2# is now away: #3#"
  MSG_PLBack = "3*** #1# is back from: #2#"
  MSG_PLBotNetBack = "3*** 14(#1#)3 #2# is back from: #3#"
  MSG_PLNote = "2*** I've got 1 note waiting for you. Type .notes to get it."
  MSG_PLNotes = "2*** I've got #1# notes waiting for you. Type .notes to get them."
  MSG_PLHelpIntro = "*** DCC COMMANDS for #1#:"
  MSG_PLHelpFlagW = "For whatis authors:"
  MSG_PLHelpFlagO = "For channel ops:"
  MSG_PLHelpFlagT = "For botnet masters:"
  MSG_PLHelpFlagCM = "For channel masters:"
  MSG_PLHelpFlagM = "For masters:"
  MSG_PLHelpFlagCN = "For channel owners:"
  MSG_PLHelpFlagN = "For owners:"
  MSG_PLHelpFlagS = "For super owners:"
  MSG_PLHelpOutro = "All commands begin with '.', everything else goes to the party line.##Type '3.help <command>' for a description of each command."
  MSG_PLDCCIncoming = "#CLInfo#[#T#] *** DCC Chat from #1# (#2#) incoming..."
  MSG_PLDCCOpened = "#CLInfo#[#T#] *** DCC Chat with #1# opened"
  MSG_PLDCCClosed = "#CLInfo#[#T#] *** DCC Chat with #1# closed"
  MSG_PLDCCFirst = "#CLInfo#[#T#] *** DCC Chat with #1# opened - first login of this user"
  MSG_PLDCCWrongPW = "#CLInfo#[#T#] *** DCC Chat with #1# closed - wrong password"
  MSG_PLDCCFailed = "#CLInfo#[#T#] *** DCC Chat with #1# failed!"
  MSG_PLDCCRefused = "#CLInfo#[#T#] *** DCC Chat from #1# (#2#) refused"
  MSG_PLDCCNoAcc = "#CLInfo#[#T#] *** DCC Chat from #1# refused - no party line access"
  MSG_PLTelNetIncoming = "#CLInfo#[#T#] *** Telnet: Incoming connection from #1#"
  MSG_PLTelNetOpened = "#CLInfo#[#T#] *** Telnet: Connection with #1# opened"
  MSG_PLTelNetClosed = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed"
  MSG_PLTelNetFirst = "#CLInfo#[#T#] *** Telnet: Connection with #1# opened - first login of this user"
  MSG_PLTelNetNoUser = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed - unknown user"
  MSG_PLTelNetWrongPW = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed - wrong password"
  MSG_PLTelNetNoAcc = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed - no party line access"
  MSG_PLTelNetBot = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed - user is a bot (wrong port!)"
  MSG_PLBotNetIncoming = "#CLInfo#[#T#] *** Botnet: Incoming connection from #1#"
  MSG_PLBotNetLinking = "#CLInfo#[#T#] *** Linking to #1#: #2#"
  MSG_PLBotNetLinkFrom = "#CLInfo#[#T#] *** Link from #1#: #2#"
  MSG_PLBotNetError = "#CLInfo#[#T#] *** Error from #1#: #2#"
  MSG_PLBotNetLost = "#CLInfo#[#T#] *** Botnet: Connection to #1# lost"
  MSG_PLBotNetNoLink = "#CLInfo#[#T#] *** Couldn't link to #1#: Connection closed."
  MSG_PLBotNetConnect = "#CLInfo#[#T#] 3*** Connected to: #1#"
  MSG_PLIdentRequest = "3*** Ident request from #1#"
  MSG_PLFloodStart = "#CLInfo#[#T#] 2-=> Channel #1# is being flooded. Closing (#2#)..."
  MSG_PLFloodEnd = "#CLInfo#[#T#] 2-=> Flood in #1# seems to be over - channel is cleared."
  MSG_PLNickDid = "#CLInfo#[#T#] *** #1# did #2#"
  MSG_PLNickTried = "#CLInfo#[#T#] *** #1# tried #2#"
  MSG_PLNickFailed = "#CLInfo#[#T#] *** #1# failed #2#"
  MSG_PLWrongBot = "#CLInfo#[#T#] *** ERROR - Wrong nick! Wanted #1#, got #2#"
  MSG_PLYouWereBooted = "4*** You were booted off the party line by #1#"
  MSG_PLYouWereBooted2 = "4*** You were booted off the party line by #1# (#2#)"
  MSG_PLNickWasBooted = "3*** #1# was booted off the party line by #2#"
  MSG_PLNickWasBooted2 = "3*** #1# was booted off the party line by #2# (#3#)"
  
  '| Botnet messages
  ' -—————————- -- -  -
  MSG_BNBooted = "Booted by #1#"
  MSG_BNBooted2 = "Booted by #1#: #2#"
  MSG_BNSendUser = "Sending username"
  MSG_BNSendUserPass = "Sending user/pass"
  MSG_BNPingTimeout = "Ping timeout"
  MSG_BNNoPass = "ERROR - Bot wants a password but I don't have one!"
  MSG_BNBadPass = "ERROR - Bad password!"
  MSG_BNNoAccess = "ERROR - I don't have access!"
  MSG_BNLoop = "Loop detected (duplicate: #2#)... disconnecting"
  MSG_BNLeafLink = "Rejected leaf³Unauthorized link to #1#"
  MSG_BNLeafLinks = "Rejected leaf³Unauthorized links"
  MSG_BNBogusLink = "Rejected bot³Bogus link"
  MSG_BNRestructure = "Disconnected from³Restructure"
  MSG_BNConnect = "Connected to: #1#"
  MSG_BNDisconnect = "Disconnected from#1#"
  MSG_BNLostBot = "Lost Bot:"
  MSG_BNWrongBot = "Wrong nick:"
  
  '| Timed messages
  ' -—————————- -- -  -
  MSG_PLSwitchLogs1 = "#CLInfo#[#T#] *** Switching logs..."
  MSG_PLSwitchLogs2 = "#CLInfo#[#T#] *** Done."
  MSG_PLRequestedGo = "#CLInfo#[#T#] *** Requested GO #1# from #2#..."
  MSG_PLTelNetNickTO = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed - nick timeout"
  MSG_PLTelNetPassTO = "#CLInfo#[#T#] *** Telnet: Connection with #1# closed - password timeout"
  MSG_PLBotNetLinkTO = "#CLInfo#[#T#] *** Botnet: Connection with #1# timed out"
  MSG_PLBotNetConnTO = "#CLInfo#[#T#] *** Botnet: Connection with #1# timed out"
  MSG_PLBotNetNickTO = "#CLInfo#[#T#] *** Botnet: Connection with #1# closed - nick timeout"
  MSG_PLBotNetPassTO = "#CLInfo#[#T#] *** Botnet: Connection with #1# closed - password timeout"
  MSG_PLDCCGetTO = "#CLInfo#[#T#] *** DCC Get from #1# timed out (#2#)"
  MSG_PLDCCSendTO = "#CLInfo#[#T#] *** DCC send to #1# timed out (#2#)"
  MSG_PLDCCConnTO = "#CLInfo#[#T#] *** DCC connection with #1# timed out"
  
  '| IRC messages
  ' -—————————- -- -  -
  IRC_NoRepeats = "#1#: I don't like to repeat myself...|#1#: I already said that!|#1#: You already know the answer.|#1#: Look some lines above :)|#1#: I already answered to that!"
  IRC_Seen_NoNick = "#1#: Seen whom?!"
  IRC_Seen_ErrNick = "#1#: This nick isn't possible! ^_^"
  IRC_Whois_YUnknown = "#1#: Sorry, I don't know who you are."
  IRC_Whois_UUnknown = "Sorry, I don't know who #2# is."
  IRC_Whois_NoInfo = "#2# has no information set."
  IRC_Whatis_NotFound = "#1#: There's no description for '#2#'."
  
  '| Error messages
  ' -—————————- -- -  -
  ERR_Login_WrongPass = "Is it a bird? No! Is it a plane? No! It's a wrong password! :)"
  ERR_Login_Unknown = "Sorry, unknown user name."
  ERR_Login_NoChat = "Sorry, you're not allowed to chat with me."
  ERR_Pass_TooShort = "Your password is too short! Please use at least 6 characters."
  ERR_Pass_NoSpaces = "Please don't use spaces in your password. Try again."
  ERR_Pass_TooWeak = "This password is too weak. Please use a more intelligent one. :)"
  ERR_Nick_TooLong = "#CLError#*** Your nick can't be longer than #1# characters!"
  ERR_Nick_Erroneous = "#CLError#*** ""#1#"" Erroneous Nickname"
  ERR_Nick_InUse = "#CLError#*** ""#1#"" nickname is already in use."
  ERR_NotOnLocalChans = "#CLError#*** I couldn't find this user on local channels."
  ERR_UserNotFound = "#CLError#*** Sorry, I couldn't find this user."
  ERR_BotNotFound = "#CLError#*** Sorry, I couldn't find this bot."
  ERR_ServerLost = "#CLCritical#*** Connection to server lost!"
  ERR_ServerFailed = "#CLError#*** connect to #1# failed: #2#"
  ERR_NotAllowed = "#CLError#*** Sorry, you're not allowed to use this command!"
  ERR_CommandUsage = "#CLError#*** Usage: #1#"
  MSG_PL_LookHelp = "#CLError#*** What? Type '.help' to see the commands available."
  MSG_FA_LookHelp = "#CLError#*** What? Type 'help' to see the commands available."
  
  LanguageFile = GetPPString("Others", "Language", "", AnGeL_INI)
  If LanguageFile <> "" Then InitLanguage = LoadLanguage(LanguageFile) Else InitLanguage = True
End Function

Public Function LoadLanguage(LngName As String) As Boolean ' : AddStack "LanguageFile_LoadLanguage(" & LngName & ")"
Dim FNum As Integer, Line As String, SepPos As Long, SepPos2 As Long
Dim Section As String, Entry As String
  FNum = FreeFile
  On Local Error Resume Next
  Open FileAreaHome & "" & LngName & ".lng" For Input As #FNum
  If Err.Number <> 0 Then
    Close #FNum
    DeletePPString "Others", "Language", AnGeL_INI
    LoadLanguage = False
    Exit Function
  End If
    Do While Not EOF(FNum)
      Line Input #FNum, Line
      SepPos = InStr(Line, Chr(9))
      SepPos2 = InStr(Line, ":")
      If (SepPos > 0) And (SepPos2 > 0) Then
        Section = Left(Line, SepPos - 1)
        Entry = Mid(Line, SepPos2 + 1)
        Select Case Section
          '| Normal messages
          ' -—————————- -- -  -
          Case "MSG_EnterPWD": MSG_EnterPWD = Entry
          Case "MSG_ChoosePWD": MSG_ChoosePWD = Entry
          Case "MSG_ThanksPWD": MSG_ThanksPWD = Entry
          Case "MSG_CanYouSeeColors": MSG_CanYouSeeColors = Entry
          Case "MSG_CanYouSeeColTwo": MSG_CanYouSeeColTwo = Entry
          Case "MSG_ConnectedTo": MSG_ConnectedTo = Entry
          Case "MSG_FirstLogin": MSG_FirstLogin = Entry
          Case "MSG_SecurityInfo": MSG_SecurityInfo = Entry
          Case "MSG_ThankYou": MSG_ThankYou = Entry
          Case "MSG_ShortIntro": MSG_ShortIntro = Entry
          Case "MSG_PLJoin": MSG_PLJoin = Entry
          Case "MSG_PLJoinChan": MSG_PLJoinChan = Entry
          Case "MSG_PLBotNetJoin": MSG_PLBotNetJoin = Entry
          Case "MSG_PLLeave": MSG_PLLeave = Entry
          Case "MSG_PLLeaveMsg": MSG_PLLeaveMsg = Entry
          Case "MSG_PLBotNetLeave": MSG_PLBotNetLeave = Entry
          Case "MSG_PLBotNetLeaveMsg": MSG_PLBotNetLeaveMsg = Entry
          Case "MSG_PLTalk": MSG_PLTalk = Entry
          Case "MSG_PLBotTalk": MSG_PLBotTalk = Entry
          Case "MSG_PLBotNetTalk": MSG_PLBotNetTalk = Entry
          Case "MSG_PLAct": MSG_PLAct = Entry
          Case "MSG_PLBotNetAct": MSG_PLBotNetAct = Entry
          Case "MSG_PLNick": MSG_PLNick = Entry
          Case "MSG_PLAway": MSG_PLAway = Entry
          Case "MSG_PLBotNetAway": MSG_PLBotNetAway = Entry
          Case "MSG_PLBack": MSG_PLBack = Entry
          Case "MSG_PLBotNetBack": MSG_PLBotNetBack = Entry
          Case "MSG_PLNote": MSG_PLNote = Entry
          Case "MSG_PLNotes": MSG_PLNotes = Entry
          Case "MSG_PLHelpIntro": MSG_PLHelpIntro = Entry
          Case "MSG_PLHelpFlagW": MSG_PLHelpFlagW = Entry
          Case "MSG_PLHelpFlagO": MSG_PLHelpFlagO = Entry
          Case "MSG_PLHelpFlagT": MSG_PLHelpFlagT = Entry
          Case "MSG_PLHelpFlagCM": MSG_PLHelpFlagCM = Entry
          Case "MSG_PLHelpFlagM": MSG_PLHelpFlagM = Entry
          Case "MSG_PLHelpFlagCN": MSG_PLHelpFlagCN = Entry
          Case "MSG_PLHelpFlagN": MSG_PLHelpFlagN = Entry
          Case "MSG_PLHelpFlagS": MSG_PLHelpFlagS = Entry
          Case "MSG_PLHelpOutro": MSG_PLHelpOutro = Entry
          Case "MSG_PLDCCIncoming": MSG_PLDCCIncoming = Entry
          Case "MSG_PLDCCOpened": MSG_PLDCCOpened = Entry
          Case "MSG_PLDCCClosed": MSG_PLDCCClosed = Entry
          Case "MSG_PLDCCFirst": MSG_PLDCCFirst = Entry
          Case "MSG_PLDCCWrongPW": MSG_PLDCCWrongPW = Entry
          Case "MSG_PLDCCFailed": MSG_PLDCCFailed = Entry
          Case "MSG_PLDCCRefused": MSG_PLDCCRefused = Entry
          Case "MSG_PLDCCNoAcc": MSG_PLDCCNoAcc = Entry
          Case "MSG_PLTelNetIncoming": MSG_PLTelNetIncoming = Entry
          Case "MSG_PLTelNetOpened": MSG_PLTelNetOpened = Entry
          Case "MSG_PLTelNetClosed": MSG_PLTelNetClosed = Entry
          Case "MSG_PLTelNetFirst": MSG_PLTelNetFirst = Entry
          Case "MSG_PLTelNetNoUser": MSG_PLTelNetNoUser = Entry
          Case "MSG_PLTelNetWrongPW": MSG_PLTelNetWrongPW = Entry
          Case "MSG_PLTelNetNoAcc": MSG_PLTelNetNoAcc = Entry
          Case "MSG_PLTelNetBot": MSG_PLTelNetBot = Entry
          Case "MSG_PLBotNetIncoming": MSG_PLBotNetIncoming = Entry
          Case "MSG_PLBotNetLinking": MSG_PLBotNetLinking = Entry
          Case "MSG_PLBotNetLinkFrom": MSG_PLBotNetLinkFrom = Entry
          Case "MSG_PLBotNetError": MSG_PLBotNetError = Entry
          Case "MSG_PLBotNetLost": MSG_PLBotNetLost = Entry
          Case "MSG_PLBotNetNoLink": MSG_PLBotNetNoLink = Entry
          Case "MSG_PLBotNetConnect": MSG_PLBotNetConnect = Entry
          Case "MSG_PLIdentRequest": MSG_PLIdentRequest = Entry
          Case "MSG_PLFloodStart": MSG_PLFloodStart = Entry
          Case "MSG_PLFloodEnd": MSG_PLFloodEnd = Entry
          Case "MSG_PLNickDid": MSG_PLNickDid = Entry
          Case "MSG_PLNickTried": MSG_PLNickTried = Entry
          Case "MSG_PLNickFailed": MSG_PLNickFailed = Entry
          Case "MSG_PLWrongBot": MSG_PLWrongBot = Entry
          Case "MSG_PLYouWereBooted": MSG_PLYouWereBooted = Entry
          Case "MSG_PLYouWereBooted2": MSG_PLYouWereBooted2 = Entry
          Case "MSG_PLNickWasBooted": MSG_PLNickWasBooted = Entry
          Case "MSG_PLNickWasBooted2": MSG_PLNickWasBooted2 = Entry
            
          '| Botnet messages
          ' -—————————- -- -  -
          Case "MSG_BNBooted": MSG_BNBooted = Entry
          Case "MSG_BNBooted2": MSG_BNBooted2 = Entry
          Case "MSG_BNSendUser": MSG_BNSendUser = Entry
          Case "MSG_BNSendUserPass": MSG_BNSendUserPass = Entry
          Case "MSG_BNPingTimeout": MSG_BNPingTimeout = Entry
          Case "MSG_BNNoPass": MSG_BNNoPass = Entry
          Case "MSG_BNBadPass": MSG_BNBadPass = Entry
          Case "MSG_BNNoAccess": MSG_BNNoAccess = Entry
          Case "MSG_BNLoop": MSG_BNLoop = Entry
          Case "MSG_BNLeafLink": MSG_BNLeafLink = Entry
          Case "MSG_BNLeafLinks": MSG_BNLeafLinks = Entry
          Case "MSG_BNBogusLink": MSG_BNBogusLink = Entry
          Case "MSG_BNRestructure": MSG_BNRestructure = Entry
          Case "MSG_BNConnect": MSG_BNConnect = Entry
          Case "MSG_BNDisconnect": MSG_BNDisconnect = Entry
          Case "MSG_BNLostBot": MSG_BNLostBot = Entry
          Case "MSG_BNWrongBot": MSG_BNWrongBot = Entry
          
          '| Timed messages
          ' -—————————- -- -  -
          Case "MSG_PLSwitchLogs1": MSG_PLSwitchLogs1 = Entry
          Case "MSG_PLSwitchLogs2": MSG_PLSwitchLogs2 = Entry
          Case "MSG_PLRequestedGO": MSG_PLRequestedGo = Entry
          Case "MSG_PLTelNetNickTO": MSG_PLTelNetNickTO = Entry
          Case "MSG_PLTelNetPassTO": MSG_PLTelNetPassTO = Entry
          Case "MSG_PLBotNetLinkTO": MSG_PLBotNetLinkTO = Entry
          Case "MSG_PLBotNetConnTO": MSG_PLBotNetConnTO = Entry
          Case "MSG_PLBotNetNickTO": MSG_PLBotNetNickTO = Entry
          Case "MSG_PLBotNetPassTO": MSG_PLBotNetPassTO = Entry
          Case "MSG_PLDCCGetTO": MSG_PLDCCGetTO = Entry
          Case "MSG_PLDCCSendTO": MSG_PLDCCSendTO = Entry
          Case "MSG_PLDCCConnTO": MSG_PLDCCConnTO = Entry
          
          '| IRC messages
          ' -—————————- -- -  -
          Case "IRC_NoRepeats": IRC_NoRepeats = Entry
          Case "IRC_Seen_NoNick": IRC_Seen_NoNick = Entry
          Case "IRC_Seen_ErrNick": IRC_Seen_ErrNick = Entry
          Case "IRC_Whois_YUnknown": IRC_Whois_YUnknown = Entry
          Case "IRC_Whois_UUnknown": IRC_Whois_UUnknown = Entry
          Case "IRC_Whois_NoInfo": IRC_Whois_NoInfo = Entry
          Case "IRC_Whatis_NotFound": IRC_Whatis_NotFound = Entry
          
          '| Error messages
          ' -—————————- -- -  -
          Case "ERR_Login_WrongPass": ERR_Login_WrongPass = Entry
          Case "ERR_Login_Unknown": ERR_Login_Unknown = Entry
          Case "ERR_Login_NoChat": ERR_Login_NoChat = Entry
          Case "ERR_Pass_TooShort": ERR_Pass_TooShort = Entry
          Case "ERR_Pass_NoSpaces": ERR_Pass_NoSpaces = Entry
          Case "ERR_Pass_TooWeak": ERR_Pass_TooWeak = Entry
          Case "ERR_Nick_TooLong": ERR_Nick_TooLong = Entry
          Case "ERR_Nick_Erroneous": ERR_Nick_Erroneous = Entry
          Case "ERR_Nick_InUse": ERR_Nick_InUse = Entry
          Case "ERR_NotOnLocalChans": ERR_NotOnLocalChans = Entry
          Case "ERR_UserNotFound": ERR_UserNotFound = Entry
          Case "ERR_BotNotFound": ERR_BotNotFound = Entry
          Case "ERR_ServerLost": ERR_ServerLost = Entry
          Case "ERR_ServerFailed": ERR_ServerFailed = Entry
          Case "ERR_CommandUsage": ERR_CommandUsage = Entry
          Case "MSG_PL_LookHelp": MSG_PL_LookHelp = Entry
          Case "MSG_FA_LookHelp": MSG_FA_LookHelp = Entry
        End Select
      End If
    Loop
  Close #FNum
  LoadLanguage = True
End Function

Public Function MakeMsg(Line As String, ParamArray args()) As String ' : AddStack "LanguageFile_MakeMsg(" & Line & ")"
Dim u As Long, ResLine As String
  ResLine = Line
  For u = 1 To UBound(args) + 1
    ResLine = Replace(ResLine, "#" & CStr(u) & "#", (args(u - 1)))
  Next u
  ResLine = Replace(ResLine, "##", vbCrLf)
  ResLine = Replace(ResLine, "#T#", Time)
  ResLine = Replace(ResLine, "#CLInfo#", "14")
  ResLine = Replace(ResLine, "#CLError#", "5")
  ResLine = Replace(ResLine, "#CLCritical#", "4")
  MakeMsg = ResLine
End Function
