Attribute VB_Name = "Partyline_CommandList"
',-======================- ==-- -  -
'|   AnGeL - Partyline - CommandList
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

Public Const Cl_User As Integer = 1
Public Const Cl_What As Integer = 2
Public Const Cl_Op   As Integer = 4
Public Const Cl_Net  As Integer = 8
Public Const Cl_CMas As Integer = 16
Public Const Cl_Mas  As Integer = 32
Public Const Cl_COwn As Integer = 64
Public Const Cl_Own  As Integer = 128
Public Const Cl_SOwn As Integer = 256


Private Type Command
  Name As String
  Class As Integer
  MatchFlags As String
  Script As String
  SpreadTo As String
  Description As String
End Type


Public Const CC_NoCommand As Long = 0
Public Const CC_NotAllowed As Long = -1


Public Commands() As Command
Public CommandCount As Long


Sub Commands_Load()
  ReDim Preserve Commands(5)
End Sub

Sub Commands_Unload()
'
End Sub

Public Sub AddCommand(Name As String, Class As Integer, MatchFlags As String, Script As String, SpreadTo As String, Description As String) ' : AddStack "Help_AddCommand(" & Name & ", " & Class & ", " & MatchFlags & ", " & Script & ", " & SpreadTo & ", " & Description & ")"
Dim i As Long
  'For i = 1 To CommandCount
  '  If LCase(Commands(i).Name) = LCase(Name) Then SpreadFlagMessage 0, "+m", "14*** AddCommand '" & Name & "' failed: Command is already existing.": Exit Sub
  'Next i
  CommandCount = CommandCount + 1
  If CommandCount > UBound(Commands()) Then ReDim Preserve Commands(((CommandCount \ 5) + 1) * 5)
  Commands(CommandCount).Name = Name
  Commands(CommandCount).Class = Class
  Commands(CommandCount).Script = Script
  Commands(CommandCount).SpreadTo = SpreadTo
  Commands(CommandCount).MatchFlags = MatchFlags
  Commands(CommandCount).Description = Description
End Sub

Public Sub RemCommand(Name As String) ' : AddStack "Help_RemCommand(" & Name & ")"
Dim i As Long, i2 As Long
  For i = 1 To CommandCount
    If LCase(Param(Commands(i).Name, 1)) = LCase(Param(Name, 1)) Then
      For i2 = i To CommandCount - 1
        Commands(i2) = Commands(i2 + 1)
      Next i2
      CommandCount = CommandCount - 1
      ReDim Preserve Commands(((CommandCount \ 5) + 1) * 5)
      Exit For
    End If
  Next i
End Sub

'Checks whether a command is available for party line user <vsock>
Function CheckCommand(vsock As Long, CommandName As String) As Long ' : AddStack "Routines_CheckCommand(" & vsock & ", " & CommandName & ")"
Dim i As Long, i2 As Long, CName As String, CLev As Integer
  CName = LCase(CommandName)
  CLev = CLevel(SocketItem(vsock).Flags, AllFlags(SocketItem(vsock).RegNick))
  For i = 1 To CommandCount
    For i2 = 1 To ParamCount(Commands(i).Name)
      If LCase(Param(Commands(i).Name, i2)) = CName Then
        If (Commands(i).MatchFlags = "") Or ((Commands(i).MatchFlags <> "") And (MatchUserFlags(SocketItem(vsock).UserNum, Commands(i).MatchFlags) = True)) Then
          If (CLev And Commands(i).Class) > 0 Then
            CheckCommand = i
            Exit Function
          Else
            CheckCommand = CC_NotAllowed
          End If
        Else
          CheckCommand = CC_NotAllowed
        End If
      End If
    Next i2
  Next i
  If CheckCommand = CC_NotAllowed Then Exit Function
  CheckCommand = CC_NoCommand
End Function

'Checks for match in commands
Function MatchCommand(CommandName As String) As String
  Dim i As Long, i2 As Long, CMatch As Integer, CName As String
  CMatch = 0
  For i = 1 To CommandCount
    For i2 = 1 To ParamCount(Commands(i).Name)
      If LCase(Param(Commands(i).Name, i2)) Like LCase(CommandName) & "*" Then
        CName = LCase(Param(Commands(i).Name, i2))
        CMatch = CMatch + 1
      End If
    Next i2
  Next i
  If CMatch = 1 Then MatchCommand = CName Else MatchCommand = ""
End Function

Public Sub InitHelp()
  Dim TempStr As String
  AddCommand "xpfix", Cl_SOwn, "", "", "", ""
  AddCommand "iscripts", Cl_SOwn, "", "", "", ""
  AddCommand "realname newident botnetnick primarynick prinick secondarynick secnick ident killprot", Cl_User, "", "", "", ""
  AddCommand "save", Cl_User, "", "", "", ""
  AddCommand "userport up botport bp", Cl_Net, "", "", "", ""
  'For all users ----------
  AddCommand "who", Cl_User, "", "", "+m", "2*** WHO" & _
      "##14  Zeigt eine Liste aller Personen an, die sich gerade" & _
      "##14  auf meinem Teil der Partyline aufhalten. Das Botnet" & _
      "##14  wird hierbei nicht beachtet." & _
      "##14  WHO kann auch benutzt werden, um den Status eines" & _
      "##14  Bots im Botnet abzufragen: .who <bot>" & _
      "##14  -> siehe auch 'whom'"
  AddCommand "whom", Cl_User, "", "", "+m", "2*** WHOM" & _
      "##14  Zeigt eine Liste aller Personen an, die sich gerade" & _
      "##14  auf den Partylines aller Bots des Botnets aufhalten." & _
      "##14  -> siehe auch 'who'"
  AddCommand "me", Cl_User, "", "", "", "2*** ME <action>" & _
      "##14  Beschreibt, was Du gerade machst. Genau wie '/me' in mIRC."
  AddCommand "msg", Cl_User, "", "", "", "2*** MSG <nick> <message>" & _
      "##14  Schickt dem User <nick> auf der Partyline die private" & _
      "##14  Message <message>. Diese kann kein anderer lesen."
  AddCommand "away", Cl_User, "", "", "", "2*** AWAY (reason)" & _
      "##14  Wenn <reason> angegeben wird, setzt dieses Kommando" & _
      "##14  dich mit dem angegebenen Grund auf away. Wenn keine" & _
      "##14  Parameter angegeben werden, wird das away gelöscht."
  AddCommand "back", Cl_User, "", "", "", "2*** BACK" & _
      "##14  Wenn du als away markiert warst, wird dies hiermit" & _
      "##14  gelöscht. Falls du Notes bekommen hast, wird eine" & _
      "##14  entsprechende Nachricht angezeigt."
  AddCommand "quit q", Cl_User, "", "", "", "2*** QUIT, Q <message>" & _
      "##14  Beendet die DCC Chat - Verbindung zur Partyline."
  AddCommand "newpass", Cl_User, "", "", "", "2*** NEWPASS <new password>" & _
      "##14  Ändert Dein Passwort. Es muß mindestens 6 Zeichen lang sein."
  AddCommand "colors color colour", Cl_User, "", "", "+m", "2*** COLORS <on/off>" & _
      "##14  Schaltet die Anzeige von Farben auf meiner Partyline an oder aus."
  AddCommand "setup console", Cl_User, "", "", "+m", "2*** SETUP" & _
      "##14  Ruft ein Menü auf, in dem du bestimmen kannst, welche" & _
      "##14  Meldungen der Bot dir auf der Partyline anzeigt."
  AddCommand "whois", Cl_User, "", "", "+m", "2*** WHOIS <user>" & _
      "##14  Zeigt detaillierte Informationen über einen User an. Es werden" & _
      "##14  unter anderem auch die ""Flags"" des Users angezeigt." & _
      "##14  Hier eine Liste, was die Buchstaben bedeuten:" & _
      "##" & EmptyLine + _
      "##1  Flag  Bedeutung" & _
      "##1  ----  -----------------------------------------" & _
      "##2  a     (A)uto-Op" & _
      "##2  b     User ist ein (B)ot" & "##2  d     Auto-(D)eop" & _
      "##2  f     (F)riend" & "##2  i     F(i)le area access" & _
      "##2  j     File area (j)anitor" & "##2  k     Auto-(K)ick" & "##2  m     User ist ein (M)aster" & _
      "##2  n     User ist ein Ow(n)er" & "##2  o     (o)p on request" & "##2  p     (P)arty line access" & _
      "##2  r     Special (R)evenge" & "##2  s     (S)uper owner" & "##2  t     Botne(T) Master" & "##2  v     Auto-(V)oice" & _
      "##2  w     (W)hatis author" & "##2  x     E(x)tra - Nick protection" & _
      "##14  ------------------ 1BOT FLAGS14 ------------------" & _
      "##2  a     (A)lternative Hub" & _
      "##2  h     (H)ub Bot - AutoLink" & "##" & _
      "##2  l     (L)eaf Bot - may not link other bots" & "##" & _
      "##2  s     (S)hared Bot - mixes userfiles with me" & "##" & EmptyLine
  AddCommand "nick", Cl_User, "", "", "", "2*** NICK <new nick>" & _
      "##14  Ändert Deinen Nick auf der Partyline. Dieser ist unabhängig" & _
      "##14  von dem Nickname, den Du gerade im IRC hast."
  AddCommand "help", Cl_User, "", "", "+m", "2*** HELP##14  Rate mal! ;o))"
  AddCommand "urls", Cl_User, "", "", "+m", "2*** URLS" & _
      "##14  Zeigt eine Liste von interessanten Web-Adressen an."
  AddCommand "note", Cl_User, "", "", "", "2*** NOTE <nick(@bot)> <message>" & _
      "##14  Schickt dem User <nick> die Nachricht <message> zu." & _
      "##14  Falls der Zusatz <@bot> angegeben wird, wird alles" & _
      "##14  an den angegebenen Bot geschickt. Dieser sollte" & _
      "##14  antworten, ob die Note zugestellt werden konnte."
  AddCommand "fwd", Cl_User, "", "", "", "2*** FWD <nick@bot>" & _
      "##14  Richtet eine Weiterleitung für Notes ein."
  AddCommand "notes", Cl_User, "", "", "+m", "2*** NOTES (read / erase)" & _
      "##14  Ohne Parameter zeigt dieser Befehl alle Notes an," & _
      "##14  die Dir geschickt wurden, und löscht sie danach." & _
      "##14  Wird der Parameter 'read' übergeben, werden deine" & _
      "##14  Notes angezeigt, aber nicht gelöscht. Bei 'erase'" & _
      "##14  werden die Notes ohne vorherige Anzeige gelöscht."
  AddCommand "info", Cl_User, "", "", "+m", "2*** INFO <your info line>" & _
      "##14  Ändert die Information, die ich anzeige, wenn" & _
      "##14  jemand in einem Raum !whois <dein name>" & _
      "##14  eingibt."
  AddCommand "-host", Cl_User, "-mt", "", "", "2*** -HOST (your nick) <hostmask>" & _
      "##14  Entfernt eine deiner Hostmasks aus meiner Liste."
  AddCommand "seen", Cl_User, "", "", "+m", "2*** SEEN <nick/hostmask>" & _
      "##14  Zeigt an, wann ich den User mit dem Namen <nick>" & _
      "##14  das letzte Mal gesehen habe. Wenn eine Hostmask" & _
      "##14  angegeben wird, gebe ich den ersten User aus, der" & _
      "##14  zu dieser Hostmask passt. Es können auch mehrere" & _
      "##14  User angegeben werden, die durch Kommata getrennt" & _
      "##14  werden müssen." & _
      "##" & EmptyLine + _
      "##14  Beispiel: .seen YarA, AlexKid, HotaruT" & _
      "##" & EmptyLine + _
      "##14  Tip: Du kannst dem Bot per DCC eine Datei mit der" & _
      "##14       Endung '.seen' schicken. In die erste Zeile" & _
      "##14       dieser Datei kannst du bis zu 14 Nicks mit" & _
      "##14       Leerzeichen dazwischen schreiben. Wenn die" & _
      "##14       Datei z.B. 'Friends.seen' heißt, werden dann" & _
      "##14       bei der Eingabe von '" & CommandPrefix & "seen Friends' die" & _
      "##14       Seen-Zeiten der angegebenen User gezeigt."
  AddCommand "uptime", Cl_User, "", "", "+m", "2*** UPTIME" & _
      "##14  Zeigt an, wie lange ich schon zum IRC verbunden" & _
      "##14  bin, und wie lange ich auf meinem Host-Rechner" & _
      "##14  bereits laufe."
  TempStr = "2*** MOTD (bot)" & _
      "##14  Zeigt die 'Message Of The Day' an, die auch beim" & "##14  Login auf der Partyline angezeigt wird. Wenn der" & "##14  Parameter <bot> angegeben wird, schickt dieser" & "##14  Bot des Botnets seine MOTD zu dir." & "##" & EmptyLine + _
      "##14  Als Owner kannst du die MOTD auch verändern, indem" & "##14  du dem Bot per DCC eine neue 'MOTD.txt' schickst." & "##14  Wenn sie 0 Byte groß ist, wird die MOTD gelöscht." & "##" & EmptyLine + _
      "##14  Die MOTD-Datei ist eine ganz normale Textdatei," & "##14  die der Bot jedem User beim Login anzeigt. Es" & "##14  können neben Text auch folgende Variablen in die" & "##14  Datei eingebaut werden:" & "##" & EmptyLine
  TempStr = TempStr + _
      "##1  Variable  Dafür eingesetzter Text" & "##1  --------  -----------------------------------------" & "##2  %B        Name des Bots" & "##2  %C        Channels, auf denen der Bot ist" & "##2  %N        Name des einloggenden Users" & "##2  %T        Aktuelle Zeit (z.B. '17:29')" & "##2  %V        Bot-Version (z.B. 'AnGeL " & BotVersion & "')" & "##" & EmptyLine + _
      "##14  Es kann auch festgelegt werden, daß nur User mit" & "##14  bestimmten Flags Zeilen der MOTD lesen können." & "##14  Dies geschieht durch die Angabe '%{<flags>}' in" & "##14  einer Zeile." & "##" & EmptyLine & "##14  Unsere Beispiel-MOTD-Datei sieht so aus:" & _
      "##14    Hey %N! Ich bin %B, ein %V." & "##14    %{+t-n}" & "##14    Zeilen für Botnet-Master, die aber keine Owner" & "##14    des Bots sind." & "##14    %{+n}"
  TempStr = TempStr + _
      "##14    Nur Bot-Owner können diese Zeile sehen." & "##14    %{+}" & "##14    Zeile für alle." & "##14  Ein User namens 'PowerGirl' würde dies sehen:" & "##14    Hey PowerGirl! Ich bin " & BotNetNick & ", ein AnGeL " & BotVersion & "." & "##14    Zeile für alle." & "##14  Ein Botnet-Master 'Kalyx' würde dies sehen:" & "##14    Hey Kalyx! Ich bin " & BotNetNick & ", ein AnGeL " & BotVersion & "." & _
      "##14    Zeilen für Botnet-Master, die aber keine Owner" & "##14    des Bots sind." & "##14    Zeile für alle." & "##14  Ein Owner 'Hippo' würde beim Login dies sehen:" & "##14    Hey Hippo! Ich bin " & BotNetNick & ", ein AnGeL " & BotVersion & "." & _
      "##14    Nur Bot-Owner können diese Zeile sehen." & "##14    Zeile für alle." & "##" & EmptyLine
  AddCommand "motd", Cl_User, "", "", "+m", TempStr
  AddCommand "whatis", Cl_User, "", "", "+m", "2*** WHATIS <item>" & _
      "##14  Sucht in der " & CommandPrefix & "whatis - Datenbank nach einem Eintrag," & _
      "##14  der <item> ungefähr entspricht. Falls mehrere Einträge" & _
      "##14  gefunden werden, zeige ich die Fundstellen an."
  AddCommand "whatsnew", Cl_User, "", "", "+m", "2*** WHATSNEW" & _
      "##14  Zeigt eine Liste der Änderungen dieser AnGeL-Version."
  AddCommand "files", Cl_User, "", "", "+m", "2*** FILES" & _
      "##14  Damit verläßt Du die Partyline und kommst in meine" & _
      "##14  File-Area, aus der man Dateien herunterladen kann."
  AddCommand "switchuser su", Cl_User, "", "", "+m", "2*** SWITCHUSER" & _
      "##14  Damit verläßt Du die Partyline und landest auf einem" & _
      "##14  Login-Screen auf dem du dich neu einloggen kannst."
  AddCommand "relay", Cl_User, "", "", "+m", "2*** Relay" & _
      "##14  Erlaubt dir eine Verbindung zu einem anderen Bot" & _
      "##14  von mir aus aufzubauen."
  'For whatis authors ----------
  AddCommand "wset", Cl_What, "", "", "+w", "2*** WSET <item> (description)" & _
      "##14  Setzt die Beschreibung des " & CommandPrefix & "whatis - Eintrags <item>" & _
      "##14  auf <description>. Wenn <description> nicht angegeben" & _
      "##14  wird, dann wird der " & CommandPrefix & "whatis - Eintrag gelöscht." & _
      "##" & EmptyLine + _
      "##14  In <item> können auch Leerzeichen enthalten sein;" & _
      "##14  dazu muß die Schreibweise ""<item>"" benutzt werden." & _
      "##" & EmptyLine + _
      "##14  Beispiel: .wset ""kleiner test"" Dies ist ein Test!"
  AddCommand "wlist", Cl_What, "", "", "+w", "2*** WLIST" & _
      "##14  Zeigt eine Liste der gesetzten " & CommandPrefix & "whatis - Einträge an."
  'For channel ops ----------
  AddCommand "op", Cl_Op, "", "", "+m", "2*** OP <nick> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Bringt mich dazu, dem User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  OP-Status zu geben. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann oppe ich ihn in jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)." & _
      "##14  -> siehe auch 'deop'"
  AddCommand "deop", Cl_Op, "", "", "+m", "2*** DEOP <nick> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Bringt mich dazu, dem User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  den OP-Status zu nehmen. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann deoppe ich ihn in jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)." & _
      "##14  -> siehe auch 'op'"
  AddCommand "kick k", Cl_Op, "", "", "+m", "2*** KICK, K <nick> ([" & ServerChannelPrefixes & "]channel) (reason)" & _
      "##14  Mit diesem Befehl kicke ich den User <nick> aus dem Raum" & _
      "##14  <[" & ServerChannelPrefixes & "]channel>. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann kicke ich ihn aus jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)." & _
      "##14  Du kannst optional als <reason> noch einen Grund für" & _
      "##14  den Kick angeben." & _
      "##14  -> siehe auch 'kickban'"
  AddCommand "kickban kb", Cl_Op, "", "", "+m", "2*** KICKBAN, KB <nick> ([" & ServerChannelPrefixes & "]channel) (reason)" & _
      "##14  Ich banne hiermit den User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  und kicke ihn dann. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann kickbanne ich ihn überall," & _
      "##14  wo Du bei mir Rechte hast (flag +o bei .whois)." & _
      "##14  Du kannst optional als <reason> noch einen Grund für" & _
      "##14  den Kickban angeben." & _
      "##14  -> siehe auch 'kick'"
  AddCommand "voice", Cl_Op, "", "", "+m", "2*** VOICE <nick> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Bringt mich dazu, dem User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Voice (+v) zu geben. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann voice ich ihn in jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)."
  AddCommand "devoice", Cl_Op, "", "", "+m", "2*** DEVOICE <nick> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Bringt mich dazu, dem User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  die Voice zu nehmen. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann devoice ich ihn in jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)."
  AddCommand "hop", Cl_Op, "", "", "+m", "2*** HOP <nick> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Bringt mich dazu, dem User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  HalfOp (+h) zu geben. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann halfope ich ihn in jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)."
  AddCommand "dehop", Cl_Op, "", "", "+m", "2*** DEHOP <nick> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Bringt mich dazu, dem User <nick> im Raum <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  den HalfOp zu nehmen. Die <[" & ServerChannelPrefixes & "]channel>-Angabe kann auch" & _
      "##14  weggelassen werden; dann dehalfope ich ihn in jedem Raum," & _
      "##14  für den Du bei mir Rechte hast (flag +o bei .whois)."
  AddCommand "invite", Cl_Op, "", "", "", "2*** INVITE <nick> <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Lädt den User <nick> in den Raum <[" & ServerChannelPrefixes & "]channel> ein."
  AddCommand "key", Cl_Op, "", "", "", "2*** KEY <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Zeigt Dir das Passwort eines Raumes, in dem ich bin."
  AddCommand "channel ch", Cl_Op, "", "", "+m", "2*** CHANNEL, CH <[" & ServerChannelPrefixes & "]channel> (+)" & _
      "##14  Zeigt Informationen über den Raum <[" & ServerChannelPrefixes & "]channel> an, so" & _
      "##14  z.B. die aktuelle Useranzahl, das Topic, die Modes" & _
      "##14  und eine Liste aller Chatter im Raum mit den Namen," & _
      "##14  unter denen ich sie erkannt habe." & _
      "##14  Wird ein + hinter dem [" & ServerChannelPrefixes & "]channel angegeben, so zeige" & _
      "##14  ich anstelle der Hosts die IP's der User an."
  AddCommand "chanbans cb", Cl_Op, "", "", "+m", "2*** CHANBANS, CB <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Zeigt alle Bans des Raum <[" & ServerChannelPrefixes & "]channel> an."
  'For botnet masters ----------
  AddCommand "+bot", Cl_Net, "", "", "", "2*** +BOT <nick> (address:port) (+botflags)" & _
      "##14  Fügt den Bot <nick> zu meiner Userliste hinzu und" & _
      "##14  gibt ihm die Flags '+bf'. Wenn der Bot sich in einem" & _
      "##14  Raum aufhält, in dem ich auch bin, trage ich die" & _
      "##14  Hostmask des Bots automatisch mit dazu. Die Angabe" & _
      "##14  (adress:port) legt fest, wohin ich connecten muß, um" & _
      "##14  den Bot zu erreichen und eine Botnet-Verbindung mit" & _
      "##14  ihm aufzubauen. Bitte hier keine Adresse der Form" & _
      "##14  '*!*bot@abc.de' angeben, sondern nur Host und Port!" & _
      "##14  Es können auch gleich Botflags angegeben werden," & _
      "##14  z.B. '+h' für einen Hub Bot (siehe '.help whois')." & _
      "##" & EmptyLine + _
      "##14  Beispiel: .+bot Alexia fireworks.com:3035 +h" & _
      "##14  -> siehe auch '-bot', 'botport', 'chaddr'"
  AddCommand "-bot", Cl_Net, "", "", "", "2*** -BOT <nick>" & _
      "##14  Entfernt den Bot <nick> aus meiner Userliste."
  AddCommand "botattr", Cl_Net, "", "", "+t", "2*** BOTATTR <bot> <+/-flags>" & _
      "##14  Ändert die Bot-Flags des Bots <bot> bei mir. Bitte" & _
      "##14  nicht die Bot-Flags mit den Flags, die durch .chattr" & _
      "##14  geändert werden können, verwechseln!" & _
      "##14  Eine Liste aller Flags und ihre jeweilige Bedeutung" & _
      "##14  kannst Du mit '.help whois' bekommen."
  AddCommand "chaddr", Cl_Net, "", "", "+t", "2*** CHADDR <nick> <address:port>" & _
      "##14  Ändert die connect-Adresse des Bots <nick>. Das" & _
      "##14  <adress:port> legt fest, wohin ich connecten muß, um" & _
      "##14  den Bot zu erreichen und eine Botnet-Verbindung mit" & _
      "##14  ihm aufzubauen. Bitte hier keine Adresse der Form" & _
      "##14  '*!*bot@abc.de' angeben, sondern nur Host und Port!" & _
      "##" & EmptyLine + _
      "##14  Beispiel: .chaddr Alexia fireworks.com:3035" & _
      "##14  -> siehe auch '+bot', 'botport'"
  AddCommand "-host", Cl_Net, "-m+t", "", "", "2*** -HOST (bot nick / your nick) <hostmask>" & _
      "##14  Entfernt die Hostmask eines Bots oder eine" & _
      "##14  deiner Hostmasks aus meiner Liste."
  AddCommand "comment", Cl_Net, "", "", "+m", "2*** COMMENT <user> <comment / ban reason>" & _
      "##14  Ändert die Kommentar-Zeile eines Users. Diese wird" & _
      "##14  angezeigt, wenn ein User durch das Flag +k von mir" & _
      "##14  gekickt wird. Sie wird auch bei '.whois <user>' mit" & _
      "##14  aufgeführt. Bei Bots hat die Zeile keine besondere" & _
      "##14  Funktion und kann z.B. für Infos genutzt werden."
  AddCommand "link", Cl_Net, "", "", "+t", "2*** LINK <nick>" & _
      "##14  Bringt mich dazu, eine Botnet-Verbindung zum Bot <nick>" & _
      "##14  aufzubauen. Dazu muß der Bot bei mir eingetragen sein" & _
      "##14  und ich muß seine connect-Adresse kennen." & _
      "##14  -> siehe auch '+bot', 'chaddr', 'botport'"
  AddCommand "unlink", Cl_Net, "", "", "+t", "2*** UNLINK <nick>/*" & _
      "##14  Beendet die Botnet-Verbindung mit dem Bot <nick>." & _
      "##14  Wenn für <nick> ein Sternchen angegeben wird, also" & _
      "##14  '.unlink *', beende ich die Verbindung zu allen Bots." & _
      "##14  -> siehe auch 'link'"
  AddCommand "chpass", Cl_Net, "", "", "", "2*** CHPASS <user/bot> <new password>" & _
      "##14  Ändert das Passwort des angegebenen Users/Bots."
  AddCommand "bottree bt botree vbottree vbt", Cl_Net, "", "", "+t", "2*** (V)BOTTREE, (V)BT" & _
      "##14  Zeigt das aktuelle Botnet in einer Baumstruktur an," & _
      "##14  in der man sehen kann, wie die einzelnen Bots des" & _
      "##14  Netzes miteinander verbunden sind und wieviele Bots" & _
      "##14  sich insgesamt im Botnet befinden." & _
      "##14  Mit '.vbottree' werden die Bot-Versionen angezeigt."
  AddCommand "botinfo", Cl_Net, "", "", "+t", "2*** BOTINFO" & _
      "##14  Veranlasst alle Bots des Botnets dazu, je eine Zeile" & _
      "##14  zu senden, die Informationen über die Version, Räume" & _
      "##14  und Laufzeit des Bots enthält." & _
      "##14  -> siehe 'chanlist','uptime'"
  AddCommand "trace", Cl_Net, "", "", "+t", "2*** TRACE" & _
      "##14  Startet eine Routenverfolgung eines Datenpakets durch" & _
      "##14  das BotNet bis zu einem speziellen Bot. Man sollte " & _
      "##14  eine Info über Weg und Zeit des Paketes bekommen."
  'For masters ----------
  AddCommand "match lmatch", Cl_Mas, "+m", "", "+m", "2*** (L)MATCH (hostmask) (+/-<flags> ([" & ServerChannelPrefixes & "]channel)) (+/-<botflags> bot)" & _
      "##14  Sucht aus meiner Userliste alle User heraus, die##14  bestimmten Kriterien entsprechen. So kann z.B. nach" & _
      "##14  einer Hostmask gesucht werden, nach einem Namen##14  oder Namensteil, und nach Flags (global oder lokal)." & _
      "##" & EmptyLine + _
      "##1  Parameter:      Gefundene User:" & _
      "##1  --------------  --------------------------------------------------" & _
      "##2  Slr*            14User, deren Nick im Bot mit 'Slr' anfängt" & _
      "##2  *!*Hippo@*.de   14User mit dieser oder einer passenden Hostmask" & _
      "##2  +n              14Globale Owner" & _
      "##2  +h bot          14Hub Bots" & _
      "##2  +n #            14Globale Owner und Channelowner jeglicher Channels" & _
      "##2  -n +n #abc      14Keine globalen Owner, nur Channelowner von #abc" & _
      "##2  +bo #flirt      14Bots mit Op in #flirt" & _
      "##2  +o #fu*         14Ops eines Channels, dessen Name mit '#fu' beginnt" & _
      "##2  +m-o            14Master ohne globalen Op" & _
      "##" & EmptyLine
  AddCommand "match lmatch", Cl_COwn, "-m", "", "+m", "2*** (L)MATCH (hostmask) (+/-<flags> ([" & ServerChannelPrefixes & "]channel)) (+/-<botflags> bot)" & _
      "##14  Sucht aus meiner Userliste alle User heraus, die##14  bestimmten Kriterien entsprechen. So kann z.B. nach" & _
      "##14  einer Hostmask gesucht werden, nach einem Namen##14  oder Namensteil, und nach Flags (global oder lokal)." & _
      "##" & EmptyLine + _
      "##1  Parameter:      Gefundene User:" & _
      "##1  --------------  --------------------------------------------------" & _
      "##2  Slr*            14User, deren Nick im Bot mit 'Slr' anfängt" & _
      "##2  *!*Hippo@*.de   14User mit dieser oder einer passenden Hostmask" & _
      "##2  +n              14Globale Owner" & _
      "##2  +h bot          14Hub Bots" & _
      "##2  +n #            14Globale Owner und Channelowner jeglicher Channels" & _
      "##2  -n +n #abc      14Keine globalen Owner, nur Channelowner von #abc" & _
      "##2  +bo #flirt      14Bots mit Op in #flirt" & _
      "##2  +o #fu*         14Ops eines Channels, dessen Name mit '#fu' beginnt" & _
      "##2  +m-o            14Master ohne globalen Op" & _
      "##" & EmptyLine
  AddCommand "status stat st", Cl_Mas, "", "", "+m", "2*** STATUS, ST" & _
      "##14  Zeigt einen umfassenden Bericht über den aktuellen" & _
      "##14  Status des Bots an; darunter Informationen über" & _
      "##14  Server, Hostmask, Uptime, Channels etc."
  AddCommand "flagnote", Cl_Mas, "", "", "", "2*** FLAGNOTE <+flags> ([" & ServerChannelPrefixes & "]channel) <message>" & _
      "##14  Schickt allen Usern, die die Flags <+flags> haben," & _
      "##14  die Nachricht <message>. Falls <[" & ServerChannelPrefixes & "]channel> angegeben" & _
      "##14  wird, bekommen die User, die im Raum <[" & ServerChannelPrefixes & "]channel> die" & _
      "##14  Flags <+flags> haben, die Nachricht geschickt."
  AddCommand "chanlist cl", Cl_Mas, "", "", "+m", "2*** CHANLIST, CL" & _
      "##14  Gibt eine Liste aller permanenten Channels" & _
      "##14  aus, die ich bei einem Neustart automatisch" & _
      "##14  betrete."
  AddCommand "say", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** SAY <[" & ServerChannelPrefixes & "]channel>/<nick> <message>" & _
      "##14  Bringt mich dazu, im Raum <[" & ServerChannelPrefixes & "]channel> oder zum User" & _
      "##14  <nick> etwas zu sagen."
  AddCommand "action act", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** ACTION, ACT <[" & ServerChannelPrefixes & "]channel>/<nick> <message>" & _
      "##14  Bringt mich dazu, im Raum <[" & ServerChannelPrefixes & "]channel> oder im Chat mit" & _
      "##14  <nick> etwas zu tun (wirkt wie '/me <message>')."
  AddCommand "chinfo", Cl_Mas, "", "", "+m", "2*** CHINFO <user> <info line>" & _
      "##14  Ändert die Information, die ich anzeige, wenn" & _
      "##14  jemand in einem Raum !whois <user> eingibt."
  AddCommand "ignores", Cl_Mas, "", "", "+m", "2*** IGNORES" & _
      "##14  Zeigt die Liste der ignorierten Hostmasks an." & _
      "##14  Private Messages von diesen Hostmasks werden immer" & _
      "##14  ignoriert; ebenso auch CTCPs und DCC Chats."
  AddCommand "+ignore", Cl_Mas, "", "", "+m", "2*** +IGNORE <hostmask>" & _
      "##14  Fügt eine Hostmask zu meiner Ignore-Liste hinzu." & _
      "##14  Private Messages von dieser Hostmask werden dann ab" & _
      "##14  sofort ignoriert; ebenso auch CTCPs und DCC Chats."
  AddCommand "-ignore", Cl_Mas, "", "", "+m", "2*** -IGNORE <hostmask or number>" & _
      "##14  Entfernt eine Hostmask von der Ignore-Liste. Sie" & _
      "##14  können wahlweise die Hostmask selbst oder die in" & _
      "##14  .IGNORES aufgeführte Nummer des Ignores angeben."
  AddCommand "splits", Cl_Mas, "", "", "+m", "2*** SPLITS" & _
      "##14  Zeigt die Server an, die ihre Verbindung zum IRC" & _
      "##14  verloren haben und auf einen Reconnect warten." & _
      "##14  Hinter den Servernamen ist die Zeit angegeben," & _
      "##14  seit dem der Server gesplittet ist."
  AddCommand "adduser", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** ADDUSER <nick> (handle)" & _
      "##14  Fügt einen User zu meiner Userliste hinzu. Der mit" & _
      "##14  <nick> angegebene User muß sich dazu in einem" & _
      "##14  meiner Chaträume aufhalten, damit ich seine sog." & _
      "##14  'Hostmask' herausfinden kann. Optional kann auch" & _
      "##14  <handle> angegeben werden. Der User wird dann" & _
      "##14  unter dem dort angegebenen Nick hinzugefügt." & _
      "##14  -> siehe auch '+user','remuser'"
  AddCommand "remuser", Cl_Mas Or Cl_CMas, "", "", "", "2*** REMUSER <user>" & _
      "##14  Entfernt den User <user> aus meiner Userliste. Er" & _
      "##14  verliert dabei all seine Einstellungen und Flags." & _
      "##14  (gleiche Funktion wie -USER)"
  AddCommand "+user", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** +USER <nick> (hostmask)" & _
      "##14  Fügt einen User zu meiner Userliste hinzu. Der" & _
      "##14  Unterschied zu ADDUSER ist, daß der User sich nicht" & _
      "##14  in einem Raum aufhalten muß, in dem ich auch gerade" & _
      "##14  bin. Der neu hinzugefügte User bekommt die globalen" & _
      "##14  Flags '" & BaseFlags & "' und wahlweise eine Hostmask oder keine." & _
      "##14  -> siehe auch '-user','adduser'"
  AddCommand "-user", Cl_Mas Or Cl_CMas, "", "", "", "2*** -USER <user>" & _
      "##14  Entfernt den User <user> aus meiner Userliste. Er" & _
      "##14  verliert dabei all seine Einstellungen und Flags." & _
      "##14  (gleiche Funktion wie REMUSER)" & _
      "##14  -> siehe auch '+user','remuser'"
  AddCommand "+host", Cl_Mas, "", "", "", "2*** +HOST <user> <hostmask>" & _
      "##14  Fügt zum Eintrag des Users <user> in meiner Liste" & _
      "##14  eine weitere Hostmask hinzu. Es ist immer wichtig," & _
      "##14  daß ich die aktuelle Hostmask meiner User besitze," & _
      "##14  damit ich sie erkennen und beschützen kann.##" & EmptyLine + _
      "##14  Wenn mehrere Hostmasks auf einen User passen, dann" & _
      "##14  wird automatisch die genaueste genommen. Beispiel:" & _
      "##14  - UserA hat die Hostmask *!*@*.surf1.de" & _
      "##14  - UserB hat die Hostmask *Hip*!*@Socks*.surf1.de" & _
      "##14  - Wenn jetzt 'Hippo!usr123@SocksProxy.surf1.de' in" & _
      "##14    einen Channel kommt, wird er als UserB erkannt," & _
      "##14    da dessen Hostmask genauer angegeben ist."
  AddCommand "-host", Cl_Mas, "+m", "", "", "2*** -HOST (nick) <hostmask>" & _
      "##14  Entfernt eine Hostmask des Users <user> aus" & _
      "##14  meiner Liste." & _
      "##14  -> siehe auch '+host'"
  AddCommand "bans", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** BANS" & _
      "##14  Zeigt eine Liste der permanenten Bans im Bot an." & _
      "##14  -> siehe auch '+ban','-ban'"
  AddCommand "+ban +sban", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** +(S)BAN <hostmask> ([" & ServerChannelPrefixes & "]channel) (comment)" & _
      "##14  Fügt die Hostmask <hostmask> zu meiner permanenten" & _
      "##14  Banliste hinzu. Wenn <[" & ServerChannelPrefixes & "]channel> angegeben wird, gilt" & _
      "##14  der Ban nur im angegebenen Channel. Das <comment>" & _
      "##14  wird später als Kick-Grund genommen. Bei '.+sban'" & _
      "##14  banne ich die angegebene Hostmask nicht sofort nach" & _
      "##14  Eingabe des Befehls in den Channels."
  AddCommand "-ban -sban", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** -(S)BAN <hostmask or number>" & _
      "##14  Entfernt den Ban mit der Hostmask <hostmask> oder der" & _
      "##14  angegebenen Nummer von meiner Banliste. Wenn der Ban" & _
      "##14  in irgendeinem Channel gesetzt ist, entferne ich ihn." & _
      "##14  Bei der Benutzung von '.-sban' geschieht dies nicht."
  AddCommand "invites", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** INVITES" & _
      "##14  Zeigt eine Liste der permanenten Invites im Bot an." & _
      "##14  -> siehe auch '+invite','-invite'"
  AddCommand "+invite", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** +INVITE <hostmask> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Fügt die Hostmask <hostmask> zu meiner permanenten" & _
      "##14  Inviteliste hinzu. Wenn <[" & ServerChannelPrefixes & "]channel> angegeben wird, gilt" & _
      "##14  der Invite nur im angegebenen Channel."
  AddCommand "-invite", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** -INVITE <hostmask or number>" & _
      "##14  Entfernt den Invite mit der Hostmask <hostmask> oder der" & _
      "##14  angegebenen Nummer von meiner Inviteliste. Wenn der Invite" & _
      "##14  in irgendeinem Channel gesetzt ist, entferne ich ihn."
  AddCommand "excepts", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** EXCEPTS" & _
      "##14  Zeigt eine Liste der permanenten Excepts im Bot an." & _
      "##14  -> siehe auch '+except','-except'"
  AddCommand "+except", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** +EXCEPT <hostmask> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Fügt die Hostmask <hostmask> zu meiner permanenten" & _
      "##14  Exceptliste hinzu. Wenn <[" & ServerChannelPrefixes & "]channel> angegeben wird, gilt" & _
      "##14  der Except nur im angegebenen Channel. Das <comment>" & _
      "##14  wird später als Kick-Grund genommen."
  AddCommand "-except", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** -EXCEPT <hostmask or number>" & _
      "##14  Entfernt den Except mit der Hostmask <hostmask> oder der" & _
      "##14  angegebenen Nummer von meiner Exceptliste. Wenn der Except" & _
      "##14  in irgendeinem Channel gesetzt ist, entferne ich ihn."
  AddCommand "stick", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** STICK <hostmask or number>" & _
      "##14  Macht einen Ban 'sticky' - er kann im Channel nicht" & _
      "##14  mehr entfernt werden, da ich ihn immer sofort wieder" & _
      "##14  setze." & _
      "##14  -> siehe auch 'unstick'"
  AddCommand "unstick", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** UNSTICK <hostmask or number>" & _
      "##14  Entfernt die 'sticky'-Eigenschaft eines Bans. Er kann" & _
      "##14  dadurch wieder im Channel entfernt werden." & _
      "##14  -> siehe auch 'stick'"
  AddCommand "chnick", Cl_Mas Or Cl_CMas, "", "", "+m", "2*** CHNICK <old nick> <new nick>" & _
      "##14  Ändert den Namen eines Users in der Userliste."
  'For owners ----------
  AddCommand "join", Cl_Own Or Cl_COwn, "", "", "", "2*** JOIN <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Bringt mich dazu, den Raum <[" & ServerChannelPrefixes & "]channel> zu betreten."
  AddCommand "part", Cl_Own Or Cl_COwn, "", "", "", "2*** PART <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Bringt mich dazu, den Raum <[" & ServerChannelPrefixes & "]channel> zu verlassen."
  AddCommand "+chan", Cl_Own, "", "", "+n", "2*** +CHAN <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Fügt einen permanenten Channel hinzu. Dieser" & _
      "##14  Channel wird nach einem Neustart automatisch" & _
      "##14  von mir betreten. Falls ich noch nicht in dem" & _
      "##14  angegebenen Channel bin, gehe ich hinein."
  AddCommand "-chan", Cl_Own, "", "", "+n", "2*** -CHAN <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Entfernt einen permanenten Channel. Ich werde" & _
      "##14  nach einem Neustart nicht in diesen Channel" & _
      "##14  zurückkehren. Falls ich mich gerade in dem" & _
      "##14  angegebenen Channel befinde, verlasse ich ihn."
  AddCommand "cycle rejoin", Cl_Own, "", "", "+n", "2*** CYCLE <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Verlässt vorübergehend einen Channel. Ich werde" & _
      "##14  nach einer kuzern Pause in diesen Channel" & _
      "##14  zurückkehren."
  AddCommand "boot", Cl_Own, "", "", "", "2*** BOOT <user> (reason)" & _
      "##14  Kickt den User <user> von der Partyline. Du kannst" & _
      "##14  auch einen Grund bei <reason> angeben."
  AddCommand "chansetup chanset cs", Cl_Own Or Cl_COwn, "", "", "+n", "2*** CHANSETUP, CHANSET, CS <[" & ServerChannelPrefixes & "]channel>" & _
      "##14  Ruft ein Menü auf, in dem man die Einstellungen" & _
      "##14  für den Raum <[" & ServerChannelPrefixes & "]channel> ändern kann."
  AddCommand "policysetup polsetup ps", Cl_SOwn Or Cl_SOwn, "", "", "+s", "2*** POLICYSETUP, POLSETUP, PS" & _
      "##14  Ruft ein Menü auf, in dem man die Sicherheitseinstellungen" & _
      "##14  für den Bot ändern kann."
  AddCommand "chattr", Cl_Own Or Cl_COwn, "", "", "+n", "2*** CHATTR <user> <+/-flags> ([" & ServerChannelPrefixes & "]channel)" & _
      "##14  Ändert die Flags des Users <user> bei mir. Du kannst" & _
      "##14  mit dem Befehl '.whois <user>' die aktuellen Flags" & _
      "##14  erfahren. Um z.B. dem User 'Blubb' das Flag 'p' zu" & _
      "##14  geben, mußt Du '.chattr Blubb +p' eingeben. Wenn" & _
      "##14  Du keinen <[" & ServerChannelPrefixes & "]channel> angibst, dann ist das Flag" & _
      "##14  global (gilt für alle Räume), ansonsten gilt es nur für" & _
      "##14  den angegebenen Raum." & _
      "##14  Eine Liste aller Flags und ihre jeweilige Bedeutung" & _
      "##14  kannst Du mit '.help whois' bekommen."
  AddCommand "servers", Cl_Own, "", "", "+n", "2*** SERVERS" & _
      "##14  Dieser Befehl zeigt die Liste der Server an, zu" & _
      "##14  denen ich zu connecten versuche, falls ich einmal" & _
      "##14  von meinem IRC-Server disconnected werde." & _
      "##14  -> siehe auch '+server','-server','jump'"
  AddCommand "+server", Cl_Own, "", "", "+n", "2*** +SERVER <address:port(:password)> (|proxy(:port))" & _
      "##14  Fügt einen neuen Server zu meiner Server-Liste hinzu." & _
      "##14  Der Standard-Port ist '6667', jedoch sollte nach" & _
      "##14  Möglichkeit ein anderer Port verwendet werden, um" & _
      "##14  'ICMP Unreach'-Attacken auszuweichen, die oft nur auf" & _
      "##14  diesen Port gehen. Optional kann ein Server-Passwort" & _
      "##14  angegeben werden, falls es vom IRC-Server verlangt" & _
      "##14  wird (gibts eher selten)." & _
      "##14  -> siehe auch '-server','servers','jump'"
  AddCommand "-server", Cl_Own, "", "", "+n", "2*** -SERVER <address:port or number> (|proxy(:port))" & _
      "##14  Entfernt einen Server von meiner Server-Liste." & _
      "##14  Der letzte Server meiner Liste kann nicht entfernt" & _
      "##14  werden, da ich im Notfall immer einen Server wissen" & _
      "##14  muß." & _
      "##14  -> siehe auch '+server','servers','jump'"
  AddCommand "reconnect", Cl_Own, "", "", "+n", "2*** RECONNECT" & _
      "##14  Verbindet mich erneut zum Server."
  AddCommand "jump", Cl_Own, "", "", "+n", "2*** JUMP <server number> or <new server(:port)> (|proxy(:port))" & _
      "##14  Läßt mich zu einem anderen Server connecten. Es kann" & _
      "##14  wahlweise eine Nummer aus der '.servers'-Liste oder" & _
      "##14  eine <IRCServer:Port> Kombination angegeben werden." & _
      "##14  Falls der connect fehlschlägt, wird der nächste Server" & _
      "##14  in meiner Server-Liste versucht.  " & _
      "##14  " & _
      "##14  Falls über einen SOCKS-Proxy connected werden soll," & _
      "##14  muß dieser mit einem führenden '|' angegeben werden." & _
      "##" & EmptyLine + _
      "##14  Beispiel: .jump irc.test.de:6667 |my.proxy.de:1080" & _
      "##14  -> siehe auch 'servers','+server','-server'"
  'For super owners ----------
  AddCommand "kisetup ks", Cl_SOwn, "", "", "+s", "2*** KISETUP" & _
      "##14  Zeigt ein Menü an, in dem man alle 'Personen'-Einstellungen" & _
      "##14  des Bots komfortabel verändern kann. Dazu zählen:" & _
      "##14  * Vor und Nachname des Bots" & _
      "##14  * Alter ... etc ..."
  AddCommand "dccstat ds", Cl_SOwn, "", "", "+m", "2*** DCCSTAT" & _
      "##14  Zeigt ein eine Tabelle mit 'dcc' verbindungen zum/vom Bot."
  AddCommand "netsetup ns", Cl_SOwn, "", "", "+s", "2*** NETSETUP" & _
      "##14  Zeigt ein Menü an, in dem man alle Netzwerk-Einstellungen" & _
      "##14  des Bots komfortabel verändern kann. Dazu zählen:" & _
      "##14  * Nicklänge, Max. Channel anzahl" & _
      "##14  * Netzmame ... etc ..."
  AddCommand "auth", Cl_SOwn, "", "", "+s", "2*** AUTH" & _
      "##14  Veranlasst den Bot sich neu zu AUTHen."
  AddCommand "authsetup", Cl_SOwn, "", "", "+s", "2*** AUTHSETUP" & _
      "##14  Zeigt ein Menü an, in dem man Einstellungen zum" & _
      "##14  AUTHen machen kann."
  AddCommand "botsetup bs", Cl_SOwn, "", "", "+s", "2*** BOTSETUP, BS" & _
      "##14  Zeigt ein Menü an, in dem man alle Einstellungen" & _
      "##14  des Bots komfortabel verändern kann. Dazu zählen:" & _
      "##14  * Erster und zweiter Bot-Nick" & _
      "##14  * Username (Ident) und Botnet-Nick" & _
      "##14  * Realname des Bots, Quit-Message" & _
      "##14  * Telnet-Ports für User und Bots" & _
      "##14  * 'IDENT'-Kommando" & _
      "##14  * ...und vieles mehr!"
  AddCommand "update", Cl_SOwn, "", "", "+s", "2*** UPDATE" & _
      "##14  Startet die AutoUpdate-Sequenz, die eine neue Version" & _
      "##14  des Bots installiert. Vorher muß dem Bot per DCC" & _
      "##14  die neueste AnGeL.exe zugeschickt werden."
  AddCommand "restart", Cl_SOwn, "", "", "+s", "2*** RESTART" & _
      "##14  Startet den Bot neu, indem sich die AnGeL.exe" & _
      "##14  beendet und sich wieder neu aufruft."
  AddCommand "die", Cl_SOwn, "", "", "+s", "2*** DIE" & _
      "##14  Lässt den Bot 'sterben', indem er sich beendet." & _
      "##14  Zur Sicherheit muß '.die sure' eingegeben werden."
  AddCommand "botnick", Cl_SOwn, "", "", "+s", "2*** BOTNICK <new nick>" & _
      "##14  Ändert meinen IRC-Nick auf <new nick>, falls der" & _
      "##14  Nick noch nicht vergeben ist. Diese Nick-Änderung" & _
      "##14  ist aber nicht permanent - wenn ich neugestartet" & _
      "##14  werde, benutze ich wieder meinen alten Nick."
  AddCommand "check", Cl_SOwn, "", "", "+s", "2*** CHECK <1 / 2 / 3>" & _
      "##14  1: Geht meine Userliste nach Usern durch, die" & _
      "##14     kein Passwort gesetzt haben, und zeigt diese" & _
      "##14     an. User ohne Passwort sind ein Risiko." & _
      "##14  2: Sucht die Userliste nach Bots ohne Passwort" & _
      "##14     durch. Ebenfalls ein Sicherheitsrisiko!" & _
      "##14  3: Checkt alle User durch, ob sie für Channels" & _
      "##14     Flags besitzen, in denen ich gar nicht mehr" & _
      "##14     bin, und entfernt die Flags gegebenenfalls." & _
      "##14  4: Zeigt alle User an, die keine Hostmask besitzen.##" & _
      "##14  5: Zeigt alle Bots an, die keine Hostmask besitzen.##" & _
      "##14  6: Zeigt alle User/Bots an, die AutoOp eingeschaltet##" & _
      "##14     haben. AutoOp ist ein Sicherheitsrisiko!"
  AddCommand "killnotes", Cl_SOwn, "", "", "+s", "2*** KILLNOTES" & _
      "##14  Löscht ALLE im Bot gespeicherten Notes."
  AddCommand "floodprot", Cl_SOwn, "", "", "+s", "2*** FLOODPROT (new max bytes/sec)" & _
      "##14  Verändert die Stärke meines Flood-Schutzes. Je nach" & _
      "##14  Server sind unterschiedliche Werte ideal. Der bei" & _
      "##14  mir voreingestellte Wert liegt bei 150 Bytes/Sek." & _
      "##14  Je kleiner die Zahl der Bytes ist, desto langsamer" & _
      "##14  arbeite ich mehrere Aufgaben ab. Es ist bei einer" & _
      "##14  vernünftigen Einstellung dieses Wertes unmöglich," & _
      "##14  mich aus dem IRC zu flooden."
  AddCommand "scripts", Cl_SOwn, "", "", "+s", "2*** SCRIPTS" & _
      "##14  Zeigt eine Liste aller geladenen AnGeL-Scripts an." & _
      "##14  Falls das Laden eines Scriptes beim Start des Bots" & _
      "##14  fehlgeschlagen ist, wird dies ebenfalls in der" & _
      "##14  Liste aufgeführt." & _
      "##14  AnGeL-Scripts haben die Endung .asc und können dem" & _
      "##14  Bot via DCC send durch einen Superowner geschickt" & _
      "##14  werden. Informationen zu Scripten und Beispiele" & _
      "##14  können beim IRCnet Chatter 'Hippo' erfragt werden." & _
      "##14  Er ist der Programmierer dieses Bots und gibt auch" & _
      "##14  gerne Hilfestellung bei Problemen, bis die AnGeL-" & _
      "##14  Homepage endlich fertig ist ;o)" & _
      "##14  -> siehe auch '+script','-script'"
  AddCommand "+script", Cl_SOwn, "", "", "+s", "2*** +SCRIPT <filename>" & _
      "##14  Lädt ein AnGeL-Script in den Bot und startet es." & _
      "##14  Falls Fehler oder Probleme dabei auftreten, wird" & _
      "##14  eine entsprechende Meldung ausgegeben. Scripts," & _
      "##14  die geladen werden können, haben die Endung .asc" & _
      "##14  und müssen sich im Verzeichnis \Scripts der" & _
      "##14  FileArea befinden. Neue Scripts können via DCC" & _
      "##14  send zum Bot geschickt werden." & _
      "##14  -> siehe auch '-script','scripts'"
  AddCommand "-script", Cl_SOwn, "", "", "+s", "2*** -SCRIPT <filename>" & _
      "##14  Entfernt ein laufendes AnGeL-Script aus dem Bot." & _
      "##14  Falls das Script aufgrund von Problemen nicht" & _
      "##14  gestartet werden konnte, wird nur der interne" & _
      "##14  Verweis auf das Script entfernt." & _
      "##14  -> siehe auch '+script','scripts'"
  AddCommand "reload rescript", Cl_SOwn, "", "", "+s", "2*** RELOAD <filename>" & _
      "##14  Entfernt ein laufendes AnGeL-Script aus dem Bot" & _
      "##14  und lädt es danach sofort wieder."
  AddCommand "talk", Cl_SOwn, "", "", "+s", "2*** TALK <message>" & _
      "##14  Bringt mich dazu, etwas im Botnet zu sagen."
  AddCommand "bytes", Cl_SOwn, "", "", "+s", "2*** BYTES" & _
      "##14  Lässt mich sagen, wieviel Traffic ich 'verbraten' hab."
  AddCommand "sharecmd", Cl_SOwn, "", "", "+s", "2*** SHARECMD" & _
      "##14  Sendet einen Befehl an alle lokalen sharing Partner."
  AddCommand "traffic", Cl_SOwn, "", "", "+s", "2*** TRAFFIC" & _
      "##14  Zeigt eine grafische Darstellung des Datenverkehrs."
  If AllowRunToS = True Then
  AddCommand "run", Cl_SOwn, "", "", "+s", "2*** RUN <APPLICATION>" & _
      "##14  Lässt mich ein Programm starten."
  End If
End Sub

Public Sub ShowHelp(vsock As Long, Line As String) ' : AddStack "Help_ShowHelp(" & vsock & ", " & Line & ")"
  Dim i As Long, i2 As Long, curp As Long, Char As String, Tex As String, CommandName As String
  Dim CurrentSection As Integer, DidIntro As Boolean, ComLine As String, ComCount As Byte
  Dim Result As Long, ShownCommands As String, UsrLevel As Integer
  CommandName = LCase(Param(Line, 2))
  Select Case CommandName
  Case ""
    UsrLevel = CLevel(SocketItem(vsock).Flags, AllFlags(SocketItem(vsock).RegNick))
    TU vsock, MakeMsg(MSG_PLHelpIntro, BotNetNick & ", AnGeL " & BotVersionEx + IIf(ServerNetwork <> "", " <" & ServerNetwork & ">", ""))
    For i2 = 0 To 8
      CurrentSection = 2 ^ i2
      DidIntro = False
      For i = 1 To CommandCount
        'Don't show hidden commands
        If Commands(i).Description <> "" Then
          If ((UsrLevel And Commands(i).Class) And CurrentSection) > 0 Then
            If (Commands(i).MatchFlags = "") Or ((Commands(i).MatchFlags <> "") And (MatchUserFlags(SocketItem(vsock).UserNum, Commands(i).MatchFlags) = True)) Then
              If DidIntro = False Then
                Select Case CurrentSection
                  Case Cl_What: TU vsock, MakeMsg(MSG_PLHelpFlagW)
                  Case Cl_Op:   TU vsock, MakeMsg(MSG_PLHelpFlagO)
                  Case Cl_Net:  TU vsock, MakeMsg(MSG_PLHelpFlagT)
                  Case Cl_CMas: TU vsock, MakeMsg(MSG_PLHelpFlagCM)
                  Case Cl_Mas:  TU vsock, MakeMsg(MSG_PLHelpFlagM)
                  Case Cl_COwn: TU vsock, MakeMsg(MSG_PLHelpFlagCN)
                  Case Cl_Own:  TU vsock, MakeMsg(MSG_PLHelpFlagN)
                  Case Cl_SOwn: TU vsock, MakeMsg(MSG_PLHelpFlagS)
                End Select
                DidIntro = True
              End If
              ComCount = ComCount + 1
              If Commands(i).Script = "" Then
                ComLine = ComLine & " " & Param(Commands(i).Name, 1) + IIf(ComCount < 5, Spaces(12, Param(Commands(i).Name, 1)), "")
              Else
                ComLine = ComLine & "3 " & Param(Commands(i).Name, 1) & "2" & IIf(ComCount < 5, Spaces(12, Param(Commands(i).Name, 1)), "")
              End If
              If ComCount = 5 Then TU vsock, "2 " & ComLine: ComLine = "": ComCount = 0
            End If
          End If
        End If
      Next i
      If ComCount > 0 Then TU vsock, "2 " & ComLine: ComLine = "": ComCount = 0
    Next i2
    TUEx vsock, SF_ExtraHelp, MakeMsg(MSG_PLHelpOutro)
    TU vsock, EmptyLine
  Case Else
    Result = CheckCommand(vsock, CommandName)
    If Result = CC_NoCommand Then
      TU vsock, "5*** Sorry, no help available on that!"
    ElseIf Result = CC_NotAllowed Then
      TU vsock, "5*** Sorry, you're not allowed to use this command!"
    Else
      curp = 1
      Do
        i2 = InStr(curp, Commands(Result).Description, "##")
        If i2 > 0 Then
          TU vsock, Mid(Commands(Result).Description, curp, i2 - curp)
          curp = i2 + 2
        Else
          TU vsock, Mid(Commands(Result).Description, curp)
          Exit Do
        End If
      Loop
    End If
  End Select
End Sub


