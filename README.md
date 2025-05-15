
# AnGeL-Bot ![angel-logo](https://github.com/user-attachments/assets/417f7c00-46d5-44ba-9316-c623b3ef43b1)


AnGeL-Bot war ein in Visual Basic entwickelter IRC-Bot für Windows, der sich durch seine Unterstützung für **AnGeL Script** (eine interne Bezeichnung für VBScript) auszeichnete. Der Bot wurde ursprünglich 1997 entwickelt und war ab Ende 1998 öffentlich verfügbar.

![angel-bot](https://github.com/user-attachments/assets/f4183243-a5f4-46fd-881b-558952548c27)


## Merkmale

- Begrüßung von Benutzern
- Protokollierung von Ein- und Austritten
- Benutzerrechteverwaltung
- Kick-/Ban-Funktionalität
- Reaktionen auf Befehle, Schlüsselwörter
- Spiele, Zufallsantworten, einfache KI
- Einbindung externer Dateien
- Integration von Skripten mit `.asc`-Endung (VBScript)

## AnGeL Script

AnGeL Script basierte auf VBScript und erlaubte benutzerdefinierte Automatisierung:

```vbscript
Sub OnJoin(User)
  If LCase(User) = "angel" Then
    SendChannelMessage "Willkommen zurück, Meister!"
  End If
End Sub
```

Events wie `OnJoin`, `OnPart`, `OnMessage` usw. konnten frei geskriptet werden.

Spätere Versionen unterstützten auch Perl, TCL, PHP, Ruby, JavaScript über den Windows Script Host (WSH).

## Addins

Erweiterungen waren über signierte COM-Komponenten (Addins) möglich. Diese konnten zusätzliche Funktionalitäten bereitstellen – wurden jedoch aus Sicherheitsgründen nur bei gültiger Signatur geladen.

## Verbreitung & Community

Das Skript-Portal `angelbot-portal.de`, das **nicht von den Entwicklern selbst betrieben wurde**, zählte im Jahr 2006 über **12.000 registrierte Nutzer**. Laut archivierten Daten stieg diese Zahl bis 2013 auf über **50.000 Mitglieder**. Die Plattform bot Foren, Hilfen und hunderte Skripte rund um AnGeL-Bot.

## Versionen

| Version         | Jahr     | Bemerkung                                          |
|-----------------|----------|----------------------------------------------------|
| 1.0.0           | ca. 1998 | Erste öffentliche Version                          |
| 1.5.5           | 1999     | Skriptunterstützung                                |
| 1.5.9           | 1999     | Mehrsprachigkeit                                   |
| 1.6.0           | 2000     | Verbesserungen bei TelNet und Scripting            |
| 1.6.2 Beta 10   | 2003     | Letzte bekannte Beta-Version                       |

## Historische Bedeutung

AnGeL-Bot wurde in zahlreichen deutschsprachigen IRC-Communities verwendet und stellte eine seltene Windows-native Alternative zu Linux-orientierten IRC-Bots wie Eggdrop dar. Die Möglichkeit, COM-Addins und VBScript in Echtzeit einzubinden, war für die Zeit ungewöhnlich fortschrittlich.

## Screenshots

![Konfiguration](https://github.com/user-attachments/assets/142dc48d-152d-4a4d-ab62-7fe3aa20f38c)

![PartyLine](https://github.com/user-attachments/assets/235134c2-35bd-4c5e-8d02-262256174cdf)

## Archivquellen

- https://web.archive.org/web/20000229085321/http://angel.chan.de/
- https://web.archive.org/web/20021005023700/http://www.ircbots.de/angels.html
- https://web.archive.org/web/20060203232521/http://mercurys.fmo-clan.de:80/angelbot/
- https://web.archive.org/web/20140328225234/http://www.angelbot-portal.de/
- https://www.vbarchiv.net/forum/id14_i4899t4899_altes-projekt-in-vb6-angel-bot-fuers-irc-mit-problemchens.html
- http://wp.xin.at/archives/4343

## Lizenz

GPLv3 – siehe [LICENSE.md](https://github.com/bhuebschen/AnGeL-Bot/blob/main/LICENSE)

## Wikipedia

Durch offenbar fehlende Erwähnung in Fachzeitschriften und herabsetzung der Seriösistät anderer Quellen, hat es der AnGeL-Bot leider nicht in die Wikipedia geschafft
https://de.wikipedia.org/wiki/Wikipedia:L%C3%B6schkandidaten/12._Mai_2025#AnGeL-Bot
