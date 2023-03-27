# Beschreibung:
Mit dem Skript ***createMailList.ps1*** wird aus der Betriebe-Excelliste aus der ASV eine Excelliste mit Emails und Betriebnamen erstellt, die z.B. für Serienbrief geeignet ist. In der Datei wird jede Emailadresse in einer neuen Zeile gespeichert und in einer weiteren Spalte hierzu der Betriebename.

# Ablauf
1. Aus der ASV die Betriebeliste im Excelformat exportieren
2. Da die ASV leider noch das veraltete Excelformat verwendet, mit dem die Powershell nicht umgehen kann, muss diese Datei in Excel geöffnet werden und unter dem aktuellen XLSX-Format und dem Namen ***Betriebe.xlsx*** abgespeichert werden.
3. Die Datei in das Verzeichnis mit dem Powershellscript abspeichern.
4. ***createMailList.ps1*** ausführen. Das Ergebnis steht in der neuen Datei ***mail.xlsx***

Falls nur eine Mailingliste für Betriebe einzelner Klassen ierstellt werden, so müssen Sie zusätzlich aus der ASV eine Schülerliste für die entsprechendn Klassen im Excelformat exportieren, in das aktuelle Format konvertieren und im gleichen Verzeichnis wie das Skript abspeichern. Die Schülerliste muss jedoch die Spalte "Ausb. Betrieb Name1" aus der ASV enthalten. 

# Installation des benötigten Moduls:
Einmalig muss das PowershellModul "Export-Excel" installiert werden. Dazu muss man die Powershell mit Administratorrechten öffenen und den Befehl "Install-Module -Name ImportExcel" ausführen.
Hinweis für Augsburger Schulen: Eventuell muss hier die IT der Stadt Augsburg helfen. Für das eigentliche Programm werden dann kein Administratorrechte mehr benötigt. Man kann der Stadt-IT auch sagen, dass sie das Modul im Profil des Users hinterlegen soll, dann funktioniert es für diesen User an allen VerwaltungsPCs. Der Pfad hierfür lautet: ***"\\swf10101\<benutzer>\Eigene Dateien\WindowsPowerShell\Modules\ImportExcel"***
