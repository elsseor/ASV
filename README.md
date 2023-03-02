# ASV - Skripte für die ASV

1.	CreateMailList.ps1

Beschreibung
Mit diesem Programm wird aus der Betriebe-Excelliste aus der ASV eine Excelliste mit Emails und Betriebnamen erstellt, die z.B. für Serienbrief geeignet ist.
Achtung: Einmalig muss das PowershellModul "Export-Excel" installiert werden. Das funktioniert mit dem Skript InstallExport.ps1, das mit Administratorrechten ausgeführt werden muss. Eventuell muss hier die IT der Stadt Augsburg helfen. Für das eigentliche Programm werden dann kein Administratorrechte mehr benötigt.

Ablauf
1)	Aus der ASV eine Excelliste mit den Betrieben exportieren. Es müssen mindestens die beiden Spalten „Betriebename Zeile 1“ und „Anschrift E-Mail“ enthalten sein.
2)	Diese Datei mit Excel öffnen und unter dem Namen „Betriebe.xlsx“ in dem gleichen Verzeichnis wie das Skript abspeichern. Dies ist nötig, da die Powershell das neue Excelformat benötigt, die ASV aber nur in das alte Format exportiert.
3)	Das Skript „createMailList.ps1“ mit der Powershell ausführen.
4)	Im Skriptverzeichnis wird nun die Exceldatei „mail.xlsx“ erstellt.

2. InstallExport.ps1
Dieses Skript muss einmalig mit Administratorrechten ausgeführt werden, damit das Export-Excel-Modul für Powershell installiert wird
