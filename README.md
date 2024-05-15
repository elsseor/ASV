# ASVImport
In diesem Verzeichnis finden Sie ein Skript mit dem aus den XML-Dateien der Schüleranmeldungen (Formularservice der Stadt Augsburg) eine Excelliste erstellt wird, die für den Import in die ASV geeignet ist.

# Mailinglisten
In diesem Verzeichnis finden Sie das Skript ***creatMailList.ps1*** mit dem aus der Betriebe-Excelliste aus der ASV eine Excelliste mit Emails und Betriebnamen erstellt wird, die z.B. für Serienbrief geeignet ist.

# Exportformatibibliothek
Hier finden Sie Exportformatdateien für die ASV. Derzeit exisitiert nur die Datei für den Export aus der ASV nach Webuntis.

# FehltageImport
Hier finden Sie ein Powershellmodul mit dem die Abwesenheiten aus Webuntis konvertiert werden können, um sie in die ASV einzuspielen.

# Installation des benötigten Moduls:
Einmalig muss das PowershellModul "Export-Excel" installiert werden. Dazu muss man die Powershell mit Administratorrechten öffenen und den Befehl "Install-Module -Name ImportExcel" ausführen. Hinweis für Augsburger Schulen: Eventuell muss hier die IT der Stadt Augsburg helfen. Für das eigentliche Programm werden dann kein Administratorrechte mehr benötigt. Man kann der Stadt-IT auch sagen, dass sie das Modul im Profil des Users hinterlegen soll, dann funktioniert es für diesen User an allen VerwaltungsPCs. Der Pfad hierfür lautet: "\\swf10101<benutzer>\Eigene Dateien\WindowsPowerShell\Modules\ImportExcel"
