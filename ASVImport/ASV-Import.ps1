############################################# Einstellungen
# Dateinamen fuer die Exceldatei
$outputExcelXLSX = 'ASV-Import.xlsx'
$visibleExcel = $false
$klasse = "Neu"
$zielverzeichnis = 'O:\Temp'
$finalExcelXLSPath = 'O:\Temp\ASV-Import.xlsx'



############################################# Programm

clear-host


if(Test-Path $finalExcelXLSPath) {
    Remove-Item $finalExcelXLSPath -Force
    }

### Festlegen der Dateipfade
$scriptpath = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
Set-Location $scriptpath

$outputExcelXLSXPath = $scriptpath + '\' + $outputExcelXLSX
# $logExcelPath = $scriptpath + '\Config\log.xlsx'
$ConfigPath = $scriptpath + '\Config\config.xlsx'
$functionPath = $scriptpath + '\Config\ASVImportFunktionen.ps1'
$global:xmlPath = $scriptpath + "\XML-Import\"
$global:xmlArchivPath = $scriptpath + "\XML-Archiv\"
$arrayXML = Get-ChildItem -Path $xmlPath


################ externe Datei fuer Funktionen (Dot Sourcing)
if(!(Test-Path $functionPath)) {
    Write-host Datei mit Funktionen ASVImportFunktionen.ps1 fehlt
    Exit
    }
. $functionPath


############################################# Programm


ConfigFill
CreateExcelFile
AddContentExcelFile


Write-Host
Write-Host ------------- Bearbeitung ist beendet! -------------------