############################################# Fehltageimport
Author: Michael Rößle
Version: 1.0


############################################# Einstellungen



############################################# Programm
clear-host
$skriptpath = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
Set-Location $skriptpath


$exportFile = "FehltageToASV"
$files = Get-ChildItem .\*.csv
$inhalt = @()
$einträge = @()
[regex]$pattern = "Text"
$exportArray = @('"Externe Id","Datum","unent."')




### Webuntis-CSV-Datei bearbeiten, damit die Spalte "Text" nicht doppelt vorkommt
foreach ($f in $files){
    if (!(Get-Content -Path $f)[0].Contains("Betrieb")) {
        $R=[Regex]'Text'
        (Get-Content -Path $f) | ForEach-Object {$R.Replace($_, "Betrieb", 1)}  | Set-Content -Path $f
        }
}



### Inhalt der CSV-Dateien einlesen
foreach ($f in $files){
    $inhalt += import-csv $f -Delimiter `t
}


### Zeilarray mit Werten fülle und in Datei exportieren
foreach ($i in $inhalt){
    if ($i.'Externe Id' -ne $null){
        if ($i.Status -eq "nicht entsch."){
            $i.Status = "1"}
        else{$i.Status = "0"
        }
        $exportArray += '"' +$i.'Externe Id' + '","' + $i.Datum + '","' + $i.Status + '"'
        }
}

$exportArray | set-content -path $exportFile".csv"
