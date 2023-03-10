############################################# Einstellungen
# Setzen Sie $MailAnspPa auf $true, wenn auch die Emailadressen der
# Ansprechpartner in die Exceldatei übernommen werden sollen
$MailAnspPa = $true
# $MailAnspPa = $true


############################################# Programm
# Prüfen ob Betriebe.xlsx und schueler.xlsx im richtigen Format vorliegt


clear-host
$skriptpath = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

## Überprüfen ob die Exceldatei exisitiert und im aktuellen Format vorliegt
if (-not(Test-path -path $skriptpath\Betriebe.xlsx)){
    if (-not(Test-path -path $skriptpath\Betriebe.xls)){
        write-host Fehler: Die Datei Betriebe.xlsx kann in dem aktuellen Verzeichnis nicht gefunden werden!!!
    } else {
        write-host Fehler: Sie haben anscheinend die Exceldatei nicht auf das XLSX-Format aktualisiert!!!
    }
    Pause
    exit
}


if ((Test-path -path $skriptpath\schueler.xls) -and -not(Test-path -path $skriptpath\schueler.xlsx)){
    Write-host Es ist eine "schueler.xls" vorhanden, die nicht im aktuellen Excelformat vorliegt
    Write-Host Diese Datei kann nicht berücksichtigt werden
    Pause
}


############################################# Betriebeliste einschränken
# Falls eine schueler.xlsx vorhanden ist, wird die Liste der Betriebe so eingeschränkt,
# dass nur die Betriebe dieser Schüler berücksichtigt werden



$betriebe = @()
if (Test-path -path $skriptpath\schueler.xlsx ){
    $file1 = Import-Excel -path $skriptpath\schueler.xlsx
    foreach ($i in $file1){
        if (-not($betriebe.Contains($i.'Ausb. Betrieb Name1'))){
            $betriebe += $i.'Ausb. Betrieb Name1'
        }
    }



    $file2 = Import-Excel -path $skriptpath\betriebe.xlsx
    $zeile = @()
    foreach ($i in $file2){
        if ($betriebe.Contains($i.'Betriebename Zeile 1')){
            $zeile += $i
        }
    }

    $zeile | Export-Excel -path $skriptpath\betriebetemp.xlsx

}
else{
    copy-Item $skriptpath\betriebe.xlsx $skriptpath\betriebetemp.xlsx
}










############################################# Mailadressen exportieren
$file = Import-Excel -path $skriptpath\betriebetemp.xlsx
$liste = @()
$allmail = @()



foreach ($i in $file) {
## Durchsuchen der Spalte "Anschrift E-Mail"
    if ($i.'Anschrift E-Mail' -ne $null -and $i.'Anschrift E-Mail' -ne "" ){
        if ($i.'Anschrift E-Mail'.contains(',')){
            $email = $i.'Anschrift E-Mail' -split ","
            foreach ($k in $email){
                $k = $k.trim("[").trim("]").trim(" ")
                if (-not($allmail.Contains($k))){
                    $allmail += $k
                    $i.'Anschrift E-Mail' = $k
                    $liste += $i.PsObject.Copy()
                }
            }
        }
        else {
            $k = $i.'Anschrift E-Mail' = $i.'Anschrift E-Mail'.trim("[").trim("]")
            if (-not($allmail.Contains($k))){
                $allmail += $k
                $i.'Anschrift E-Mail' = $k
                $liste += $i.PsObject.Copy()
            }

        }
    } 
    
## Durchsuchen der Spalte "Ansprechpartner"
    if ($i.'Ansprechpartner Kommunikationen Alle' -ne $null -and $i.'Ansprechpartner Kommunikationen Alle'.contains('@') -and $MailAnspPa -eq $true){
        $email = $i.'Ansprechpartner Kommunikationen Alle' -split " "
        foreach ($k in $email){
            if ($k.Contains("@")){
                $k = $k.trim("[").trim("]").trim(" ").trim(";")
                if (-not($allmail.Contains($k))){
                    $allmail += $k
                    $i.'Anschrift E-Mail' = $k
                    $liste += $i.PsObject.Copy()
                }
            }
        }
    }
}





##### Über Select-Object kann festgelegt, welche Spalten in die neue Exceldatei aufgenommen wird und in welcher Reihenfolge
$liste | Select-Object -Property 'Betriebename Zeile 1','Anschrift E-Mail' | Export-Excel -path $skriptpath\mail.xlsx
Remove-Item $skriptpath\betriebetemp.xlsx
