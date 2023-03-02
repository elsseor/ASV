############################################# Einstellungen
# Setzen Sie $MailAnspPa auf $true, wenn auch die Emailadressen der
# Ansprechpartner in die Exceldatei übernommen werden sollen
$MailAnspPa = $false
# $MailAnspPa = $true

$filename = Betriebe.xlsx

############################################# Programm


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

$file = Import-Excel -path $skriptpath\Betriebe.xlsx
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
