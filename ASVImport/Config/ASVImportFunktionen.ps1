function ConfigFill {
    write-host Ueberpruefung der Konfigurationsdatei

    $konfessionsZuordnung = Import-Excel -Path $ConfigPath -WorksheetName Konfession
    $berufsZuordnung = Import-Excel -Path $ConfigPath -WorksheetName BerufsID
    $staatZuordnung = Import-Excel -Path $ConfigPath -WorksheetName Staatsangehoerigkeit
    
    $berufMissing = @()
    $konfessionsMissing = @()
    $staatZuordnungMissing = @()

    $aenderung = $false

    $i = 2
    foreach ($xml in $arrayXML){
    
        [xml] $xmlContent = Get-Content -path $xmlPath$xml -Encoding UTF8
        # write-host  Bearbeite Schueler $xmlContent.myForm.azubi_familienname

        ### fehlende Konfession in Config-File speichern
        if ($konfessionsZuordnung.'XML-Konfession' -notcontains $xmlContent.myForm.azubi_religion){
            if ($konfessionsMissing -notcontains $xmlContent.myForm.azubi_religion){
                $konfessionsMissing += $xmlContent.myForm.azubi_religion
                $aenderung = $true
            }
        }



        ### fehlende Staatsangehoerigkeit in Config-File speichern
        if ($staatZuordnung.'StaatsangehoerigkeitFormular' -notcontains $xmlContent.myForm.azubi_staatsangehoerigkeit){
            if ($staatZuordnungMissing -notcontains $xmlContent.myForm.azubi_staatsangehoerigkeit){
                $staatZuordnungMissing += $xmlContent.myForm.azubi_staatsangehoerigkeit
                $aenderung = $true
            }
        }



        ### fehlenden Ausbildungsberuf in Config-File speichern
        if ($berufsZuordnung.BerufsbezeichnungFormular -notcontains $xmlContent.myForm.beruf_name){
            if ($berufMissing -notcontains $xmlContent.myForm.beruf_name){
                $berufMissing += $xmlContent.myForm.beruf_name
                $aenderung = $true
            }
        }


        $i++
    }


    $berufMissing | Export-Excel -path $ConfigPath -WorksheetName BerufsID -Append                               # Ergaenzt die config-Datei um fehlende Berufe
    $konfessionsMissing | Export-Excel -path $ConfigPath -WorksheetName Konfession -Append                       # Ergaenzt die config-Datei um fehlende Konfessionszuordnungen 
    $staatZuordnungMissing | Export-Excel -path $ConfigPath -WorksheetName Staatsangehoerigkeit -Append          # Ergaenzt die config-Datei um fehlende Staatsangehoerigkeitszuordnungen 

    write-host Ueberpruefung beendet

    if ($aenderung -eq $true){
        Write-Host
        write-host ---------------------------------------
        Write-Host
        write-host In der Konfiguration wurden Einstellungen ergaenzt.
        write-host Vervollstaendigen Sie diese Aenderungen und fuehren Sie das Programm erneut aus!
        $defaultPromptValue = 'n'
        $promptOutput = Read-Host "Wollen Sie das Skrt dennoch weiter ausfuehren ohne die Aenderungen zu uebernehmen? (Standard:n) (j/n) [$($defaultPromptValue)]"
     
        #if (!$promptOutput -eq "") {$defaultPromptValue = $promptOutput}
        if ($promptOutput -ne "j") {Exit} 
    
    }


}

function CreateExcelFile {

    # Exceldatei anlegen
    $MyExcel = New-Object -ComObject excel.application
    $MyExcel.visible = $false
    $Myworkbook = $MyExcel.workbooks.add()
    $Sheet1 = $Myworkbook.worksheets.item(1)
    $Sheet1.name = "ASV"

    ### Alle Felder anlegen entsprechend ASV Importdatei
    $Sheet1.cells.item(1,1) = 'Klasse'
    $Sheet1.cells.item(1,2) = 'Name'
    $Sheet1.cells.item(1,3) = 'Vornamen'
    $Sheet1.cells.item(1,4) = 'Rufname'
    $Sheet1.cells.item(1,5) = 'Geburtstag'
    $Sheet1.cells.item(1,6) = 'Geburtsdatum Gültigkeit'
    $Sheet1.cells.item(1,7) = 'Geburtsort'
    $Sheet1.cells.item(1,8) = 'Geburtsland'
    $Sheet1.cells.item(1,9) = 'Geschlecht'
    $Sheet1.cells.item(1,10) = 'Konfession'
    $Sheet1.cells.item(1,11) = 'RU'
    $Sheet1.cells.item(1,12) = 'Land'
    $Sheet1.cells.item(1,13) = 'Land2'
    $Sheet1.cells.item(1,14) = 'Strasse'
    $Sheet1.cells.item(1,15) = 'HausNr'
    $Sheet1.cells.item(1,16) = 'PLZ'
    $Sheet1.cells.item(1,17) = 'Ort'
    $Sheet1.cells.item(1,18) = 'Teilort'
    $Sheet1.cells.item(1,19) = 'Staat'
    $Sheet1.cells.item(1,20) = 'Telefon'
    $Sheet1.cells.item(1,21) = 'Handy'
    $Sheet1.cells.item(1,22) = 'Email'
    $Sheet1.cells.item(1,23) = 'Muttersprache'
    $Sheet1.cells.item(1,24) = 'Schuleintrittam'
    $Sheet1.cells.item(1,25) = 'Einschulungam'
    $Sheet1.cells.item(1,26) = 'im Schriftverkehrverteiler'
    $Sheet1.cells.item(1,27) = 'auskunftsberechtigt'
    $Sheet1.cells.item(1,28) = 'Erz1Art'
    $Sheet1.cells.item(1,29) = 'Erz1Anrede'
    $Sheet1.cells.item(1,30) = 'Erz1Name'
    $Sheet1.cells.item(1,31) = 'Erz1Vorname'
    $Sheet1.cells.item(1,32) = 'Erz1Strasse'
    $Sheet1.cells.item(1,33) = 'Erz1Hausnr'
    $Sheet1.cells.item(1,34) = 'Erz1PLZ'
    $Sheet1.cells.item(1,35) = 'Erz1Ort'
    $Sheet1.cells.item(1,36) = 'Erz1Teilort'
    $Sheet1.cells.item(1,37) = 'Erz1Telefon'
    $Sheet1.cells.item(1,38) = 'Erz1Handy'
    $Sheet1.cells.item(1,39) = 'Erz1Email'
    $Sheet1.cells.item(1,40) = 'Erz1Schriftverkehrverteiler'
    $Sheet1.cells.item(1,41) = 'Erz1auskunftsberechtigt'
    $Sheet1.cells.item(1,42) = 'Erz1Hauptansprechpartner'
    $Sheet1.cells.item(1,43) = 'Erz2Art'
    $Sheet1.cells.item(1,44) = 'Erz2Anrede'
    $Sheet1.cells.item(1,45) = 'Erz2Name'
    $Sheet1.cells.item(1,46) = 'Erz2Vorname'
    $Sheet1.cells.item(1,47) = 'Erz2Strasse'
    $Sheet1.cells.item(1,48) = 'Erz2Hausnr'
    $Sheet1.cells.item(1,49) = 'Erz2PLZ'
    $Sheet1.cells.item(1,50) = 'Erz2Ort'
    $Sheet1.cells.item(1,51) = 'Erz2Teilort'
    $Sheet1.cells.item(1,52) = 'Erz2Telefon'
    $Sheet1.cells.item(1,53) = 'Erz2Handy'
    $Sheet1.cells.item(1,54) = 'Erz2Email'
    $Sheet1.cells.item(1,55) = 'Erz2Schriftverkehrverteiler'
    $Sheet1.cells.item(1,56) = 'Erz2auskunftsberechtigt'
    $Sheet1.cells.item(1,57) = 'Erz2Hauptansprechpartner'
    $Sheet1.cells.item(1,58) = 'Fremdsprache1'
    $Sheet1.cells.item(1,59) = 'Fremdsprache2'
    $Sheet1.cells.item(1,60) = 'Fremdsprache3'
    $Sheet1.cells.item(1,61) = 'Fremdsprache4'
    $Sheet1.cells.item(1,62) = 'Zuzugsart'
    $Sheet1.cells.item(1,63) = 'AbgebendeSchule'
    $Sheet1.cells.item(1,64) = 'Ausbildungsbetrieb'
    $Sheet1.cells.item(1,65) = 'Ausbild_beruf_id'
    $Sheet1.cells.item(1,66) = 'AusbArtDerBescheaftigung'
    $Sheet1.cells.item(1,67) = 'VorbldgSchulischAbschluss'
    $Sheet1.cells.item(1,68) = 'VorbldgSchulart'
    $Sheet1.cells.item(1,69) = 'SchulbesuchAmStichtag'
    $Sheet1.cells.item(1,70) = 'ArtDerHeimunterbringung'
    $Sheet1.cells.item(1,71) = 'Unterbringung'
    $Sheet1.cells.item(1,72) = 'Gastschï¿½ler'
    $Sheet1.cells.item(1,73) = 'Zuzugsdatum'
    $Sheet1.cells.item(1,74) = 'Austrittsdatum'
    $Sheet1.cells.item(1,75) = 'BeruflicherAbschluss'
    $Sheet1.cells.item(1,76) = 'Lokales_DM'

    ### Exceldatei speichern
    $MyExcel.displayalerts = $false                       # Ueberschreibt bestehnde Datei ohne Rueckfrage
    $Myworkbook.Saveas($outputExcelXLSXPath)
    $MyExcel.displayalerts = $true
    $Myworkbook.close($true)
    $MyExcel.Quit()
    $Myexcel.displayalerts = $true
    Remove-Variable MyExcel
}

function AddContentExcelFile {

    $konfessionsZuordnung = Import-Excel -Path $ConfigPath -WorksheetName Konfession
    $berufsZuordnung = Import-Excel -Path $ConfigPath -WorksheetName BerufsID
    $staatZuordnung = Import-Excel -Path $ConfigPath -WorksheetName Staatsangehoerigkeit
    $reliUnterrichtZuordnung = Import-Excel -Path $ConfigPath -WorksheetName Reliunterricht
    
    $Myexcel = New-Object -ComObject excel.application
    $Myexcel.visible = $visibleExcel
    $Myworkbook = $Myexcel.Workbooks.Open($outputExcelXLSXPath)
    $Sheet1 = $Myworkbook.worksheets.item(1)
    $Sheet1.name = "ASV"

    $berufMissing = @()
    $konfessionsMissing = @()
    $staatZuordnungMissing = @()

    $i = 2
    foreach ($xml in $arrayXML){
    
        [xml] $xmlContent = Get-Content -path $xmlPath$xml -Encoding UTF8
        write-host  Bearbeite Schueler $xmlContent.myForm.azubi_familienname
        $Sheet1.cells.item($i,1) = $klasse
        $Sheet1.cells.item($i,2) = $xmlContent.myForm.azubi_familienname
        $Sheet1.cells.item($i,3) = $xmlContent.myForm.azubi_vorname
        $Sheet1.cells.item($i,4) = $xmlContent.myForm.azubi_vorname                 # eventuell Rufname weglassen
        $Sheet1.cells.item($i,5) = $xmlContent.myForm.azubi_geburtsdatum
        $Sheet1.cells.item($i,6) = 'G'
        $Sheet1.cells.item($i,7) = $xmlContent.myForm.azubi_geburtsort
        $Sheet1.cells.item($i,8) = $xmlContent.myForm.azubi_geburtsland
        $Sheet1.cells.item($i,9) = $xmlContent.myForm.azubi_geschlecht

        ### Konfessionszuordnung
        # $Sheet1.cells.item($i,10) = $xmlContent.myForm.azubi_religion
        if ($konfessionsZuordnung.'XML-Konfession' -contains $xmlContent.myForm.azubi_religion){
            $Sheet1.cells.item($i,10) = ($konfessionsZuordnung |Where-Object XML-Konfession -eq $xmlContent.myForm.azubi_religion).'ASV-Reli-Kurzform'
        } else {
            if ($konfessionsMissing -notcontains $xmlContent.myForm.azubi_religion){
                $konfessionsMissing += $xmlContent.myForm.azubi_religion
            }

        }



        ### Relligionsunterricht
        # $Sheet1.cells.item($i,11) = $xmlContent.myForm.'reliunterricht'
        if ($reliUnterrichtZuordnung.'XML-Konfession' -contains $xmlContent.myForm.azubi_religion){
            $Sheet1.cells.item($i,11) = ($reliUnterrichtZuordnung |Where-Object XML-Konfession -eq $xmlContent.myForm.azubi_religion).'ASV-RU'
        } else {
            $Sheet1.cells.item($i,11) = ($reliUnterrichtZuordnung |Where-Object XML-Konfession -eq default).'ASV-RU'
        }



        ### Staatsangehoerigkeit fuer ASV aendern
        # $Sheet1.cells.item($i,12) = $xmlContent.myForm.azubi_staatsangehoerigkeit
        if ($staatZuordnung.'StaatsangehoerigkeitFormular' -contains $xmlContent.myForm.azubi_staatsangehoerigkeit){
            $Sheet1.cells.item($i,12) = ($staatZuordnung |Where-Object StaatsangehoerigkeitFormular -eq $xmlContent.myForm.azubi_staatsangehoerigkeit).'ASV-Staats-Kurzform'
        } else {
            if ($staatZuordnungMissing -notcontains $xmlContent.myForm.azubi_staatsangehoerigkeit){
                $staatZuordnungMissing += $xmlContent.myForm.azubi_staatsangehoerigkeit
            }

        }




        $Sheet1.cells.item($i,14) = $xmlContent.myForm.azubi_strasse
        $Sheet1.cells.item($i,15).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,15) = $xmlContent.myForm.azubi_hnr
        $Sheet1.cells.item($i,16).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,16) = $xmlContent.myForm.azubi_plz
        $Sheet1.cells.item($i,17) = $xmlContent.myForm.azubi_wohnort
        $Sheet1.cells.item($i,19) = 'Deutschland'                                    # Staat
        $Sheet1.cells.item($i,20).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,20) = $xmlContent.myForm.azubi_telefon
        # $Sheet1.cells.item($i,21) = $xmlContent.myForm.'Handy'                     # Azubi Handynummer vs Telefon
        $Sheet1.cells.item($i,22) = $xmlContent.myForm.azubi_email
        # $Sheet1.cells.item($i,23) = 'Muttersprache'                                # Muttersprache ? 
        $Sheet1.cells.item($i,24) = $xmlContent.myForm.ausbildung_beginn             # Schuleintritt
        # $Sheet1.cells.item($i,25) = $xmlContent.myForm.''                          # Einschulung am
        # $Sheet1.cells.item($i,26) = 'im Schriftverkehrverteiler'                   # brauchen wir das?


        ################  Erziehungsberechtigter 1
        # $Sheet1.cells.item($i,27) = 'auskunftsberechtigt'                          # brauchen wir das?
        # $Sheet1.cells.item($i,28) = $xmlContent.myForm.person1_form
        # $Sheet1.cells.item($i,29) = $xmlContent.myForm.'Erz1Anrede'                # haben wir das?
        $Sheet1.cells.item($i,30) = $xmlContent.myForm.person1_familienname
        $Sheet1.cells.item($i,31) = $xmlContent.myForm.person1_vorname
        if ($xmlContent.myForm.person1_adressegleich -eq $true) {
            $Sheet1.cells.item($i,32) = $xmlContent.myForm.azubi_strasse
            $Sheet1.cells.item($i,33).NumberFormatlocal = "@"
            $Sheet1.cells.item($i,33) = $xmlContent.myForm.azubi_hnr
            $Sheet1.cells.item($i,34).NumberFormatlocal = "@"
            $Sheet1.cells.item($i,34) = $xmlContent.myForm.azubi_plz
            $Sheet1.cells.item($i,35) = $xmlContent.myForm.azubi_wohnort
        } else {
            $Sheet1.cells.item($i,32) = $xmlContent.myForm.person1_strasse
            $Sheet1.cells.item($i,33) = $xmlContent.myForm.person1_hnr
            $Sheet1.cells.item($i,34) = $xmlContent.myForm.person1_plz
            $Sheet1.cells.item($i,35) = $xmlContent.myForm.person1_wohnort
        }
        $Sheet1.cells.item($i,37).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,37) = $xmlContent.myForm.person1_telefon
        $Sheet1.cells.item($i,38).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,38) = $xmlContent.myForm.person1_handy
        $Sheet1.cells.item($i,39) = $xmlContent.myForm.person1_email
        # $Sheet1.cells.item($i,40) = $xmlContent.myForm.'Erz1Schriftverkehrverteiler'
        # $Sheet1.cells.item($i,41) = $xmlContent.myForm.'Erz1auskunftsberechtigt'
        # $Sheet1.cells.item($i,42) = $xmlContent.myForm.'Erz1Hauptansprechpartner'


        ################  Erziehungsberechtigter 2
        # $Sheet1.cells.item($i,43) = $xmlContent.myForm.person2_form
        # $Sheet1.cells.item($i,44) = $xmlContent.myForm.'Erz2Anrede'                # haben wir das?
        $Sheet1.cells.item($i,45) = $xmlContent.myForm.person2_familienname
        $Sheet1.cells.item($i,46) = $xmlContent.myForm.person2_vorname
        if ($xmlContent.myForm.person1_adressegleich -eq $true) {
            $Sheet1.cells.item($i,47) = $xmlContent.myForm.azubi_strasse
            $Sheet1.cells.item($i,48).NumberFormatlocal = "@"
            $Sheet1.cells.item($i,48) = $xmlContent.myForm.azubi_hnr
            $Sheet1.cells.item($i,49).NumberFormatlocal = "@"
            $Sheet1.cells.item($i,49) = $xmlContent.myForm.azubi_plz
            $Sheet1.cells.item($i,50) = $xmlContent.myForm.azubi_wohnort
        } else {
            $Sheet1.cells.item($i,47) = $xmlContent.myForm.person2_strasse
            $Sheet1.cells.item($i,48).NumberFormatlocal = "@"
            $Sheet1.cells.item($i,48) = $xmlContent.myForm.person2_hnr
            $Sheet1.cells.item($i,49).NumberFormatlocal = "@"
            $Sheet1.cells.item($i,49) = $xmlContent.myForm.person2_plz
            $Sheet1.cells.item($i,50) = $xmlContent.myForm.person2_wohnort
        }
        $Sheet1.cells.item($i,52).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,52) = $xmlContent.myForm.person2_telefon
        $Sheet1.cells.item($i,53).NumberFormatlocal = "@"
        $Sheet1.cells.item($i,53) = $xmlContent.myForm.person2_handy
        $Sheet1.cells.item($i,54) = $xmlContent.myForm.person2_email
        # $Sheet1.cells.item($i,55) = 'Erz2Schriftverkehrverteiler'
        # $Sheet1.cells.item($i,56) = $xmlContent.myForm.'Erz2auskunftsberechtigt'
        # $Sheet1.cells.item($i,57) = $xmlContent.myForm.'Erz2Hauptansprechpartner'


        ################  Weitere Info

        # $Sheet1.cells.item($i,62) = $xmlContent.myForm.'Zuzugsart'
        $Sheet1.cells.item($i,64) = $xmlContent.myForm.unternehmen_id_label

        ### Ausbildungsberuf in BerufsID ï¿½ndern
        if ($berufsZuordnung.BerufsbezeichnungFormular -contains $xmlContent.myForm.beruf_name){
            $Sheet1.cells.item($i,65) = ($berufsZuordnung |Where-Object BerufsbezeichnungFormular -eq $xmlContent.myForm.beruf_name).'ASV-Berufs-ID'
        } else {
            if ($berufMissing -notcontains $xmlContent.myForm.beruf_name){
                $berufMissing += $xmlContent.myForm.beruf_name
            }

        }

        # $Sheet1.cells.item($i,66) = $xmlContent.myForm.'AusbArtDerBescheaftigung'
        # $Sheet1.cells.item($i,67) = $xmlContent.myForm.'VorbldgSchulischAbschluss'
        # $Sheet1.cells.item($i,68) = $xmlContent.myForm.'VorbldgSchulart'
        # $Sheet1.cells.item($i,69) = $xmlContent.myForm.'SchulbesuchAmStichtag'
        # $Sheet1.cells.item($i,70) = 'ArtDerHeimunterbringung'
        # $Sheet1.cells.item($i,71) = 'Unterbringung'
        # $Sheet1.cells.item($i,72) = $xmlContent.myForm.'Gastschüler'
        # $Sheet1.cells.item($i,73) = $xmlContent.myForm.'Zuzugsdatum'
        $Sheet1.cells.item($i,74) = $xmlContent.myForm.ausbildung_ende
        # $Sheet1.cells.item($i,75) = 'BeruflicherAbschluss'



        ######## Hochschule Dual Exportieren
        if ($xmlContent.myForm.'ausbildung_dual' -eq 'true') {
            $hdschueler = $xmlContent.myForm.'azubi_familienname' + ", " + $xmlContent.myForm.'azubi_vorname' + ", " + $xmlContent.myForm.'beruf_name' + ", " + $xmlContent.myForm.'unternehmen_id_label'
            $hdschueler | Export-Excel -path .\HD-schueler.xlsx -Append
        }




        ################ XML-Datei in Archiv verschieben
        Move-Item $xmlPath$xml $xmlArchivPath$xml                                               


        $i++
    }

    # Exceldatei speichern
    $Myexcel.displayalerts = $false
    $Myworkbook.Save()
    $Myworkbook.close($true)
    $MyExcel.Quit()
    
    $Myexcel.displayalerts = $true

    $berufMissing | Export-Excel -path $ConfigPath -WorksheetName BerufsID -Append                                                   # Ergänzt die config-Datei um fehlende Berufe
    $konfessionsMissing | Export-Excel -path $ConfigPath -WorksheetName Konfession -Append                                           # Ergänzt die config-Datei um fehlende Konfessionszuordnungen 
    $staatZuordnungMissing | Export-Excel -path $ConfigPath -WorksheetName Staatsangehoerigkeit -Append                              # Ergänzt die config-Datei um fehlende Staatsangehï¿½rigkeitszuordnungen 

    Remove-Variable MyExcel

}
