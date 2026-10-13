param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$dateiname
)

Write-Host "Input Datei: $dateiname"

if (-not (Test-Path -LiteralPath $dateiname -PathType Leaf)) {
    throw "Die Datei '$dateiname' wurde nicht gefunden."
}

$delimiter = ";"
#Zulässige Werte: ";",","
$encoding = "UTF8"
#Zulässige Werte: ASCII, BigEndianUnicode, BigEndianUTF32, OEM, Unicode, UTF7, UTF8, UTF8BOM, UTF8NoBOM, UTF32
$useQuotes = "Never"
#This parameter was added in PowerShell 7.0. Zulässige Werte: Never, Always, AsNeeded
$quelleIsSchuldock = $true
$createInventor = $false
$dateiname_ohne_csv = [System.IO.Path]::GetFileNameWithoutExtension($dateiname)
$ausgabe_AD = [System.IO.Path]::Combine((Split-Path -Parent $dateiname), ($dateiname_ohne_csv + "_AD.csv"))
$ausgabe_Moodle = [System.IO.Path]::Combine((Split-Path -Parent $dateiname), ($dateiname_ohne_csv + "_Moodle.csv"))
$ausgabe_Inventor = [System.IO.Path]::Combine((Split-Path -Parent $dateiname), ($dateiname_ohne_csv + "_Inventor.csv"))

Write-Host "ausgabe_AD: " $ausgabe_AD
Write-Host "ausgabe_Moodle: " $ausgabe_Moodle
Write-Host "ausgabe_Inventor: " $ausgabe_Inventor

function ConvertText {
    param ([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    $Text = [string]$Text
    $Text = $Text.Replace('Ä','Ae').Replace('Ö','Oe').Replace('Ü','Ue').Replace('ä','ae').Replace('ö','oe').Replace('ü','ue').Replace('ß','ss').Replace('ć','c').Replace('ǎ','a').Replace('ễ','e').Replace('ğ','g').Replace('Ë','E').Replace('ë','e').Replace('Ï','I').Replace('ï','i').Replace('À','A').Replace('à','a').Replace('Á','A').Replace('á','a').Replace('Â','A').Replace('â','a').Replace('Ç','C').Replace('ç','c').Replace('È','E').Replace('è','e').Replace('é','e').Replace('Ê','E').Replace('ê','e').Replace('Ñ','N').Replace('ñ','n').Replace('Ò','O').Replace('ò','o').Replace('Ó','O').Replace('ó','o').Replace('Ô','O').Replace('ô','o').Replace('õ','o').Replace('Ÿ','Y').Replace('ÿ','y').Replace(' ','')
    return $Text
}

function Get-Klassenname {
    param ([pscustomobject]$Zeile)

    if ($quelleIsSchuldock) {
        return $Zeile.'Klassen (aktiv)'
    }

    return $Zeile.Klassenname
}

$daten = Import-Csv -Path $dateiname -Delimiter $delimiter -Encoding $encoding

$ergebnis_AD = foreach ($zeile in $daten) {
    $vorname = ConvertText $zeile.Vorname
    $nachname = ConvertText $zeile.Nachname
    $myKlassenname = Get-Klassenname -Zeile $zeile

    [PSCustomObject]@{
        S_K = $myKlassenname
        Name = $nachname
        Vorname = $vorname
        Geburtsdat = $zeile.Geburtsdatum
    }
}

$ergebnis_Moodle = foreach ($zeile in $daten) {
    $vorname = ConvertText $zeile.Vorname
    $nachname = ConvertText $zeile.Nachname

    if ([string]::IsNullOrWhiteSpace($vorname)) { $vorname = 'Unbekannt' }
    if ([string]::IsNullOrWhiteSpace($nachname)) { $nachname = 'Unbekannt' }

    $myKlassenname = Get-Klassenname -Zeile $zeile
    $username = ($nachname.Substring(0, [Math]::Min(6, $nachname.Length)) + $vorname.Substring(0, [Math]::Min(3, $vorname.Length))).ToLower()

    [PSCustomObject]@{
        username = $username
        firstname = $vorname
        lastname = $nachname
        email = "$username@bs19hh.de"
        password = $zeile.Geburtsdatum
        cohort1 = "5916_bs19"
        cohort2 = "5916_bs_" + $myKlassenname
        profile_field_schoolno = "5916"
        profile_field_schoolname = "Berufliche Schule Farmsen BS19"
        profile_field_schoolpersona = "Schüler/in"
    }
}
If ($createInventor) {
    $ergebnis_Inventor = foreach ($zeile in $daten) {
        $vorname = ConvertText $zeile.Vorname
        $nachname = ConvertText $zeile.Nachname
        if ([string]::IsNullOrWhiteSpace($vorname)) { $vorname = 'Unbekannt' }
        if ([string]::IsNullOrWhiteSpace($nachname)) { $nachname = 'Unbekannt' }

        $username = ($nachname.Substring(0, [Math]::Min(6, $nachname.Length)) + $vorname.Substring(0, [Math]::Min(3, $vorname.Length))).ToLower()

        [PSCustomObject]@{
            Vorname = $vorname
            Nachname = $nachname
            email = "$username@bs19hh.de"
        }
    }
}

$ergebnis_AD | Export-Csv `
    -Path $ausgabe_AD `
    -Delimiter $delimiter `
    -NoTypeInformation `
    -UseQuotes $useQuotes `
    -Encoding $encoding

$ergebnis_AD | ConvertTo-Csv -NoTypeInformation -Delimiter $delimiter

$ergebnis_Moodle | Export-Csv `
    -Path $ausgabe_Moodle `
    -Delimiter $delimiter `
    -NoTypeInformation `
    -UseQuotes $useQuotes `
    -Encoding $encoding

$ergebnis_Moodle | ConvertTo-Csv -NoTypeInformation -Delimiter $delimiter

If ($createInventor) {
    $ergebnis_Inventor | Export-Csv `
        -Path $ausgabe_Inventor `
        -Delimiter $delimiter `
        -NoTypeInformation `
        -UseQuotes $useQuotes `
        -Encoding $encoding

    $ergebnis_Inventor | ConvertTo-Csv -NoTypeInformation -Delimiter $delimiter
}