param ([string]$dateiname)
Write-Host "Input Datei: $dateiname"

$delimiter = ";"
#Zulässige Werte: ";",","
$encoding = "UTF8"
#Zulässige Werte: ASCII, BigEndianUnicode, BigEndianUTF32, OEM, Unicode, UTF7, UTF8, UTF8BOM, UTF8NoBOM, UTF32
$useQuotes = "Never"
#This parameter was added in PowerShell 7.0. Zulässige Werte: Never, Always, AsNeeded
$dateiname_ohne_csv = $dateiname.Substring(0, $dateiname.Length - 4)
$ausgabe_AD = $dateiname_ohne_csv+"_AD.csv"
$ausgabe_Moodle = $dateiname_ohne_csv+"_Moodle.csv"

Write-Host "ausgabe_AD: " $ausgabe_AD
Write-Host "ausgabe_Moodle: " $ausgabe_Moodle

function ConvertText{
	param ([string]$Text)
	$Text = $Text -replace 'ä','ae' -replace 'ö','oe' -replace 'ü','ue' -replace 'ß','ss' -replace 'Ä','Ae' -replace 'Ö','Oe' -replace 'Ü','Ue' -replace 'ć','c' -replace 'ǎ','a' -replace 'ễ','e' -replace 'ğ','g'
	return $Text
}

$daten = Import-Csv -Path $dateiname -Delimiter $delimiter -Encoding $encoding

$ergebnis_AD = foreach ($zeile in $daten) {
	$vorname = ConvertText $zeile.Vorname
	$nachname = ConvertText $zeile.Nachname

	[PSCustomObject]@{
		S_K = $zeile.Klassenname
		Name = $nachname
		Vorname = $vorname
		Geburtsdat = $zeile.Geburtsdatum
	}
}

$ergebnis_Moodle = foreach ($zeile in $daten) {
	$vorname = ConvertText $zeile.Vorname
	$nachname = ConvertText $zeile.Nachname
	
	#Benutzername: erste 6 Zeichen des Nachnamens + erste 3 Zeichen des Vornamens
	$username = ($nachname.Substring(0, [Math]::Min(6, $nachname.Length)) + $vorname.Substring(0, [Math]::Min(3, $vorname.Length))).ToLower()

	[PSCustomObject]@{
		username = $username
		firstname = $vorname
		lastname = $nachname
		email = "$username@bs19hh.de"
		password = $zeile.Geburtsdatum
		cohort1 = "5916_bs19"
		cohort2 = "5916_bs_" + $zeile.Klassenname
		profile_field_schoolno = "5916"
		profile_field_schoolname = "Berufliche Schule Farmsen BS19"
		profile_field_schoolpersona = "Schüler/in"
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
