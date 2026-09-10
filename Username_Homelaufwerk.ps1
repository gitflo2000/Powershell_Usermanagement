#Write-Host "Username: $env:UserName"
$cred = Get-Credential -Credential Contoso\ServiceAccount
New-PSDrive -Name "H" -Root "\\bs19hh.de\Unterricht\Lehrerdaten\$env:UserName" -Persist -PSProvider "FileSystem" -Credential $cred
Net Use