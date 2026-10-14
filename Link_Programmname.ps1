$WscriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WscriptShell.CreateShortcut("c:\Users\Public\Desktop\Programmname.lnk")
$Shortcut.TargetPath = "c:\Program Files (x86)\Hersteller\Programmname\Programmname.exe"
$Shortcut.IconLocation = "c:\Program Files (x86)\Hersteller\Programmname\Programmname.exe,0" # Optional: Spezifisches Icon
$Shortcut.Save()