# Written by Kevin Bein

# Variables - Path definitions
$CurrentDirectory = (Resolve-Path .)
$TargetFile =    "$CurrentDirectory\WCMS.ps1"
$ShortcutFile =  "$CurrentDirectory\WCMS.lnk"
$TaskBarShortcutFile = "$env:AppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\WCMS.lnk"
$LightModeIcon = "$CurrentDirectory\light-mode.ico"
$DarkModeIcon =  "$CurrentDirectory\dark-mode.ico"

# Change windows color mode
$AppsUseLightTheme = (Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").AppsUseLightTheme
If ($AppsUseLightTheme -eq $true) {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0
    $ColorMode = "Dark"
    $ColorModeIcon = $DarkModeIcon
} else {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 1
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 1
    $ColorMode = "Light"
    $ColorModeIcon = $LightModeIcon
}
Write-Host "Changed Windows color mode to: $ColorMode"

# Change shortcut icon
$WScriptShell = New-Object -ComObject WScript.Shell

# Change original shortcut icon
$OriginalShortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$OriginalShortcut.TargetPath = $TargetFile
$OriginalShortcut.IconLocation = $ColorModeIcon 
$OriginalShortcut.Save()

# Change task bar icon
$TaskBarShortcut = $WScriptShell.CreateShortcut($TaskBarShortcutFile)
$TaskBarShortcut.TargetPath = $TargetFile
$TaskBarShortcut.IconLocation = $ColorModeIcon 
$TaskBarShortcut.Save()