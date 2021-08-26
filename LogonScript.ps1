$WallPaperUrl = 'https://bc365v834p.blob.core.windows.net/config/1003/Wallpaper.png?sp=r&st=2020-03-31T13:15:16Z&se=2030-03-31T21:15:16Z&spr=https&sv=2019-02-02&sr=b&sig=J5oVrTDLZUT6VJ284YOlQVQCkWdNeaKO7z7oaPkPVhM%3D'
$WallPaperFile = "$Env:ProgramData\BC365\Wallpaper.png"
Write-Host "Installing Logon Script (Pr. User)"

$ScriptBlock = @"
# Set Registry Keys
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_NotifyNewApps" -Type DWord -Value 0 | Out-Null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "StartMenuAdminTools" -Type DWord -Value 0 | Out-Null
Set-ItemProperty -path "HKCU:\Control Panel\Desktop\" -name wallpaper -value %%WallPaper%%
rundll32.exe user32.dll, UpdatePerUserSystemParameters
"@

if (!(Test-Path "$Env:ProgramData\BC365"))
{
    New-Item -ItemType Directory -Path "$Env:ProgramData\BC365" -Force | Out-Null
}

Write-Host "Downloading Wallpaper to $WallPaperFile"
(New-Object System.Net.WebClient).DownloadFile($WallpaperUrl, $WallPaperFile)

$ScriptBlock.Replace("%%WallPaper%%","`"$WallPaperFile`"") | Out-File -FilePath "$Env:ProgramData\BC365\LogonScript.ps1" -Force
#Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "LogonScript" -Value "powershell.exe -file %ProgramData%\BC365\LogonScript.ps1 -executionpolicy bypass -windowstyle hidden -noprofile" -Type ExpandString

$Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File $Env:ProgramData\BC365\LogonScript.ps1"
$Trigger =  New-ScheduledTaskTrigger -AtLogon
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -Hidden -DontStopIfGoingOnBatteries -Compatibility Win8
$Principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545"
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal
Register-ScheduledTask -InputObject $Task -TaskName "LogonScript" -Force