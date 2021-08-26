$Url = 'https://bc365v834p.blob.core.windows.net/software/teamviewer/15.3.8497.0/TeamViewerQS.exe?st=2020-04-08T10%3A47%3A31Z&se=2030-04-09T10%3A47%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=Mj58UMhx0xq2Rm6LfEZ%2Fb4P4LwY%2B8EyqPPcGXA42QFY%3D'

$DownloadFolder = "$Env:ProgramFiles\TeamviewerQS"
$Output = "$DownloadFolder\TeamviewerQS.exe"
$ShortcutLocation = "$Env:PUBLIC\Desktop\TeamViewer Zentura.lnk"

if (!(Test-Path $DownloadFolder))
{
    new-item -ItemType Directory -Force -Path $DownloadFolder | Out-Null
}

Write-Host "Downloading TeamviewerQS.exe to $Output"
(New-Object System.Net.WebClient).DownloadFile($url, $output)

$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
$Shortcut.TargetPath = $Output
$Shortcut.Save()
