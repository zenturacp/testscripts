$Url = 'https://bc365v834p.blob.core.windows.net/software/teamviewer/15.23.9.0/TeamViewerQS.exe?sv=2020-04-08&st=2021-11-15T07%3A36%3A53Z&se=2031-11-16T07%3A36%3A00Z&sr=b&sp=r&sig=shzt6FF7on3Z8DEkBdn7gBal36Hg8%2FGudo8IcUVY7zM%3D'

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
