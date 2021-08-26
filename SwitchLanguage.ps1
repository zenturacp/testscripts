$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/SwitchLanguage/switchlang.exe?st=2020-04-27T13%3A07%3A01Z&se=2030-04-28T13%3A07%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=uX%2BG35U4lP8eq%2BjvWozW6M12QPJc2aMlOQ%2F4KdOnBDw%3D'
$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"

$Product = "Langauge Switcher"
$SetupFile = "switchlang.exe"
$CurrentDir = Get-Location

if (!(Test-Path $InstallTempFolder)) {
    New-Item -ItemType Directory -Force -Path $InstallTempFolder | Out-Null
}

Write-Host "$Product`: Downloading installer to $InstallTempFolder\$SetupFile"
(New-Object System.Net.WebClient).DownloadFile($url, "$InstallTempFolder\$SetupFile")

if (!(Test-Path "$InstallTempFolder\$SetupFile")) {
    Write-Error "$Product`: Installer not downloaded"
    exit 1
}

Set-Location $InstallTempFolder

if (!(Test-Path "$Env:ProgramData\BC365"))
{
    New-Item -ItemType Directory -Path "$Env:ProgramData\BC365" -Force | Out-Null
}

Write-Host "$Product`: Extracting files from $SetupFile to $InstallTempFolder"
Start-Process -FilePath "$InstallTempFolder\$SetupFile" -Argumentlist "-o$InstallTempFolder -y" -NoNewWindow -Wait

if (Test-Path "$InstallTempFolder\$SetupFile") {
    Write-Host "$Product`: Moving powershell script to $Env:ProgramData\BC365"
    Get-ChildItem -Path $InstallTempFolder -Filter "*.ps1" | Move-Item -Destination "$Env:ProgramData\BC365"
    Write-Host "$Product`: Adding StartMenu Shortcut"
    Get-ChildItem -Path $InstallTempFolder -Filter "*.lnk" | Move-Item -Destination "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs"
}

Set-Location $CurrentDir

Write-Host "$Product`: Performing cleanup"
Remove-Item -Path $InstallTempFolder -Recurse -Force
