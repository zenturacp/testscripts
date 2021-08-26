$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/SharePointLink/SharePointLink.exe?sv=2019-02-02&st=2020-05-06T18%3A55%3A18Z&se=2030-05-07T18%3A55%3A00Z&sr=b&sp=r&sig=qNQ0r%2Fpim6CcxOY0szzkM%2BBsmLoBpqKjKYaLYCB6jeI%3D'
$Product = 'Sharepoint Shortcut'
$SetupFile = 'SharePointLink.exe'

$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"
$CurrentDir = Get-Location

if (!(Test-Path $InstallTempFolder))
{
    new-item -ItemType Directory -Force -Path $InstallTempFolder | Out-Null
}

Write-Host "Downloading $Product installer to $InstallTempFolder\$SetupFile"
(New-Object System.Net.WebClient).DownloadFile($Url, "$InstallTempFolder\$SetupFile")

if (!(Test-Path "$InstallTempFolder\$SetupFile"))
{
    Write-Error "$Product installer not downloaded"
    exit 1
}

Set-Location $InstallTempFolder

Write-Host "$Product`: Extracting files from $SetupFile to $InstallTempFolder"
Start-Process -FilePath "$InstallTempFolder\$SetupFile" -Argumentlist "-o$InstallTempFolder -y" -NoNewWindow -Wait

if (!(Test-Path "$Env:ProgramData\BC365"))
{
    New-Item -ItemType Directory -Path "$Env:ProgramData\BC365" -Force | Out-Null
}

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    Copy-Item -Path "$InstallTempFolder\Sharepoint.ico" -Destination "$Env:ProgramData\BC365"
    Copy-Item -Path "$InstallTempFolder\SharePoint ScanMetals.url" -Destination "$Env:Public\Desktop\"
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
