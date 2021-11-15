$Url = 'https://bc365v834p.blob.core.windows.net/software/office365/ODTBusiness.zip?sv=2020-04-08&st=2021-11-15T07%3A26%3A12Z&se=2031-11-16T07%3A26%3A00Z&sr=b&sp=r&sig=HhaO0NgeUADqNRyJ%2FltRKn13a78mDsm%2Bf0kBVEgXwuM%3D'
$Product = 'Office365 Business'
$ArchiveFile = 'ODTBusiness.zip'

$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"
$CurrentDir = Get-Location

if (!(Test-Path $InstallTempFolder)) {
    new-item -ItemType Directory -Force -Path $InstallTempFolder | Out-Null
}

Write-Host "$Product`: Downloading installer to $InstallTempFolder\$ArchiveFile"
(New-Object System.Net.WebClient).DownloadFile($Url, "$InstallTempFolder\$ArchiveFile")

if (!(Test-Path "$InstallTempFolder\$ArchiveFile")) {
    Write-Error "$Product`: installer not downloaded"
    exit 1
}

Set-Location $InstallTempFolder

Write-Host "$Product`: Extracting files from $ArchiveFile to $InstallTempFolder"
Expand-Archive -Path "$InstallTempFolder\$ArchiveFile" -DestinationPath $InstallTempFolder

if (Test-Path "$InstallTempFolder\setup.exe")
{
    Write-Host "$Product`: Downloading files for installer"
    Start-Process -FilePath "$InstallTempFolder\setup.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/download configuration-Office365-x64.xml" -NoNewWindow -Wait
    Write-Host "$Product`: Running installer"
    Start-Process -FilePath "$InstallTempFolder\setup.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/configure configuration-Office365-x64.xml" -NoNewWindow -Wait
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
