$Url = 'https://bc365v834p.blob.core.windows.net/software/nemid/nemid.zip?sp=r&st=2021-08-27T09:17:31Z&se=2031-08-27T17:17:31Z&spr=https&sv=2020-08-04&sr=b&sig=yHHj9Z%2Bwc%2Fn3UOOGXDYCAPaGj0MsQGnMVtiMevLo60A%3D'
$Product = 'NemID'
$ArchiveFile = 'nemid.zip'

$installers = @(
    "csp.msi",
    "NemidNoglefilsprogram.msi",
    "NemID KSP 1.5.msi"
)

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

if (Test-Path "$InstallTempFolder\$ArchiveFile") {
    foreach ($installer in $installers) {
        Write-Host "$Product`: Starting installer $installer"
        Start-Process -FilePath "msiexec.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/i `"$installer`" /qn" -NoNewWindow -Wait
    }
}

Set-Location $CurrentDir
Remove-Item -Path $InstallTempFolder -Recurse -Force
