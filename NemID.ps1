$Url = 'https://bc365v834p.blob.core.windows.net/software/nemid/20200424/nemid.exe?st=2020-04-24T05%3A48%3A33Z&se=2030-04-25T05%3A48%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=tFyn0xBFKF6KdI3u%2FddmOu%2FhBCV3zYxlh5nb73Ix%2Fcg%3D'
$Product = 'NemID'
$SetupFile = 'nemid.exe'

$installers = @(
    "csp.msi",
    "NemidNoglefilsprogram.msi",
    "NemID KSP 1.5.msi"
)

$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"
$CurrentDir = Get-Location

if (!(Test-Path $InstallTempFolder))
{
    new-item -ItemType Directory -Force -Path $InstallTempFolder | Out-Null
}

Write-Host "$Product`: Downloading installer to $InstallTempFolder\$SetupFile"
(New-Object System.Net.WebClient).DownloadFile($Url, "$InstallTempFolder\$SetupFile")

if (!(Test-Path "$InstallTempFolder\$SetupFile"))
{
    Write-Error "$Product`: installer not downloaded"
    exit 1
}

Set-Location $InstallTempFolder

Write-Host "$Product`: Extracting files from $SetupFile to $InstallTempFolder"
Start-Process -FilePath "$InstallTempFolder\$SetupFile" -Argumentlist "-o$InstallTempFolder -y" -NoNewWindow -Wait


if (Test-Path "$InstallTempFolder\$SetupFile")
{
    foreach ($installer in $installers)
    {
        Write-Host "$Product`: Starting installer $installer"
        Start-Process -FilePath "msiexec.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/i `"$installer`" /qn" -NoNewWindow -Wait
    }
}

Set-Location $CurrentDir
Remove-Item -Path $InstallTempFolder -Recurse -Force
