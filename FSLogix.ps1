$Url = 'https://bc365v834p.blob.core.windows.net/software/fslogix/2.9.7237.48865/x64/Release/FSLogixAppsSetup.exe?st=2020-04-14T06%3A53%3A25Z&se=2030-04-15T06%3A53%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=bGtIBwVDTQqXtLqixpmCd%2FKTA4XgA%2BGp5c9pYV5iF1U%3D'
$Product = 'FSLogix'
$SetupFile = 'FSLogixAppsSetup.exe'
$VHDLocation = '\\mftas1003profilespremium.file.core.windows.net\profiles'
$SetupArguments = @(
    "/install",
    "/quiet",
    "/norestart"
)

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
Write-Host "Starting $Product installer"

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    Start-Process -FilePath "$InstallTempFolder\$SetupFile" -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

# Create Registry Keys
New-Item -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "VHDLocations" -Type String -Value $VHDLocation
Set-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "VolumeType" -Type String -Value "vhdx"
Set-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "Enabled" -Type DWord -Value 1

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
