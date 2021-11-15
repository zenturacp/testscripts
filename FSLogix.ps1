$Url = 'https://bc365v834p.blob.core.windows.net/software/fslogix/2.9.7979.62170/x64/Release/FSLogixAppsSetup.exe?sv=2020-04-08&st=2021-11-15T07%3A20%3A38Z&se=2031-11-16T07%3A20%3A00Z&sr=b&sp=r&sig=KMdqreYoWJiaVRdyxiqO09JFfgBwF4FzXxN1SM%2BzLGA%3D'
$Product = 'FSLogix'
$SetupFile = 'FSLogixAppsSetup.exe'
$VHDLocation = '\\wlhvx1003profiles.file.core.windows.net\profiles\desktop'
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
