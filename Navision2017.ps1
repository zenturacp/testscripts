$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/NAVI2017/NAV_10_0_18197_DK_DVD.exe?sp=r&st=2020-04-08T06:11:57Z&se=2030-04-08T14:11:57Z&spr=https&sv=2019-02-02&sr=b&sig=C67vZALzeihh%2Fjogntysyc%2FEaC6a880k6PaF%2FqK7JZo%3D'
$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"

$DownloadFolder = "$InstallTempFolder\Download"
$NavisionFolder = "$InstallTempFolder\Navi2017"
$Output = "$DownloadFolder\NAV_10_0_18197_DK_DVD.exe"
$CurrentDir = Get-Location

if (!(Test-Path $DownloadFolder)) {
    New-Item -ItemType Directory -Force -Path $DownloadFolder | Out-Null
}

if (!(Test-Path $NavisionFolder)) {
    New-Item -ItemType Directory -Force -Path $NavisionFolder | Out-Null
}

Write-Host "Downloading Navision client installer to $Output"
(New-Object System.Net.WebClient).DownloadFile($url, $output)

if (!(Test-Path $Output)) {
    Write-Error "Navision Client not downloaded"
    exit 1
}

Set-Location $NavisionFolder
Write-Host "Extracting files from $Output to $NavisionFolder"
Start-Process -FilePath $Output -Argumentlist "-o$NavisionFolder -y" -NoNewWindow -Wait

if (Test-Path "$NavisionFolder\setup.exe") {
    Start-Process -FilePath "$NavisionFolder\setup.exe" -Argumentlist "/config $NavisionFolder\NaviScanmetals.xml /quiet" -NoNewWindow -Wait
    Get-ChildItem -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Filter "Dynamics NAV 2017*" | Remove-Item -Force | Out-Null
    Get-ChildItem -Path "$InstallTempFolder\Navi2017\Shortcut" | Copy-Item -Destination "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Force | Out-Null
}

# Navision Fix
New-Item -Path "HKLM:\SOFTWARE\Microsoft\Clock" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Clock" -Name "AutoDST" -Value 1 -Type DWord

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
