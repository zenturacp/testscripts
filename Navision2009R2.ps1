$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/NAVI2009R2/CSideClient2009R2.exe?st=2020-04-16T09%3A32%3A28Z&se=2030-04-17T09%3A32%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=nsraz1%2FH4rIC8MYbFeCkfAKFHrEHFt0%2FNwaDNRp3XOY%3D'
$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"

$Product = "Navision 2009R2"
$SetupFile = "CSideClient2009R2.exe"
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

Write-Host "$Product`: Extracting files from $SetupFile to $InstallTempFolder"
Start-Process -FilePath "$InstallTempFolder\$SetupFile" -Argumentlist "-o$InstallTempFolder -y" -NoNewWindow -Wait

if (Test-Path "$InstallTempFolder\$SetupFile") {
    Write-Host "$Product`: Running installers"
    $SetupArguments = @(
        "/i `"$InstallTempFolder\Microsoft Dynamics NAV Classic.msi`"",
        "/qn"
    )
    Write-Host "$Product`: Running vcredist_x86.exe"
    Start-Process -FilePath "$InstallTempFolder\vcredist_x86.exe" -WorkingDirectory $InstallTempFolder -ArgumentList "/q" -Wait -NoNewWindow
    Write-Host "$Product`: Running ReportViewer2008.exe"
    Start-Process -FilePath "$InstallTempFolder\ReportViewer2008.exe" -WorkingDirectory $InstallTempFolder -ArgumentList "/q" -Wait -NoNewWindow
    Write-Host "$Product`: Running Navision Classic Client installer"
    Start-Process -FilePath "msiexec.exe" -WorkingDirectory $InstallTempFolder -ArgumentList $SetupArguments -Wait -NoNewWindow
    Write-Host "$Product`: Running Document Capture Client Setup"
    Start-Process -FilePath "$InstallTempFolder\DC\Setup.exe" -WorkingDirectory "$InstallTempFolder\DC" -ArgumentList "/S /v/qn" -Wait -NoNewWindow
    Get-ChildItem -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Filter "Microsoft Dynamics NAV 2009 R2*" | Remove-Item -Force | Out-Null
    Get-ChildItem -Path "$InstallTempFolder\Shortcut" | Copy-Item -Destination "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Force | Out-Null
}

Set-Location $CurrentDir

Write-Host "$Product`: Performing cleanup"
Remove-Item -Path $InstallTempFolder -Recurse -Force
