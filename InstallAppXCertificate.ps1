$Url = 'https://bc365v834p.blob.core.windows.net/software/zentura/ZenturaAppXCert.pfx?st=2020-04-15T13%3A19%3A36Z&se=2030-04-16T13%3A19%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=4aagupi8jjzOeKTHShliy%2Bdxb9lhuB%2FWj5zXWVgT9%2Fk%3D'
$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"
$CertFile = 'ZenturaAppXCert.pfx'
$CertPassword = ConvertTo-SecureString -String "8hbtpuhswd3jcr46" -Force -AsPlainText
$Product = "Zentura AppX Cert"

$CurrentDir = Get-Location

if (!(Test-Path $InstallTempFolder))
{
    new-item -ItemType Directory -Force -Path $InstallTempFolder | Out-Null
}

Write-Host "$Product`: Downloading $CertFile to $InstallTempFolder"
(New-Object System.Net.WebClient).DownloadFile($Url, "$InstallTempFolder\$CertFile")

if (Test-Path "$InstallTempFolder\$CertFile")
{
    Write-Host "$Product`: Importing Certificate"
    Import-PfxCertificate -FilePath "$InstallTempFolder\$CertFile" -CertStoreLocation "Cert:\LocalMachine\Root" -Password $CertPassword | Out-Null
    Write-Host "$Product`: Allowing Sideload of Applications"
    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Name "AllowAllTrustedApps" -Value 1 -Type DWord | Out-Null
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force