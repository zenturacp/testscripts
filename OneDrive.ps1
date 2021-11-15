$Url = 'https://bc365v834p.blob.core.windows.net/software/onedrive/21.205.1003.3/OneDriveSetup.exe?sv=2020-04-08&st=2021-11-15T07%3A31%3A15Z&se=2031-11-16T07%3A31%3A00Z&sr=b&sp=r&sig=jI8wc0rFeybetVwHn4heYLZr6TXsyGPdFlMv0mqMg30%3D'
$Product = 'OneDrive'
$SetupFile = 'OneDriveSetup.exe'
$SetupArguments = @(
    "/allusers",
    "/silent"
)

$TenantID = "cef642e3-be23-4a67-9261-f7cc5db4cddb"

#$TenantAutoMounts = @{
#    'Accounting'='tenantId=cef642e3%2Dbe23%2D4a67%2D9261%2Df7cc5db4cddb&siteId=%7B5967e153%2De504%2D4eac%2Db4ae%2Dbf7463b4bcb8%7D&webId=%7B48ac3bb7%2D410e%2D4343%2D99d7%2Dfe353c97be11%7D&listId=%7B244BD895%2D881E%2D41D6%2D8BD8%2D0766027A6F32%7D&webUrl=https%3A%2F%2Fscanmetals%2Esharepoint%2Ecom%2Fsites%2Fsmfiles&version=1';
#    'Data UK'='tenantId=cef642e3%2Dbe23%2D4a67%2D9261%2Df7cc5db4cddb&siteId=%7B5967e153%2De504%2D4eac%2Db4ae%2Dbf7463b4bcb8%7D&webId=%7B48ac3bb7%2D410e%2D4343%2D99d7%2Dfe353c97be11%7D&listId=%7B53043921%2DE9E1%2D4D02%2D955B%2D2EB31F3FA22D%7D&webUrl=https%3A%2F%2Fscanmetals%2Esharepoint%2Ecom%2Fsites%2Fsmfiles&version=1';
#    'MIS UK'='tenantId=cef642e3%2Dbe23%2D4a67%2D9261%2Df7cc5db4cddb&siteId=%7B5967e153%2De504%2D4eac%2Db4ae%2Dbf7463b4bcb8%7D&webId=%7B48ac3bb7%2D410e%2D4343%2D99d7%2Dfe353c97be11%7D&listId=%7BF9BBF9FE%2D6D07%2D495D%2DAD25%2D0425EE2F7AEA%7D&webUrl=https%3A%2F%2Fscanmetals%2Esharepoint%2Ecom%2Fsites%2Fsmfiles&version=1';
#    'Shared'='tenantId=cef642e3%2Dbe23%2D4a67%2D9261%2Df7cc5db4cddb&siteId=%7B5967e153%2De504%2D4eac%2Db4ae%2Dbf7463b4bcb8%7D&webId=%7B48ac3bb7%2D410e%2D4343%2D99d7%2Dfe353c97be11%7D&listId=%7BD7CDC935%2DB11F%2D4E07%2D811C%2D1360CEEF7C7E%7D&webUrl=https%3A%2F%2Fscanmetals%2Esharepoint%2Ecom%2Fsites%2Fsmfiles&version=1'
#}

$InstallTempFolder = "D:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"
$CurrentDir = Get-Location

if (!(Test-Path $InstallTempFolder)) {
    New-Item -ItemType Directory -Force -Path $InstallTempFolder | Out-Null
}

Write-Host "Downloading $Product installer to $InstallTempFolder\$SetupFile"
(New-Object System.Net.WebClient).DownloadFile($Url, "$InstallTempFolder\$SetupFile")

if (!(Test-Path "$InstallTempFolder\$SetupFile")) {
    Write-Error "$Product installer not downloaded"
    exit 1
}

Set-Location $InstallTempFolder
Write-Host "Starting $Product installer"

if (Test-Path "$InstallTempFolder\$SetupFile") {
    Start-Process -FilePath "$InstallTempFolder\$SetupFile" -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

# OneDrive Default Settings
Write-Host "Setting Default Settings in Registry"
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Force | Out-Null
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\AllowTenantList" -Force | Out-Null
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantAutoMount" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "DehydrateSyncedTeamSites" -Value 1 -Type DWord
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "FilesOnDemandEnabled" -Value 1 -Type DWord
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "KFMSilentOptIn" -Value $TenantID -Type String
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "KFMSilentOptInWithNotification" -Value 0 -Type DWord
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "SilentAccountConfig" -Value 1 -Type DWord
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\AllowTenantList" -Name $TenantID -Value $TenantID -Type String
#foreach($TenantAutoMount in $TenantAutoMounts.GetEnumerator())
#{
#    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantAutoMount" -Name $TenantAutoMount.Name -Value $TenantAutoMount.Value -Type String
#}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
