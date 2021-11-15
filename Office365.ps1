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

if (Test-Path "$InstallTempFolder\setup.exe") {
    Write-Host "$Product`: Downloading files for installer"
    Start-Process -FilePath "$InstallTempFolder\setup.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/download configuration-Office365-x64.xml" -NoNewWindow -Wait
    Write-Host "$Product`: Running installer"
    Start-Process -FilePath "$InstallTempFolder\setup.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/configure configuration-Office365-x64.xml" -NoNewWindow -Wait
}

Set-Location $CurrentDir

# Setting up GPO
$keyName = Get-Random -Minimum 100000 -Maximum 999999
reg load "HKU\$keyName" "C:\Users\Default\NTUSER.DAT"
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS

# Default Profile
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\general") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\general" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\graphics") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\graphics" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\signin") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\signin" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\excel") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\excel" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\excel\options") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\excel\options" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\firstrun") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\firstrun" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\search") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\search" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\security") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\security" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\powerpoint") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\powerpoint" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\powerpoint\options") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\powerpoint\options" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\word") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\word" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\word\options") -ne $true) { New-Item "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\word\options" -force -ea SilentlyContinue };
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\general" -Name "skydrivesigninoption" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\general" -Name "optindisable" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\graphics" -Name "disablehardwareacceleration" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\common\signin" -Name "signinoptions" -Value 2 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\excel\options" -Name "disableboottoofficestart" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\firstrun" -Name "disablemovie" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\firstrun" -Name "bootedrtm" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm" -Name "enableupload" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm" -Name "enablelogging" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "wdsolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "xlsolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "pptsolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "olksolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "accesssolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "projectsolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "publishersolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "visiosolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedapplications" -Name "onenotesolution" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes" -Name "documentfiles" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes" -Name "templatefiles" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes" -Name "comaddins" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes" -Name "appaddins" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\osm\preventedsolutiontypes" -Name "agave" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -Name "cachedexchangemode" -Value 2 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -Name "syncwindowsetting" -Value 6 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -Name "syncwindowsettingdays" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -Name "enable" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -Name "syncpffav" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\cached mode" -Name "downloadsharedfolders" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "doaging" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "everydays" -Value 14 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "promptforaging" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "deleteexpired" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "archiveold" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "archivemount" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "archiveperiod" -Value 6 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "archivegranularity" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name "archivedelete" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\search" -Name "disabledownloadsearchprompt" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\outlook\security" -Name "markinternalasunsafe" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\powerpoint\options" -Name "disableboottoofficestart" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKU:\$keyName\Software\Policies\Microsoft\office\16.0\word\options" -Name "disableboottoofficestart" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;

# Local Machine
if ((Test-Path -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office") -ne $true) { New-Item "HKLM:\SOFTWARE\Policies\Microsoft\office" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0") -ne $true) { New-Item "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common") -ne $true) { New-Item "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common" -force -ea SilentlyContinue };
if ((Test-Path -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate") -ne $true) { New-Item "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" -force -ea SilentlyContinue };
New-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" -Name "enableautomaticupdates" -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" -Name "hideenabledisableupdates" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" -Name "hideupdatenotifications" -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;

# Unload default profile again
reg unload "HKU\$keyName"
Remove-PSDrive -Name "HKU"

Remove-Item -Path $InstallTempFolder -Recurse -Force
