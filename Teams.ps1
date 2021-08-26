
$Url = 'https://bc365v834p.blob.core.windows.net/software/teams/Teams_windows_x64.msi?st=2020-04-17T10%3A35%3A27Z&se=2030-04-18T10%3A35%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=5ecVhQt%2FPLsqcBOmdYE73T59O8bHmcadbSGQmuygSB0%3D'

$Product = 'Teams'
$SetupFile = 'Teams_windows_x64.msi'

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

# Create Registry Keys
New-Item -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Name "IsWVDEnvironment" -Type DWord -Value 1

Set-Location $InstallTempFolder
Write-Host "Starting $Product installer"

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    $SetupArguments = @(
        "/i $InstallTempFolder\$SetupFile",
        '/qn',
        'ALLUSER=1',
        'OPTIONS="noAutoStart=true"'
    )
    Start-Process -FilePath "msiexec.exe " -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

if(Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run")
{
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -ErrorAction SilentlyContinue -Force
} 

if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
{
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -ErrorAction SilentlyContinue -Force
} 

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
