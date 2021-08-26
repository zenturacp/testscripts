$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/Printix/CLIENT_%7Bscanmetals.printix.net%7D_%7B75fa1191-9a07-4eef-acf0-0896d728affb%7D.MSI?st=2020-04-14T08%3A19%3A07Z&se=2030-04-15T08%3A19%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=JJRGMVcR80XSKY%2BrewHKNEH4rLnEo3KSPjh%2FUjZWUD4%3D'
$Product = 'Printix'
$SetupFile = 'CLIENT_{scanmetals.printix.net}_{75fa1191-9a07-4eef-acf0-0896d728affb}.MSI'
$SetupArguments = @(
    "/i",
    $SetupFile,
    "/quiet",
    "WRAPPED_ARGUMENTS=/id:75fa1191-9a07-4eef-acf0-0896d728affb:oms"
)

$InstallTempFolder = "C:\Install-$(Get-Random -Minimum 10000 -Maximum 99999)"
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

New-Item -Path "HKLM:\SOFTWARE\printix.net\Printix Client" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\printix.net\Printix Client" -Name 'StartAsVDI' -Value 1 -Type DWord

Set-Location $InstallTempFolder
Write-Host "Starting $Product installer"

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    Start-Process -FilePath "msiexec.exe" -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
