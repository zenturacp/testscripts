$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/sqlnclient/x64/sqlncli.msi?st=2020-05-05T11%3A17%3A22Z&se=2030-05-06T11%3A17%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=wD24IpIkkRPxWh4JZhfIGgTuMrgfbBqdu0zp8WV1yZ0%3D'
$Product = 'SQL Native Client'
$SetupFile = 'sqlncli.msi'

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

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    $SetupArguments = @(
        "/i",
        "$InstallTempFolder\$SetupFile",
        "IACCEPTSQLNCLILICENSETERMS=YES",
        "/qn"
    )
    Start-Process -FilePath "msiexec.exe" -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
