$Url = 'https://bc365v834p.blob.core.windows.net/software/fod/dotnet/1909/dotnet35.exe?st=2020-04-15T10%3A48%3A49Z&se=2030-04-16T10%3A48%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=LHSKpPZiUQlpUVlIvBwp17Fm7%2Bz1YTU0XNBrRe7Fiwk%3D'
$Product = 'DotNet 3.5'
$SetupFile = 'dotnet35.exe'
$SetupArguments = @(
    "/online",
    "/Add-Package",
    "/NoRestart"
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

Write-Host "$Product`: Extracting files from $SetupFile to $InstallTempFolder"
Start-Process -FilePath "$InstallTempFolder\$SetupFile" -Argumentlist "-o$InstallTempFolder -y" -NoNewWindow -Wait

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    $cabFiles = (Get-ChildItem -Path $InstallTempFolder -Filter "*.cab").VersionInfo.FileName
    Write-Host "$Product`: Installing packages`r`n$cabFiles"
    foreach ($cabFile in $cabFiles)
    {
        $SetupArguments += "/PackagePath:$cabFile" 
    }
    Start-Process -FilePath "dism.exe" -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
