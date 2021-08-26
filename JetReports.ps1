$Url = 'https://bc365v834p.blob.core.windows.net/software/1003/JetReports/JetExcelAddin_20_5_20042_1.exe?st=2020-04-16T08%3A49%3A33Z&se=2030-04-17T08%3A49%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=PZTztJI4BiU%2BmwzUX2p9p7sk4uu8%2Fl4Ot1%2FmFqh%2BOjQ%3D'
$Product = 'Jet Reports'
$SetupFile = 'JetExcelAddin_20_5_20042_1.exe'

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
    $SetupArguments = @(
        "/i",
        "$InstallTempFolder\Jet.Setup.ExcelAddIn_x64.msi",
        "EXCELPLATFORM=x64",
        "/qn"
    )
    Start-Process -FilePath "msiexec.exe" -WorkingDirectory $InstallTempFolder -Argumentlist $SetupArguments -NoNewWindow -Wait
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
