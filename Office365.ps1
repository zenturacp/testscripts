$Url = 'https://bc365v834p.blob.core.windows.net/software/office365/ODTBusiness.exe?st=2020-05-06T11%3A30%3A48Z&se=2030-05-07T11%3A30%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=iqb5Qz6LMlnrA0D27IZetesVKmkCP1qna2huKTfFdHg%3D'
$Product = 'Office365 Business'
$SetupFile = 'ODTBusiness.exe'

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

Write-Host "Extracting files from $InstallTempFolder\$SetupFile"
Start-Process -FilePath $InstallTempFolder\$SetupFile -Argumentlist "-o$InstallTempFolder -y" -NoNewWindow -Wait

if (Test-Path "$InstallTempFolder\$SetupFile")
{
    Write-Host "$Product`: Downloading files for installer"
    Start-Process -FilePath "$InstallTempFolder\setup.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/download configuration-Office365-x64.xml" -NoNewWindow -Wait
    Write-Host "$Product`: Running installer"
    Start-Process -FilePath "$InstallTempFolder\setup.exe" -WorkingDirectory $InstallTempFolder -Argumentlist "/configure configuration-Office365-x64.xml" -NoNewWindow -Wait
}

Set-Location $CurrentDir

Remove-Item -Path $InstallTempFolder -Recurse -Force
