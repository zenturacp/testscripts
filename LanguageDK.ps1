$Url = 'https://bc365v834p.blob.core.windows.net/software/languagepack/win10/1909/da-DK.exe?st=2020-04-15T08%3A31%3A19Z&se=2030-04-16T08%3A31%3A00Z&sp=rl&sv=2018-03-28&sr=b&sig=Wre2iOdmFoOrLzvC9136%2FzckvARkt6oKm1Us8QRPUk4%3D'
$Product = 'Win10 DK Language Pack'
$SetupFile = 'da-DK.exe'
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

#####################
# Global script vars
#####################
$lp_root_folder = "$InstallTempFolder" #Root folder where the copied sourcefiles are
$architecture = "x64" #Architecture of cab files
$systemlocale = 'en-US' #System local when script finishes

#####################
# Start installation of language pack on Win10 1809 and higher
#####################
$installed_lp = New-Object System.Collections.ArrayList
foreach ($language in Get-ChildItem -Path "$lp_root_folder\LXP") {
    #check if files exist

    $appxfile = $lp_root_folder + "\LXP\" + $language.Name + "\LanguageExperiencePack." + $language.Name + ".Neutral.appx"
    $licensefile = $lp_root_folder + "\LXP\" + $language.Name + "\License.xml"
    $cabfile = $lp_root_folder + "\LangPacks\Microsoft-Windows-Client-Language-Pack_" + $architecture + "_" + $language.Name + ".cab"
    
    if (!(Test-Path $appxfile)) {
        Write-Host $language.Name " - File missing: $appxfile" -ForegroundColor Red
        Write-Host "Skipping installation of "  + $language.Name
    } elseif (!(Test-Path $licensefile)) {
        Write-Host $language.Name " - File missing: $licensefile" -ForegroundColor Red
        Write-Host "Skipping installation of "  + $language.Name
    } elseif (!(Test-Path $cabfile)) {
        Write-Host $language.Name " - File missing: $cabfile" -ForegroundColor Red
        Write-Host "Skipping installation of "  + $language.Name
    } else {
        Write-Host $language.Name " - Installing $cabfile" -ForegroundColor Green
        Start-Process -FilePath "dism.exe" -WorkingDirectory "C:\Windows\System32" -ArgumentList "/online /Add-Package /PackagePath=$cabfile /NoRestart" -Wait -NoNewWindow

        Write-Host $language.Name " - Installing $appxfile" -ForegroundColor Green
        Start-Process -FilePath "dism.exe" -WorkingDirectory "C:\Windows\System32" -ArgumentList "/online /Add-ProvisionedAppxPackage /PackagePath=$appxfile /LicensePath=$licensefile /NoRestart" -Wait -NoNewWindow

        Write-Host $language.Name " - CURRENT USER - Add language to preffered languages (User level)" -ForegroundColor Green
        $prefered_list = Get-WinUserLanguageList
        $prefered_list.Add($language.Name)
        Set-WinUserLanguageList($prefered_list) -Force 

        $installed_lp.Add($language.Name)
    }
}

Write-Host "$systemlocale - Setting the system locale" -ForegroundColor Green
Set-WinSystemLocale -SystemLocale $systemlocale

Set-Location $CurrentDir
Write-Host "$Product`: Performing CleanUp"
Remove-Item -Path $InstallTempFolder -Recurse -Force
