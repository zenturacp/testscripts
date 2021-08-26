$Ngens = Get-ChildItem -Recurse -Filter "ngen.exe" -Path "$Env:SystemRoot\Microsoft.NET"
foreach ($Ngen in $Ngens) {
    Start-Process -FilePath $Ngen.VersionInfo.FileName -ArgumentList "executeQueuedItems" -Wait -NoNewWindow
}
