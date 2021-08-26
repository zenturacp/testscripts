# Run Windows Update
Write-Host "Searching for Updates"
$Updates = Start-WUScan -SearchCriteria "IsInstalled=0 AND IsHidden=0 AND IsAssigned=1" -Verbose
Write-Host "Installing Updates"
Install-WUUpdates -Updates $Updates -Verbose
Write-Host "Waiting 60 seconds"
Start-Sleep -Seconds 60