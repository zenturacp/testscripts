while ((Get-Service RdAgent).Status -ne 'Running') { Start-Sleep -s 5 }
while ((Get-Service WindowsAzureGuestAgent).Status -ne 'Running') { Start-Sleep -s 5 }

Get-Service "wuauserv" | Stop-Service
Get-Service "wuauserv" | Set-Service -StartupType Disabled

if( Test-Path $Env:SystemRoot\windows\system32\Sysprep\unattend.xml ){ Remove-Item $Env:SystemRoot\windows\system32\Sysprep\unattend.xml -Force}
& $env:SystemRoot\System32\Sysprep\Sysprep.exe /oobe /generalize /quiet /quit /mode:vm
while($true) { $imageState = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State | Select-Object ImageState; Write-Output $imageState.ImageState; if($imageState.ImageState -ne 'IMAGE_STATE_GENERALIZE_RESEAL_TO_OOBE') { Start-Sleep -s 10 } else { break } }
