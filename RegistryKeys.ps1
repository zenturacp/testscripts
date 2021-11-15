$MaxConnectionTimeInSeconds = 28800
$MaxDisconnectionTimeInSeconds = 28800
$MaxIdleTimeInSeconds = 14400
$EdgeAllowDomains = "*scanmetals.com,autologon.microsoftazuread-sso.com"

# Terminal Server
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "RemoteAppLogoffTimeLimit" -Type DWord -Value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "fResetBroken" -Type DWord -Value 1
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "MaxConnectionTime" -Type DWord -Value $($MaxConnectionTimeInSeconds*1000)
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "RemoteAppLogoffTimeLimit" -Type DWord -Value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "MaxDisconnectionTime" -Type DWord -Value $($MaxDisconnectionTimeInSeconds*1000)
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "MaxIdleTime" -Type DWord -Value $($MaxIdleTimeInSeconds*1000)
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "fEnableTimeZoneRedirection" -Value 1 -Type DWord

# Hide A,B,C,D,E Drives
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoDrives" -Value 31 -Type DWord

# Block Workplace Join
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin" -Name "BlockAADWorkplaceJoin" -Value 1 -Type DWord
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin" -Name "autoWorkplaceJoin" -Value 0 -Type DWord

# Windows 10
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" -Name "HideRecentlyAddedApps" -Value 1 -Type DWord

# Default Profile
$keyName = Get-Random -Minimum 100000 -Maximum 999999
reg load "HKU\$keyName" "C:\Users\Default\NTUSER.DAT"
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS

# Edge Policy
New-Item -Path "HKU:\$keyName\Software\Policies\Microsoft\Edge" -Force | Out-Null
Set-ItemProperty -Path "HKU:\$keyName\Software\Policies\Microsoft\Edge" -Name "AuthNegotiateDelegateAllowlist" -Type String -Value $EdgeAllowDomains
Set-ItemProperty -Path "HKU:\$keyName\Software\Policies\Microsoft\Edge" -Name "AuthServerAllowlist" -Type String -Value $EdgeAllowDomains

# ZoneMaps
New-Item -Path "HKU:\$keyName\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -Force | Out-Null
Set-ItemProperty -Path "HKU:\$keyName\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -Name "https" -Value 1 -Type DWord

New-Item -Path "HKU:\$keyName\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1" -Force | Out-Null
Set-ItemProperty -Path "HKU:\$keyName\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1" -Name "2103" -Value 0 -Type DWord

# Unload default profile again
Start-Sleep -Seconds 10
reg unload "HKU\$keyName"
Start-Sleep -Seconds 10
Remove-PSDrive -Name "HKU"
