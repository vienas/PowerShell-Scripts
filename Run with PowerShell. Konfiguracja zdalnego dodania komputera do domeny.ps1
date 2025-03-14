Start-Process powershell.exe -Verb RunAs {
Set-Service -Name dot3svc -Status Running -StartupType Automatic
Ipconfig /flushdns
Ipconfig /release
Ipconfig /renew
New-LocalUser -Name "Wasko" -Description "Wasko" -Password ( ConvertTo-SecureString "12Wasko3" -AsPlainText -Force)
Add-LocalGroupMember -Group "Administratorzy" -Member "Wasko"
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -name "fDenyTSConnections" -value 0
Set-NetFirewallProfile -Profile Domain, Public, Private -Enabled False
Shutdown /r
}