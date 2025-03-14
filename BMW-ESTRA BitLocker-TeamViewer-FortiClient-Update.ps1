# Sets BitLocker key (Pin) and save backup file on USB-Drive via name Estra
# & install FortiClient, TeamViewer 
# for estra-automotive.com
# byWojas

Add-Type -AssemblyName System.Web

Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

Clear-Host

Function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length
    )
    $charSet = '0123456789'.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
 
    $rng.GetBytes($bytes)

    $result = New-Object char[]($length)

    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i]%$charSet.Length]
    }

    return (-join $result)
}


### Set PowerShell execution policies ###

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force




### BitLocker ###

$usbDriveEstra = Get-WmiObject -Class Win32_Volume | Where-Object { $_.DriveType -eq 2 -and $_.Label -eq 'Estra' } | Select-Object -ExpandProperty DriveLetter

$bitlockerYes = $(Write-Host "Do you want enable pin BITLOCKER ? (Y/N) " -ForegroundColor yellow -NoNewline -BackgroundColor red ; Read-Host)

if ($bitlockerYes -eq "y") {
    if ($usbDriveEstra -inotlike $null) {
        Write-host -ForegroundColor DarkRed -BackgroundColor White "USB Drive Connected Volume" $usbDriveEstra
        Start-Sleep -Seconds 5
        $hostname = $env:COMPUTERNAME
        $password = Get-RandomPassword 6
        $firstLine = (Manage-BDE -protectors -get C: | Select-String -Pattern "ID:" | Select-Object -First 1).Line
        $idBitlocker = ($firstLine -split ':', 2)[1].Trim()

        try {
            Manage-BDE -protectors -adbackup -id $idBitlocker
            Manage-BDE -protectors -get C: | Out-File "$usbDriveEstra\$hostname-$idBitlocker-$password.txt"
            Manage-BDE -protectors -add C: -TPMAndPIN $password

            Write-Host "Set BitLocker PIN: $($password) for computer $($hostname)"
            Write-Host "The file has been saved to USB $($usbDriveEstra)"
            [console]::beep(1200, 300)
            [console]::beep(1200, 300)
            [console]::beep(1200, 400)
            [console]::beep(700, 400)
        }  
        catch [System.Security.SecurityException] {
            Write-Warning "Run the file with administrator rights"
        } 
        catch {
            Write-Host "An error occurred: $_"
        }

        Write-host -ForegroundColor DarkRed -BackgroundColor White "Current PIN:" $password
        Pause
    } 
    else {
        Write-Host -ForegroundColor White -BackgroundColor Red "USB Drive 'Estra' Not Found."
        [console]::beep(700, 1500)
        Pause
    }
} 
else { }


### TeamViewer ###

    Write-Host ""
    $teamViewerYes = $(Write-Host "Do you want uninstall/install TEAMVIEWER 64-bit ? (Y/N) " -ForegroundColor yellow -NoNewline -BackgroundColor DarkBlue ; Read-Host)

    if ($teamViewerYes -eq "y") {

    $teamViewerInstalled = Test-Path 'C:\Program Files (x86)\TeamViewer\TeamViewer.exe'
    $teamViewerInstalled2 = Test-Path 'C:\Program Files\TeamViewer\TeamViewer.exe'

    Get-Process -Name TeamViewer | Stop-Process -Force

if ( ($teamViewerInstalled -eq $true) -or ($teamViewerInstalled2 -eq $true) ) {
    
    Write-Host "Uninstalls TeamViewer.."
    
    try {
    Start-Process -FilePath 'C:\Program Files (x86)\TeamViewer\uninstall.exe' -ArgumentList "/S" -Wait -NoNewWindow
    } catch { Write-Host "Not find uninstall.exe under program files (x86)" -ForegroundColor DarkRed }

    try {
    Start-Process -FilePath 'C:\Program Files\TeamViewer\uninstall.exe' -ArgumentList "/S" -Wait -NoNewWindow
    } catch { Write-Host "Not find uninstall.exe under program files" -ForegroundColor DarkRed }

    try {
    Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*TeamViewer*"} | ForEach-Object { $_.Uninstall() }
    } catch { Write-Host "Not find proccess like TeamViewer" -ForegroundColor DarkRed }
    
    Start-Sleep 1

    Write-Host "Install TeamViewer with additional settings.."
    
    $msiFile = "$($usbDriveEstra)\install\Teamviewer\TeamViewer_Host.msi"
    $customConfig = "$($usbDriveEstra)\install\Teamviewer\672qxne.zip"
    $settingsFile = "$($usbDriveEstra)\install\tvconfig.tvopt"
    
    Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$msiFile`" /qn SETTINGSFILE=`"$settingsFile`"" -Wait -NoNewWindow
    Start-Sleep -Seconds 1
    Start-Process -FilePath "C:\Program Files\TeamViewer\TeamViewer.exe" -ArgumentList "customize --remove --path `"$customConfig`" --restart-gui"
} 

else {
    
    $msiFile = "$($usbDriveEstra)\install\Teamviewer\TeamViewer_Host.msi"
    $customConfig = "$($usbDriveEstra)\install\Teamviewer\672qxne.zip"
    $settingsFile = "$($usbDriveEstra)\install\tvconfig.tvopt"
    
    Write-Host "Install TeamViewer with additional settings.."

    Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$msiFile`" /qn SETTINGSFILE=`"$settingsFile`"" -Wait -NoNewWindow
    Start-Sleep -Seconds 1
    Start-Process -FilePath "C:\Program Files\TeamViewer\TeamViewer.exe" -ArgumentList "customize --remove --path `"$customConfig`" --restart-gui"
} }



### FortiClient ###

Write-Host ""
$FortiYes = $(Write-Host "Check version FortiClient and install new one ? (Y/N) " -ForegroundColor yellow -NoNewline -BackgroundColor DarkBlue ; Read-Host)

if ($FortiYes -eq "y") {

$fortiClientInstalledVersion = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | 
    Where-Object {$_.DisplayName -eq "FortiClient"}).DisplayVersion

if ([version]$fortiClientInstalledVersion -lt [version]"7.2.0") {
    
    $fortiClientProcesses = Get-Process | Where-Object { $_.ProcessName -like "FortiClient*" }
    foreach ($process in $fortiClientProcesses) {
    Stop-Process -Id $process.Id -Force
}
    Start-Process -FilePath "$($usbDriveEstra)\install\FortiClientSetup_7.2.3_x64.exe" -ArgumentList "/quiet /norestart" -Wait -NoNewWindow
    Write-Host "FortiClient has been installed.."
} else {
    Write-Host "The current version of FortiClient ($fortiClient Installed Version) is new enough."
} }



### Install Windows Update ###

Write-Host ""
$updateYes = $(Write-Host "Do you want UPDATE the system ? (Y/N) " -ForegroundColor White -NoNewline -BackgroundColor red ; Read-Host)
if ($updateYes -eq "y") {


    Write-Host -ForegroundColor DarkRed -BackgroundColor Yellow "Enable Receive Updates for Other Microsoft Products"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name AUOptions -Value 3
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name AllowMUUpdateService -Value 1
    Write-Host -ForegroundColor Black -BackgroundColor White "Done"
    
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -ForceBootstrap -Confirm:$false -Force
    Install-Module -Name PSWindowsUpdate -SkipPublisherCheck -AllowClobber -Confirm:$false -Force
 
    Write-Host -ForegroundColor DarkRed -BackgroundColor Yellow "Installing optional Windows updates.."
    $optionalUpdates = Get-WindowsUpdate -Category Optional
    $optionalUpdates | foreach {
        $_ | Install-WindowsUpdate -AcceptAll
    }
    Write-Host -ForegroundColor Black -BackgroundColor White "Done"

    Write-Host -ForegroundColor DarkRed -BackgroundColor Yellow "Installing Windows updates.."
    
    Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -ForceDownload -ForceInstall
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -ForceDownload -ForceInstall -AutoReboot
    Write-Host -ForegroundColor Black -BackgroundColor White "Done"
    Pause

} else { Start-Sleep 2 }