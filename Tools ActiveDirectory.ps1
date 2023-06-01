CLS

Write-Host "1. OBIEKTY NIEZALOGOWANE DO DOMENY X DNI"
Write-Host "2. WYSZUKAJ OBIEKTY BEZ MENADŻERA"
Write-Host "3. ZNAJDŹ OBIEKTY UŻYTKOWNIKA"
Write-Host "4. WYSZUKAJ WSZYSTKIE OBIEKTY WYŁĄCZONE OU=WASKO I PRZENIEŚ DO OU=wylaczone komputery"

$zadanie = Read-Host "Numer zadania"

$Credentials = Get-StoredCredential -Target wasko.pl
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wasko.pl"
$Credentials = New-Object System.Management.Automation.PSCredential $Username, $Password


if ($zadanie -eq 1) {
            CLS
            Write-host -ForegroundColor Red "UWAGA! Skrypt wyszukuje, przenosi i wyłącza obiekty od wskazanego dnia wstecz. Minimalnie 40 dni!"

            Write-host ""
            $days = Read-Host "Podaj liczbę dni"

            $Date = (Get-Date).AddDays(-$days)

            $objects = Get-ADComputer -Filter { (Enabled -eq $true) -and (LastLogonDate -lt $Date) } -Properties LastLogonDate -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Select-Object Name, LastLogonDate | Sort-Object LastLogonDate

            $objects | Format-Table Name, LastLogonDate

            $objects | Select-Object DNSHostName, LastLogonDate | Sort-Object LastLogonDate | Out-File C:\Lista_obiektów.txt

            Write-Output ""
            Write-Output "Znaleziono obiektów: $($objects.Count)"
            Write-Host -ForegroundColor Yellow "Obiekty zostały zapisane: C:\Lista_obiektów.txt"
            Invoke-Item C:\Lista_obiektów.txt

    $next1 = Read-Host "Czy przenieść obiekty do OU=wylaczone komputery? (Tak/Nie):"
            if ($next1 -eq "Tak") {
                Get-ADComputer -Filter { (Enabled -eq $true) -and (LastLogonDate -lt $Date) } -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Move-ADObject -TargetPath "ou=wylaczone komputery,dc=wasko,dc=pl"
            }

    $next2 = Read-Host "Czy wyłączyć przeniesione obiekty? (Tak/Nie):"
            if ($next2 -eq "Tak") {
                Get-ADComputer -Filter { (Enabled -eq $true) -and (LastLogonDate -lt $Date) } -Credential $Credentials -Server $Server -SearchBase "ou=wylaczone komputery,dc=wasko,dc=pl" | Disable-ADAccount
            }

            Write-Host -ForegroundColor DarkYellow "End file"
            Wait-Debugger
}
elseif ($zadanie -eq 2) {
            Write-Output "Liczba znalezionych obiektów: $($(Get-ADComputer -LDAPFilter "(!managedby=*)" -Credential $Credentials -Server $Server -Properties Manager, Description -SearchBase "ou=wasko,dc=wasko,dc=pl" | Select-Object samAccountName, Description).Count)"

            Get-ADComputer -LDAPFilter "(!managedby=*)" -Properties Manager, Description -SearchBase "ou=wasko,dc=wasko,dc=pl" -Credential $Credentials -Server $Server | Select-Object samAccountName, Description

            Write-Host -ForegroundColor DarkYellow "End file"
            Wait-Debugger
}
elseif ($zadanie -eq 3) {
            $next3 = Read-Host "Podaj login użytkownika: "

            Get-ADComputer -Filter "ManagedBy -eq '$next3'" -Credential $Credentials -Server $Server -Properties ManagedBy | Format-Table Name, DistinguishedName -Wrap

            Write-Host -ForegroundColor DarkYellow "End file"
            Wait-Debugger
}
elseif ($zadanie -eq 4) {
            $disabledObjects = Get-ADComputer -Filter { Enabled -eq $false } -Credential $Credentials -Server $Server -SearchBase "dc=wasko,dc=pl" | Select-Object Name

            $disabledObjects | Format-Table Name

            Write-Output "Liczba znalezionych obiektów: $($disabledObjects.Count)"

    $next4 = Read-Host "Czy przenieść obiekty do OU=wylaczone komputery? (Tak/Nie):"
            if ($next4 -eq "Tak") {
                Get-ADComputer -Filter { Enabled -eq $false } -Credential $Credentials -Server $Server -SearchBase "dc=wasko,dc=pl" | Move-ADObject -TargetPath "ou=wylaczone komputery,dc=wasko,dc=pl"
            }

            Write-Host -ForegroundColor DarkYellow "End file"
            Wait-Debugger
}
