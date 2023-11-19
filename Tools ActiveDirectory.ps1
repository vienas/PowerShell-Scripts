Clear-Host
Write-Host "1. OBIEKTY NIEZALOGOWANE DO DOMENY X DNI"
Write-Host "2. WYSZUKAJ OBIEKTY BEZ MENADŻERA ou=wasko,dc=wasko,dc=pl"
Write-Host "3. ZNAJDŹ OBIEKTY UŻYTKOWNIKA"
Write-Host "4. WYSZUKAJ OBIEKTY WYŁĄCZONE OU=WASKO I PRZENIEŚ DO OU=wylaczone komputery"
Write-Host "5. L4 - WYSZUKAJ OSOBY NA ZWOLNIENIU LEKARSKIM"

$case = $(Write-Host "Wybierz numer zadania: " -ForegroundColor Cyan -NoNewLine; Read-Host);

$Credentials = Get-StoredCredential -Target wasko.pl
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wasko.pl"
$GLOBAL:Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password

switch ( $case ) {

    '1' {
    do {
    Write-Host -ForegroundColor Red "UWAGA ! Skrypt wyszukuje, przenosi i wyłącza obiekty od wskazanego dnia wstecz. Minimalnie 40 Dni !"
    Write-Host ""
    $days = Read-Host "Podaj liczbę dni (empty=all)"
    $Date = (Get-Date).AddDays(-$days)
    Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Properties LastLogonDate -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Select-Object Name, LastLogonDate | Sort-Object LastLogonDate

    $countobject = (Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Properties LastLogonDate -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Select-Object LastLogonDate).count

    Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Properties DNSHostName, LastLogonDate, whenCreated, OperatingSystem, OperatingSystemVersion, DistinguishedName -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Select-Object DNSHostName, LastLogonDate, whenCreated, OperatingSystem, OperatingSystemVersion, DistinguishedName | Sort-Object LastLogonDate | Out-File C:\Lista_obiektów-$countobject.txt

    Write-Output ""

    Write-Host "Znaleziono obiektów: $countobject"
    Write-Host -ForegroundColor Yellow "Obiekty zostały zapisane: C:\Lista_obiektów-$countobject.txt"
    Invoke-Item C:\Lista_obiektów-$countobject.txt

    $next1 = Read-Host "Czy przenieść obiekty do OU=wylaczone komputery ? Tak [Tak]"
    if ($next1 -eq "tak") { Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Move-ADObject -TargetPath "ou=wylaczone komputery, dc=wasko, dc=pl" }
    else { break }

    $next2 = Read-Host "Czy wyłączyć przeniesione obiekty ? Tak [Tak]"
    if ($next2 -eq "tak") { Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Credential $Credentials -Server $Server -SearchBase "ou=wylaczone komputery, dc=wasko, dc=pl" | Disable-ADAccount }
    else { break }
    Write-Host -ForegroundColor DarkYellow "End file"
    } while ($back = Read-Host "Powrót do ekranu głównego ? T [Tak]")
    break
}

    '2' {
    Clear-Host
    $resultCount = (Get-ADComputer -LDAPFilter "(!managedby=*)" -Credential $Credentials -Server $Server -Properties name -SearchBase "ou=wasko,dc=wasko,dc=pl" | Measure-Object).count
    

    Get-ADComputer -ldapFilter "(!managedby=*)" -Credential $Credentials -Server $Server -Properties ManagedBy -SearchBase "ou=wasko,dc=wasko,dc=pl" | FT Name, DistinguishedName -Wrap
    Write-Host "Liczba znalezionych obiektów: $resultCount"
    #Select-Object name, CanonicalName, whenCreated
    Write-Host -ForegroundColor DarkYellow "End file"

    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    break
}

    '3' {
    Clear-Host
    $next3 = Read-Host "Podaj login użytkownika"

    Get-ADComputer -Filter "ManagedBy -eq '$next3'" -Credential $Credentials -Server $Server -Properties ManagedBy | FT Name, DistinguishedName -Wrap
    Write-Host -ForegroundColor DarkYellow "End file"
    
    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    break
}

    '4' {
    Get-ADComputer -Filter {enabled -eq $false} -Properties Name -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl"

    $obiekty = Get-ADComputer -Filter {enabled -eq $false} -Properties Name, DNSHostName -SearchBase "ou=wasko,dc=wasko,dc=pl" -Credential $Credentials -Server $Server | Select-Object DNSHostName

    $liczobiekt2 = (Get-ADComputer -Filter {enabled -eq $false} -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Measure-Object).count

    Write-Output "Liczba znalezionych obiektów: $liczobiekt2"
    if ($liczobiekt2 -eq 0) {
    $back = Read-Host "Powrót do ekranu głównego? T [Tak]"
    break
    } 
    else {
    $next4 = Read-Host "Czy przenieść obiekty do OU=wylaczone komputery ? Tak [Tak]"
    if ($next4 -eq "tak") { Get-ADComputer -Filter {enabled -eq $false} -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Move-ADObject -TargetPath "ou=wylaczone komputery,dc=wasko,dc=pl" }

    Write-Host -ForegroundColor DarkYellow "End file"

    $back = Read-Host "Powrót do ekranu głównego? T [Tak]"
    break
    }
    }

    '5' {  
    Clear-Host
    $Credentials = Get-StoredCredential -Target ente.local
    $Username = $Credentials.UserName
    $Password = $Credentials.Password
    $Server = "ente.local"
    
    Get-ADUser -filter * -Properties SamAccountName, description, Company -Server 'wasko.pl' | Where-Object description -like "*pow. 14 dni*" | select Company, SamAccountName, description;
    Get-ADUser -filter * -Properties SamAccountName, description, Company -Server 'ente.local' | Where-Object description -like "*pow. 14 dni*" | select Company, SamAccountName, description;
    
    Write-Host -ForegroundColor DarkYellow "End file"

    $back = Read-Host "Powrót do ekranu głównego? T [Tak]"
    break
    }

}
if ($back -eq "t") {C:\Tools_ActiveDirectory.ps1}
