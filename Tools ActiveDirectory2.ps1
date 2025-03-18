Clear-Host
Write-Host "1. OBIEKTY NIEZALOGOWANE DO DOMENY X DNI"
Write-Host "2. ZNAJDŹ OBIEKTY BEZ MENADŻERA"
Write-Host "3. ZNAJDŹ OBIEKTY UŻYTKOWNIKA"
Write-Host "4. ZNAJDŹ OBIEKTY WYŁĄCZONE "
Write-Host "5. WYSZUKAJ"
Write-Host "6. SZUKAJ WYŁĄCZONYCH UŻYTKOWNIKÓW"
Write-Host "7. SZUKAJ UŻYTKOWNIKA PO TELEFONIE"
Write-Host "8. SZUKAJ UŻYTKOWNIKA BEZ MANAGERA W DOMENIE EN..TE"
Write-Host "9. SZUKAJ UŻYTKOWNIKA BEZ MANAGERA W DOMENIE WA..SKO"

$case = $(Write-Host "Wybierz numer zadania: " -ForegroundColor Cyan -NoNewLine; Read-Host);

$Credentials = Get-StoredCredential -Target wa
$Credentials_Ente = Get-StoredCredential -Target ente
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wa"
$Server_Ente = "en"
$GLOBAL:Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password

switch ($case) {

    '1' {
    do {
    Write-Host -ForegroundColor Red "UWAGA ! Skrypt wyszukuje, przenosi i wyłącza obiekty od wskazanego dnia wstecz. Minimalnie 40 Dni !"
    Write-Host ""
    $days = Read-Host "Podaj liczbę dni (empty=all)"
    $Date = (Get-Date).AddDays(-$days)
    Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Properties LastLogonDate -Credential $Credentials -Server $Server -SearchBase "l" | Select-Object Name, LastLogonDate | Sort-Object LastLogonDate

    $countobject = (Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Properties LastLogonDate -Credential $Credentials -Server $Server -SearchBase "1" | Select-Object LastLogonDate).count

    Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Properties DNSHostName, LastLogonDate, whenCreated, OperatingSystem, OperatingSystemVersion, DistinguishedName -Credential $Credentials -Server $Server -SearchBase "ou=wasko,dc=wasko,dc=pl" | Select-Object DNSHostName, LastLogonDate, whenCreated, OperatingSystem, OperatingSystemVersion, DistinguishedName | Sort-Object LastLogonDate | Out-File C:\Lista_obiektów-$countobject.txt

    Write-Output ""

    Write-Host "Znaleziono obiektów: $countobject"
    Write-Host -ForegroundColor Yellow "Obiekty zostały zapisane: C:\L.txt"
    Invoke-Item C:\Lista_obiektów-$countobject.txt

    $next1 = Read-Host "Czy przenieść obiekty do OU=wylaczone komputery ? Tak [Tak]"
    if ($next1 -eq "tak") { Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Credential $Credentials -Server $Server -SearchBase "1" | Move-ADObject -TargetPath "1" }
    else { break }

    $next2 = Read-Host "Czy wyłączyć przeniesione obiekty ? Tak [Tak]"
    if ($next2 -eq "tak") { Get-ADComputer -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -Credential $Credentials -Server $Server -SearchBase "1" | Disable-ADAccount }
    else { break }
    Write-Host -ForegroundColor DarkYellow "End file"
    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    } while ($back -eq "T")
    break
}

    '2' {
    cls
    do {
    cls
    $resultCount = (Get-ADComputer -LDAPFilter "(!managedby=*)" -Credential $Credentials -Server $Server -Properties name -SearchBase "1" | Measure-Object).count
    

    Get-ADComputer -ldapFilter "(!managedby=*)" -Credential $Credentials -Server $Server -Properties ManagedBy -SearchBase "1" | FT Name, DistinguishedName -Wrap
    Write-Host "Liczba znalezionych obiektów: $resultCount"
    Write-Host -ForegroundColor DarkYellow "End file"

    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    } while ($back -eq "T")
    break
}

    '3' {
    cls
    do {
    
    $next3 = Read-Host "Podaj login użytkownika"

    Get-ADComputer -Filter "ManagedBy -eq '$next3'" -Credential $Credentials -Server $Server -Properties ManagedBy | Format-Table Name, DistinguishedName -Wrap

    Write-Host -ForegroundColor DarkYellow "End file"
    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    } while ($back -eq "T")
    break
}

    '4' {
    cls
    do {
            $disabledObjects = Get-ADComputer -Filter { Enabled -eq $false } -Credential $Credentials -Server $Server -SearchBase "1" | Select-Object Name

            $disabledObjects | Format-List Name

            Write-Output "Liczba znalezionych obiektów: $($disabledObjects.Count)"

    $next4 = Read-Host "Czy przenieść obiekty do OU=wylaczone komputery? (Tak/Nie)"
            
            if ($next4 -eq "Tak") {
                Get-ADComputer -Filter { Enabled -eq $false } -Credential $Credentials -Server $Server -SearchBase "1" | Move-ADObject -TargetPath "1"
            }

            Write-Host -ForegroundColor DarkYellow "End file"
    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    } while ($back -eq "T")
    break
    }

    '5' {  
    cls
    do {
    $Credentials = Get-StoredCredential -Target en
    $Username = $Credentials.UserName
    $Password = $Credentials.Password
    $Server = "en"

    $Credentials = Get-StoredCredential -Target f
    $Username = $Credentials.UserName
    $Password = $Credentials.Password
    $Server = "f"

 (Get-ADUser -filter * -Properties SamAccountName, description, Company -Server 'wasko' | Where-Object description -like "*nieob*" | select Company, SamAccountName, description)
,(Get-ADUser -filter * -Properties SamAccountName, description, Company -Server 'ente' | Where-Object description -like "*nieob*" | select Company, SamAccountName, description)
,(Get-ADUser -filter * -Properties SamAccountName, description, Company -Server 'fk' | Where-Object description -like "*nieob*" | select Company, SamAccountName, description)
    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    } while ($back -eq "T")
    break
    }

    '6' {
     do {
            $disabledUsers = Get-ADUser -Filter { Enabled -eq $false } -Credential $Credentials -Server $Server -SearchBase "1" | Select-Object Name

            $disabledUsers | Format-List Name |  Out-File C:\lis.txt

            Write-Output "Liczba znalezionych wyłączonych kont: $($disabledUsers.Count)"

    $next5 = Read-Host "Czy przenieść konta do OU=wylaczone konta? (Tak/Nie)"
            
            if ($next5 -eq "Tak") {
                Get-ADUser -Filter { Enabled -eq $false } -Credential $Credentials -Server $Server -SearchBase "1" | Move-ADObject -TargetPath "1"
            }

            Write-Host -ForegroundColor DarkYellow "End file"
    $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"
    } while ($back -eq "T")
      break
    }

    '7' {
    do {

            $telNumber = Read-Host "Podaj fragment nr. telefonu"
            Get-AdUser -Filter * -Properties MobilePhone, HomePhone, OfficePhone, DisplayName -Credential $Credentials -Server $Server | Select-Object DisplayName, `
            @{Name = "MobilePhone";Expression = {($_.MobilePhone -replace '[^0-9]')}},`
            @{Name = "OfficePhone";Expression = {($_.OfficePhone -replace '[^0-9]')}},`
            @{Name = "HomePhone";Expression = {($_.HomePhone -replace '[^0-9]')}}|`
            Where-Object {($_.MobilePhone -like ("*$telNumber*")) `
            -or ($_.OfficePhone -like ("*$telNumber*"))`
            -or ($_.HomePhone -like ("*$telNumber*"))}

           Write-Host -ForegroundColor DarkYellow "End file"
           $back = Read-Host "Powrót do ekranu głównego ? T [Tak]"           
    } while ($back -eq "T")           
    break
    }
'8' {
    do {
        # Pobierz użytkowników bez menedżera
        $users = Get-ADUser -LDAPFilter "(!(manager=*))" -Properties Name -Credential $Credentials_Ente -Server $Server_Ente -SearchBase "1"
        
        # Wyświetl listę użytkowników
        if ($users) {
            foreach ($user in $users) {
                Write-Host "Nazwa użytkownika: $($user.Name)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "Brak użytkowników z pustym atrybutem 'manager'." -ForegroundColor Red
        }

        Write-Host -ForegroundColor DarkYellow "End file"
        $back = Read-Host "Powrót do ekranu głównego? T [Tak]"
    } while ($back -eq "T")
    break
}


'9' {
    do {
        # Pobierz użytkowników bez menedżera
        $users = Get-ADUser -LDAPFilter "(!(manager=*))" -Properties Name -Credential $Credentials -Server $Server -SearchBase "OU=wasko, DC=wasko, DC=pl"
        
        # Wyświetl listę użytkowników
        if ($users) {
            foreach ($user in $users) {
                Write-Host "Nazwa użytkownika: $($user.Name)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "Brak użytkowników z pustym atrybutem 'manager'." -ForegroundColor Red
        }

        Write-Host -ForegroundColor DarkYellow "End file"
        $back = Read-Host "Powrót do ekranu głównego? T [Tak]"
    } while ($back -eq "T")
    break
}

}
if ($back -eq "t") {C:\Tools.ps1}
