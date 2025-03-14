# Pobranie danych logowania z zaufanych poświadczeń
$Credentials = Get-StoredCredential -Target 'wasko.pl'
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wasko.pl"
$GLOBAL:Credentials = New-Object System.Management.Automation.PSCredential ($Username, $Password)

# Definiowanie lokalizacji OU, z której chcemy pobrać użytkowników
#$OU = "OU=FONON,OU=WSPOLPRACOWNICY,OU=WASKO,DC=wasko,DC=pl"
$OU = "OU=WASKO,DC=wasko,DC=pl"

# Definiowanie filtra - tutaj pobieramy wszystkich użytkowników
$filter = '*'

# Pobieranie użytkowników z wybranymi właściwościami
$ADUsers = Get-ADUser -Filter $filter -SearchBase $OU -Properties Enabled, DisplayName, UserPrincipalName, Mail -Credential $Credentials -Server $Server |
    Select-Object @{
        Name = 'AccountEnabled'; Expression = { $_.Enabled }
    },
    @{
        Name = 'SamAccountName'; Expression = { $_.SamAccountName }
    },
    @{
        Name = 'DisplayName'; Expression = { $_.DisplayName }
    },
    @{
        Name = 'ImmutableId'; Expression = { $_.ImmutableId }
    },
    @{
        Name = 'Mail'; Expression = { $_.Mail }
    }

# Wczytanie danych z pliku CSV, który nie ma nagłówków
# Definiowanie nagłówków kolumn w pliku
$Headers = @("SamAccountName", "Enabled")  # Zakładamy, że mamy tylko dwie kolumny: SamAccountName i Enabled
$CsvData = Import-Csv -Path "d:\fonon_email\mxcoig.csv" -Header $Headers

# Tworzymy listę wyników porównań
$ComparisonResults = @()

# Przeglądanie każdego wpisu w CSV i szukanie odpowiedniego użytkownika w AD
foreach ($CsvEntry in $CsvData) {
    # Szukanie użytkownika w AD o takim samym SamAccountName
    $MatchedADUser = $ADUsers | Where-Object { $_.SamAccountName -eq $CsvEntry.SamAccountName }

    if ($MatchedADUser) {
        # Jeśli znaleziono użytkownika w AD, dodajemy go do wyników
        $ComparisonResults += [PSCustomObject]@{
            CsvSamAccountName    = $CsvEntry.SamAccountName
            CsvEnabled           = $CsvEntry.Enabled
            ADUserSamAccountName = $MatchedADUser.SamAccountName
            ADUserDisplayName    = $MatchedADUser.DisplayName
            ADUserEnabled        = $MatchedADUser.AccountEnabled
            MatchFound           = $true
        }
    } else {
        # Jeśli nie znaleziono użytkownika w AD
        $ComparisonResults += [PSCustomObject]@{
            CsvSamAccountName    = $CsvEntry.SamAccountName
            CsvEnabled           = $CsvEntry.Enabled
            ADUserSamAccountName = $null
            ADUserDisplayName    = $null
            ADUserEnabled        = $null
            MatchFound           = $false
        }
    }
}

# Opcjonalnie: Zapisz wyniki do pliku CSV
$ComparisonResults | Export-Csv -Path "d:\fonon_email\Porownanie.csv" -NoTypeInformation

# Wyświetlenie wyników w formie tabeli, tylko pasujące wiersze
$ComparisonResults | Where-Object { $_.MatchFound -eq $true } | Format-Table CsvSamAccountName, CsvEnabled, ADUserSamAccountName, ADUserDisplayName, ADUserEnabled -AutoSize
