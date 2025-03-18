$Credentials = Get-StoredCredential -Target 'wa'
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wa"
$GLOBAL:Credentials = New-Object System.Management.Automation.PSCredential ($Username, $Password)

$OU = "WA"

$filter = '*'

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

$Headers = @("SamAccountName", "Enabled")  
$CsvData = Import-Csv -Path "d:\fo\mx.csv" -Header $Headers

$ComparisonResults = @()

foreach ($CsvEntry in $CsvData) {
    $MatchedADUser = $ADUsers | Where-Object { $_.SamAccountName -eq $CsvEntry.SamAccountName }

    if ($MatchedADUser) {
        $ComparisonResults += [PSCustomObject]@{
            CsvSamAccountName    = $CsvEntry.SamAccountName
            CsvEnabled           = $CsvEntry.Enabled
            ADUserSamAccountName = $MatchedADUser.SamAccountName
            ADUserDisplayName    = $MatchedADUser.DisplayName
            ADUserEnabled        = $MatchedADUser.AccountEnabled
            MatchFound           = $true
        }
    } else {
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

$ComparisonResults | Export-Csv -Path "d:\fo\Po.csv" -NoTypeInformation

$ComparisonResults | Where-Object { $_.MatchFound -eq $true } | Format-Table CsvSamAccountName, CsvEnabled, ADUserSamAccountName, ADUserDisplayName, ADUserEnabled -AutoSize
