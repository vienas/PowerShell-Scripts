function Set-CredentialMx {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$Target
    )
    
    Process {
    $GLOBAL:Cred = Get-StoredCredential -Target $Target
    $GLOBAL:password = $Cred.GetNetworkCredential().Password
    $GLOBAL:userNameMx = $Cred.GetNetworkCredential().UserName
    }
}

function Get-SearchAliases {

    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$mxAccount,
        [Parameter(Mandatory=$true)]
        [string]$box

    )
    
    Process {

        $authBody = @{
            username = $userNameMx
            password = $password
        }
        
        $response = Invoke-RestMethod -Uri https://mx.coig.pl/api/v1/auth/authenticate-user -Method Post -Body $authBody
        
        $header = @{
            'Authorization' = "Bearer " + $response.accessToken
        }

        $webRequest = Invoke-RestMethod -uri https://mx.coig.pl/api/v1/settings/domain/aliases/$mxAccount -Method Get -Headers $header -body $user -ContentType "application/json"
        $viewAliasesPerUser = ($webRequest | Select-Object -Expand gridinfo).name
        
        foreach ($findAlias in $viewAliasesPerUser) {

        if ( $viewAliasesPerUser -notlike $null) {
        $tableResult1 =@()
        
        Write-Host -ForegroundColor Green ==  $mxAccount - Aliasy # ; $viewAliasesPerUser
        
        $tableResult1 += [PSCustomObject]@{
        Alias = $findAlias
        Adres_email = $mxAccount
        SmarterMailMX = $box }
                
        $tableResult1
        }}
        
    }
}

function Get-SearchMailingList {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$mxAccount,
        [Parameter(Mandatory=$true)]
        [string]$box
    )
    
    Process {
        

        $authBody = @{
            username = $userNameMx
            password = $password
        }
        
        $response = Invoke-RestMethod -Uri https://mx.coig.pl/api/v1/auth/authenticate-user -Method Post -Body $authBody
        
        $header = @{
            'Authorization' = "Bearer " + $response.accessToken
        }

        try { $webRequest2 = Invoke-RestMethod -uri https://mx.coig.pl/api/v1/settings/domain/mailing-lists/subscribers/$mxAccount/$requestMailListId -Method Get -Headers $header -body $user -ContentType "application/json"
            }
        catch { Write-Information "Brak użytkownika w Listach Mailigowych "}
        
        $webRequest2Id = $webRequest2 | Select-Object -ExpandProperty subscribedLists
        $IncludeUserIdML = $webRequest2 | Select-Object -ExpandProperty subscribedLists
        

        $webRequest1 = Invoke-RestMethod -uri https://mx.coig.pl/api/v1/settings/domain/mailing-lists/list -Method Get -Headers $header -body $user -ContentType "application/json"
        $ApiMailList = $webRequest1 | Select-Object -ExpandProperty items
        $ApiMailListId = ($webRequest1 | Select-Object -Expand items).id
        
       
        $dataMailList = @()
        

        if ( $IncludeUserIdML -cmatch $findid.id ) { Write-Host -ForegroundColor Red == $mxAccount - Listy Mailingowe }
       
        foreach ( $info in $ApiMailList ) {
        
        $dataMailList += [PSCustomObject]@{
        Id = $info.id
        listValue = $info.listAddress }
        }
        
        foreach ($findId in $dataMailList) {
        $tableResult = @()
        
        if ($IncludeUserIdML -cmatch $findid.id) {
        
        $tableResult += [PSCustomObject] @{
        Lista_mailingowa = $findId.listvalue
        Adres_email = $mxAccount
        SmarterMailMX = $box }
        $tableResult

        } 
        
        }
    } 
}

function Get-Final {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$false)]
        [string]$mxAccount
    )
    
    Process {

$Global:Data = Get-Content -Path 'C:\Users\itmw\Documents\disableMailwasko.txt'

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@



foreach ($mxAccount in $Data) {

#Write-Host "WASKO" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxwasko"
#Get-SearchAliases -mxAccount $mxAccount -box "wasko"
Get-SearchMailingList -mxAccount $mxAccount  -box "wasko"

#Write-Host "FONON" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxfonon"
#Get-SearchAliases -mxAccount $mxAccount -box "fonon"
Get-SearchMailingList -mxAccount $mxAccount -box "fonon"

#Write-Host "ENTE" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxente"
#Get-SearchAliases -mxAccount $mxAccount -box "ente"
Get-SearchMailingList -mxAccount $mxAccount -box "ente"

#Write-Host "COIG" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxcoig"
#Get-SearchAliases -mxAccount $mxAccount -box "coig"
Get-SearchMailingList -mxAccount $mxAccount  -box "coig"

#Write-Host "WASKO4B" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxwasko4b"
#Get-SearchAliases -mxAccount $mxAccount -box "w4b"
Get-SearchMailingList -mxAccount $mxAccount -box "w4b"

#Write-Host "GABOS" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxgabos"
#Get-SearchAliases -mxAccount $mxAccount -box "gabos"
Get-SearchMailingList -mxAccount $mxAccount -box "gabos"

#Write-Host "DE2ES" -ForegroundColor DarkYellow
Set-CredentialMx -Target "mxde2es"
#Get-SearchAliases -mxAccount $mxAccount -box "d2s"
Get-SearchMailingList -mxAccount $mxAccount -box "d2s"

  }
}
}


#Get-Final | convertto-html -Head $Header -Property Adres_email, Alias, SmarterMailMX | Out-File -FilePath c:\Aliasy.html -Append
Get-Final | convertto-html -Head $Header -Property Adres_email, Lista_mailingowa, SmarterMailMX -Title "List Mailingowa kont wylaczonych" -Body get-date | Out-File -FilePath c:\Listy_mailingowe.html -Append


#Export-Csv -Path 'C:\Users\itmw\Documents\MailAlias.csv' -InputObject * 
#$Global:mxAccount = $(Write-Host "USUNĄĆ WSZYSKIE ALIASY ?" -ForegroundColor yellow -NoNewline -BackgroundColor red ; Read-Host)