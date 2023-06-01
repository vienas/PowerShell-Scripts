Remove-Item -Path "D:\Skrypty\New_Users.txt"

$dane = Import-excel D:\Skrypty\users.xlsx

$Credentials = Get-StoredCredential -Target ente.local
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wasko.pl"
$Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password



function ADandMail {

$SamAccountUser = ($Firstname[0] + '.' + $Lastname).ToLower()

$hash = @{'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ó'='o'; 'ś'='s'; 'ń'='n'; 'ż'='z'; 'ź'='z'}
Foreach ($key in $hash.keys) {
    $SamAccountUser = $SamAccountUser.Replace($key, $hash.$key)
  }
#=========================================================
$Mainpath = "OU=ente, DC=ente, DC=local"
#=========================================================
$Pass = "ChangePa`$`$w0rd"
#=========================================================

  New-ADUser `
            -SamAccountName $SamAccountUser `
            -UserPrincipalName $SamAccountUser@wasko.pl `
            -Server ente.local `
            -Credential $Credentials `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Displayname "$Firstname $Lastname" `
            -Description "$Department, $Firstname $Lastname" `
            -Department $Department `
            -Path "$mainpath" `
            -EmailAddress $SamAccountUser@wasko.pl `
            -Title "Praktykant" `
            -Manager $Manager `
            -ChangePasswordAtLogon $true `
            -StreetAddress "Berbeckiego 6" `
            -City "Gliwice" `
            -Company "WASKO S.A." `
            -Country "pl" `
            -PostalCode "44-100" `
            -AccountExpirationDate "31-07-2023" `
            -Enabled $true `
            -AccountPassword (ConvertTo-SecureString $Pass -AsPlainText -Force)


    Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File  D:\Skrypty\New_Users.txt -append
    Write-Output "Użytkownik: DOMENA-WASKO\$SamAccountUser" | Out-File  D:\Skrypty\New_Users.txt -append
    Write-Output "E-mail: $SamAccountUser@wasko.pl" | Out-File  D:\Skrypty\New_Users.txt -append
    Write-Output "Hasło do zmiany: $Pass" | Out-File  D:\Skrypty\New_Users.txt -append
    Write-Output "" | Out-File  D:\Skrypty\New_Users.txt -append

<#
$Cred = Get-StoredCredential -Target mxente
$password = $Cred.GetNetworkCredential().Password
$username = $Cred.GetNetworkCredential().UserName

$authBody = @{
username = $username
password = $password }

$response = Invoke-RestMethod -Uri https://mx.coig.pl/api/v1/auth/authenticate-user -Method Post -Body $authBody

$header = @{
'Authorization' = "Bearer " + $response.accessToken
'Content'='application/json' }
 
$SamAccount = "$SamAccountUser"
$Password = "$Pass"
$mx = "$SamAccountUser@owa.wasko.pl"
$FullName = "$Firstname $Lastname"

$user = @('{
"userData": {
		"userName": "' + $SamAccount + '",
        "fullName": "' + $FullName + '",
		"password": "' + $Password + '",

    "securityFlags": {
		"authType": 0,
		"authenticatingWindowsDomain": null,
		"isDomainAdmin": false,
        "isDisabled": false
    },
	 "isPasswordExpired": false
    },

"forwardList": {
     "forwardList":
       [
        "' + $mx + '",
       ],

     "keepRecipients": false,
     "deleteOnForward": true
    },

"userMailSettings": {
     "canReceiveMail": true,
     "enableMailForwarding": false
    },
}')

Invoke-RestMethod -uri https://mx.coig.pl/api/v1/settings/domain/user-put -Method Post -Headers $header -body $user -ContentType "application/json"

#>
}


foreach ($line in $dane) {

$Firstname = $line.imie
$Lastname = $line.nazwisko
$Department = $line.dzial
$Manager = $line.menadzer


    ADandMail


Write-Host -ForegroundColor Red "Dodałem użytkownika $Firstname $Lastname"
Write-Host ""

}