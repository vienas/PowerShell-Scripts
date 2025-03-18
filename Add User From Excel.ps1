Remove-Item -Path "D:\new_users.txt"

$dane = Import-excel D:\users.xlsx

$Credentials = Get-StoredCredential -Target ente
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "ente"
$Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password

function ADandMail {

$SamAccountUser = ($Firstname[0] + '.' + $Lastname).ToLower()

$hash = @{'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ó'='o'; 'ś'='s'; 'ń'='n'; 'ż'='z'; 'ź'='z'}
Foreach ($key in $hash.keys) {
    $SamAccountUser = $SamAccountUser.Replace($key, $hash.$key)
  }
#=========================================================
$Mainpath = "ente"
#=========================================================
$Pass = "password"
#=========================================================

# Create User in Active Directory

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
            -Title "User" `
            -Manager $Manager `
            -ChangePasswordAtLogon $true `
            -StreetAddress "Adress" `
            -City "Gliwice" `
            -Company "WASKO S.A." `
            -Country "pl" `
            -PostalCode "44-100" `
            -AccountExpirationDate "31-03-2022" `
            -Enabled $true `
            -AccountPassword (ConvertTo-SecureString $Pass -AsPlainText -Force)


    Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File  D:\new_users.txtt -append
    Write-Output "Użytkownik: Company\$SamAccountUser" | Out-File  D:\new_users.txt -append
    Write-Output "E-mail: $SamAccountUser@wasko.pl" | Out-File  D:\new_users.txt -append
    Write-Output "Hasło do zmiany: $Pass" | Out-File  D:\new_users.txt -append
    Write-Output "" | Out-File  D:\new_users.txt -append

#Create Email Account - Smartermail 

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
$mx = "$SamAccountUser@wasko"
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
