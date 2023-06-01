function Get-SendEmails {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SendDW,

        [Parameter(Mandatory=$false)]
        [string]$SendUDW
    )

    Process {
        $Data = Get-Content -Path 'C:\Users\itmw\Documents\data1.txt'

        foreach ($D in $Data) {
            $ComputerName = $($D.Split("|")[0])
            $EmailAddress = $($D.Split("|")[1])

            $Outlook = New-Object -ComObject Outlook.Application
            $Email = $Outlook.CreateItem(0)
            $Email.Recipients.Add($EmailAddress) | Out-Null
            $Recip = $Email.Recipients.Add($SendDW)
            $Recip.Type = 2

            if (![string]::IsNullOrEmpty($SendUDW)) {
                $Recip2 = $Email.Recipients.Add($SendUDW) 
                $Recip2.Type = 3
            }

            $Email.Subject = "Bitlocker na komputerze $ComputerName"
            $Email.Body = @"
Cześć,

Na Twoim komputerze $ComputerName jest wyłączony Bitlocker.

Prosimy o informację zwrotną po włączeniu funkcji Bitlocker (wszystkie partycje).

Komputer wymaga wpięcia do sieci firmowej.

W razie pytań prosimy o kontakt.

Administratorzy IT
"@

            $Email.Attachments.Add("C:\bitlocker.png") | Out-Null
            $Email.Attachments.Add("C:\bitlocker2.png") | Out-Null
            $Email.Attachments.Add("C:\bitlocker3.png") | Out-Null

            $Email.Send()

            Write-Host -ForegroundColor Cyan "Email was sent to $EmailAddress"
            Start-Sleep -Seconds 1
        }
    }

    End {
        Write-Host -ForegroundColor Red "End of file"
    }
}