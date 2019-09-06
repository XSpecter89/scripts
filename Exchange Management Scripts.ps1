# Collection of scripts for managing an Exchange server remotely through Powershell. Lines 4, 7, 10, and 31 must stay uncommented. Uncomment the Invoke-Command lines you want to run, as needed.

# Connection URI to your Exchange server PowerShell (e.g. http://server.domain.com/Powershell/)
$uri = ''

# Get admin credential to Exchange server
$cred = Get-Credential

# Create the Exchange shell session
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred

# Test command that should just return the name of the Exchange server, confirming you're connected to Exchange Management Shell
#Invoke-Command -Session $session { Get-ExchangeServer | Select-Object Name }

# Create new mail for specified user
#Invoke-Command -Session $session { Enable-Mailbox -Identity '' -Alias '' -Database '' }

# Modify primary SMTP record of mailbox
#Invoke-Command -Session $session { Set-Mailbox '' -EmailAddressPolicyEnabled $false -EmailAddresses '' }

# Mailbox move
#Invoke-Command -Session $session { '' | New-MoveRequest -TargetDatabase '' }

# Mailbox database populations
#Invoke-Command -Session $session { Get-Mailbox -WarningAction silentlycontinue } | Group-Object -Property Database | Select-Object Name,Count | Sort-Object Name | FT -Auto

# Enter a custom Exchange shell command between the braces { }
#Invoke-Command -Session $session {  }

#Clear Powershell sessions
Get-PSSession | Remove-PSSession