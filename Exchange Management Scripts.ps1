# Allow PowerShell to run Exchange Management Shell commands by connecting to the Exchange server
$UserCredential = Get-Credential
Enter-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri '' -Authentication Kerberos -Credential $UserCredential # Enter the connection URI for Powershell on the Exchange server (e.g. https://mail.domain.com/powershell)

# Create a mailbox for a specific user in the specified database
Enable-Mailbox -Identity '' -Alias '' -Database '' # Enter full DN (e.g. Domain.com/Users/John Doe) of the user object, the username/alias, and the database in which the mailbox should be created

# Move a specific user's mailbox to a new database
'' | New-MoveRequest -TargetDatabase '' # Enter full DN (e.g. Domain.com/Users/John Doe) of the user object and the destination database to which the mailbox should be moved

# Query all databases and display the number of mailboxes in each
Get-Mailbox | Group-Object -Property:Database | Select-Object Name,Count | Sort-Object Name | FT -Auto