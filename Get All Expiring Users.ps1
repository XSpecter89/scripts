# Queries AD users for accounts expiring with the specified number of days
$daysToExpire = 90

Get-ADUser -Filter * -Properties * `
| Where {($_.accountexpirationdate -ne $null) -and ($_.accountexpirationdate -le ([DateTime]::Now).AddDays($daysToExpire)) -and ($_.enabled -eq 'true')} `
| Select name, distinguishedName, accountExpirationDate, @{n='manager';e={(Get-ADUser $_.manager).name}} `
| Export-CSV -Path "\\glk-fs01\shared\it\Scripts\Expired Users.csv" -NoTypeInformation