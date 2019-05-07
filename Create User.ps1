<#
This script locates the 'New Users.xlsx' file on your non-admin account Desktop. Since PowerShell is ran as an administrator to work in AD,
your current username will include the '-admin' suffix. This script parses the current username for everything before the '-', and if a '-'
doesn't exist, it takes the full username. The 'New Users.xlsx' file must have first name in the first column, last name in the second column,
and password in the third column. It takes the data from the Excel file, parses it appropriately, and creates the AD user object. If the user
already exists in AD, it will display a warning and skip to the next user in the list. Skipped users will need to be created manually.

Edit lines 31 and 40 appropriately before running the script.
#>

# Set non-admin username to find appropriate Desktop where 'New Users.xlsx' is stored
$username = if ($env:USERNAME.IndexOf('-') -eq -1) { $env:USERNAME } else { $env:USERNAME.Substring(0,$env:USERNAME.IndexOf('-')) }

# Create new Excel object and opens the 'New Users.xlsx' workbook and first worksheet
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workBook = $excel.Workbooks.Open('C:\users\' + $username + '\desktop\New Users.xlsx')
$workSheet = $workBook.Sheets.Item(1)

# Find the number of records in the worksheet
$workSheetRange = $workSheet.UsedRange
$rowCount = $workSheetRange.Rows.Count

# Create arrays to store user data from Excel and DN path variable
$arrGivenNames = @()
$arrSurnames = @()
$arrPasswords = @()
$arrNames = @()
$arrSamAccountNames = @()
$arrUserPrincipalNames = @()
$path = "" # Enter the DN where you want the user objects to be created in AD

# Loop through each record, assigning data to appropriate arrays
For ($i = 1; $i -le $rowCount; $i++) {
    $arrGivenNames += $workSheet.Cells.Item($i, 1).Value()
    $arrSurnames += $workSheet.Cells.Item($i, 2).Value()
    $arrPasswords += $worksheet.Cells.Item($i, 3).Value()
    $arrNames += $arrGivenNames[$i-1] + ' ' + $arrSurnames[$i-1]
    $arrSamAccountNames += $arrGivenNames[$i-1].Substring(0, 1) + ($arrSurnames[$i-1] -replace '\s+', '')
    $arrUserPrincipalNames += $arrSamAccountNames[$i-1] + '' # Enter the suffix for the principal name (e.g. @domain.com)
}

# Clean up open processes and disposes of open Excel objects
$workBook.Close()
$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheetRange)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[GC]::Collect()

# Check if user already exists in AD and, if so, display warning, else create the new user and display confirmation of user object creation
For ($i = 1; $i -le $rowCount; $i++) {

    $name = $arrNames[$i-1]
    $givenName = $arrGivenNames[$i-1]
    $surname = $arrSurnames[$i-1]
    $samAccountName = $arrSamAccountNames[$i-1]
    $userPrincipalName = $arrUserPrincipalNames[$i-1]
    $password = $arrPasswords[$i-1]

    If (Get-ADUser -Filter {SamAccountName -eq $samAccountName}) {
        Write-Warning "The user $samAccountName already exists."
    }
    Else {
        New-ADUser `
            -Name $name `
            -DisplayName $name `
            -GivenName $givenName `
            -Surname $surname `
            -SamAccountName $samAccountName `
            -UserPrincipalName $userPrincipalName `
            -StreetAddress '229 Avenida Fabricante,' `
            -City 'San Clemente' `
            -State 'CA' `
            -PostalCode '92672' `
            -Path $path `
            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
            -Enabled $true `
            -ChangePasswordAtLogon $false
        Write-Host "User $samAccountName has been created in $path."
    }
}