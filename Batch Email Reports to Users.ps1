<# Specify the full path, including file name and extension, to the email list that specifies who should receive which report. #>
$DataFilePath = ''
<# Specify which folder contains the reports. #>
$ReportsFolderPath = ''

<# Specify the from account, subject, and body of the email. #>
$fromAccount = ''
$subjectLine = ''
$body = ''

<# Create an Excel COM object, open the email list, and count the number of records. #>
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $workBook = $excel.Workbooks.Open($DataFilePath)
}
catch {
    $_ | Out-File -FilePath "C:\ReportEmailErrors.txt" -Append
}
$workSheet = $workBook.Sheets.Item(1)
$workSheetRange = $workSheet.UsedRange
$rowCount = $workSheetRange.Rows.Count

<# Create arrays to store the email addresses and report file names defined by the email list. #>
$arrEmails = @()
$arrReports = @()

<# Populate the arrays with the information from the email list. #>
for ($i = 1; $i -le $rowcount; $i++) {
    $arrEmails += $workSheet.Cells.Item($i, 1).Value()
    $arrReports += $workSheet.Cells.Item($i, 2).Value()
}

<# Close the email list and Excel, then release the memory. #>
try {
    $workBook.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheetRange)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
}
catch {
    $_ | Out-File -FilePath "C:\ReportEmailErrors.txt" -Append
}

<# Send the emails to the individuals with their respective reports. #>
for ($i = 1; $i -le $rowCount; $i++) {
    try {
        Send-MailMessage `
            -From $fromAccount `
            -To $arrEmails[$i-1] `
            -Cc '' `
            -Subject ($arrReports[$i-1].substring(0,5) + $subjectLine) `
            -Body $body `
            -Attachments ($ReportsFolderPath + $arrReports[$i-1]) `
            -SmtpServer '' `
            -DeliveryNotificationOption OnFailure `
            -ErrorAction Stop
        $_ | Out-File -FilePath "C:\ReportEmailErrors.txt" -Append
        <# Write-Host (([string]$i) + (" | From: " + $fromAccount) + (" | To: " + $arrEmails[$i-1]) + (" | Subject: " + $subjectLine) + (" | Body: " + $body) + (" | Attachments: " + $arrReports[$i-1])) #><# unit test line #>
    }
    catch {
        $_ | Out-File -FilePath "C:\ReportEmailErrors.txt" -Append
    }
}