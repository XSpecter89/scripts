# This script ingests a CSV of host names, disables those computer objects, then moves them to a disabled objects OU

$hostList = Get-Content -Path '' # Enter path to CSV containing host names to be disabled
$disabledObjectOU = '' # Enter DN of the disabled objects OU

foreach ($hostName in $hostList) {
    $hostToDisable = $hostName + '$'
    disable-adaccount -Identity $hostToDisable
    get-adcomputer -Identity $hostName | move-adobject -TargetPath $disabledObjectOU
}