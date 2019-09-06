# This script reads a CSV full of usernames and adds them all to the specified distribution group

$users = get-content -path '' # Enter full path to CSV containing usernames that should be added to the distribution group
foreach($un in $users) {
    add-distributiongroupmember -Identity '' -Member $un # Enter the name of the distribution group
}