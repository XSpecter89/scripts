# This script reads in all files with the specified folder, then prompts for the username of the user that should receive permissions to that file

$dataFolder = Get-ChildItem '' # Enter path to the folder on the network that contains the files
$permissions = 'ReadAndExecute' # Enter the permissions level the user should be granted
foreach ($dataFile in $dataFolder) {
    try {
        $username = Read-Host -Prompt "Input username (in the format 'domain\username') to receive permissions to $DataFile (press Enter to skip this file)"
        If($username -eq "") {
            continue
        } else {
            $filePath = $dataFile.FullName
            $acl = (Get-Item $filePath).GetAccessControl('Access')
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($username, $permissions, 'None', 'None', 'Allow')
            $acl.SetAccessRule($rule)
            Set-Acl -path $filePath -AclObject $acl
            Write-Host "$permissions permissions set on $dataFile for user $username"
        }
    } catch {
        Write-Host "Error setting permissions on $dataFile for username $username. Verify the username was entered correctly."
    }
}
Pause