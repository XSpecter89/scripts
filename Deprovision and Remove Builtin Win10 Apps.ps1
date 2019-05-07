# This script deprovisions builtin Microsoft Office and and removes bloatware apps in Windows 10

Get-AppxProvisionedPackage –Online | where-object {$_.packagename –like “Microsoft.Office*”} | Remove-AppxProvisionedPackage -Online # Deprovisions Microsoft Office apps so they are not re-added to new user profiles
Get-AppxPackage -AllUsers | where-object {$_.name -like '*nordcurrent*' -OR $_.name -like '*king.com*' -OR $_.name -like '*delldigital*'} | Remove-AppxPackage -AllUsers # Removes bloatware apps such as Candy Crush and Dell Digital Delivery