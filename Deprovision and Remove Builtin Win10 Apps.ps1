# This script deprovisions builtin Microsoft Office and and removes bloatware apps in Windows 10


Get-AppxProvisionedPackage –Online | where-object {$_.packagename –like 'Microsoft.Office*' -OR $_.packagename -like 'Microsoft.windowscommunicationsapps*' -OR $_.packagename -like '*dell*' -AND $_.packagename -notlike '*command*'} | Remove-AppxProvisionedPackage -Online # Deprovisions Microsoft Office apps so they are not re-added to new user profiles
Get-AppxPackage -AllUsers | where-object {$_.name -like '*nordcurrent*' -OR $_.name -like '*king.com*' -OR $_.name -like 'Microsoft.windowscommunicationsapps*' -OR $_.name -like '*dell*' -AND $_.name -notlike '*command*'} | Remove-AppxPackage -AllUsers # Removes bloatware apps such as Candy Crush and Dell Digital Delivery