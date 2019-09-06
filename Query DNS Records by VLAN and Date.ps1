<#
This script queries DNS A records on the specified VLANs that are greater than 8 months old
#>

$vlanlist = '' # Enter VLAN IP's to query separated by a comma (e.g. '192.168.1.*', '192.168.2.*', '192.168.5.*')

foreach($vlan in $vlanlist) {
    Get-DnsServerResourceRecord -ZoneName 'glaukosnet.local' -RRType 'A' | where-object {($_.Timestamp -lt [DateTime]::Now.AddMonths(-8)) -and ($_.RecordData.ipv4address -like $vlan)} # Enter ZoneName to query (e.g. domain.com)
}