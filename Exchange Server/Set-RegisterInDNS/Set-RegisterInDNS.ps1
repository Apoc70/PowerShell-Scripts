# get network adapters with status UP (1) 
$adapters = Get-NetAdapter | Where-Object {
    $_.InterfaceOperationalStatus -eq 1  
} 

# sort adapters and select primary adapter (lowest metrik)
$primaryAdapter = $adapters | Sort-Object InterfaceIndex | Select-Object -First 1

#  get DNS client settings
$dnsClient = Get-DnsClient | Where-Object { $_.Interfaceindex -eq $primaryAdapter.Interfaceindex }

if ($dnsClient.RegisterThisConnectionsAddress -eq $false) {
    Write-Host ('Setting RegisterThisConnectionsAddress to TRUE for {0}' -f $dnsClient.InterfaceAlias)
}
else {
    Write-Host 'Nothing to change.'
}

# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
# RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
