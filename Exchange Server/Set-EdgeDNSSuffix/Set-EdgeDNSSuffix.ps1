[cmdletbinding()]
param()

# DNS Suffix for Edge Transport Server
$DNSSuffix = "varunagroup.de"

# Fetch current DNS Suffix
$keyPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\'
$valueName = 'NV Domain'

try{
    # $oldDNSSuffix = Get-ItemProperty -Path $keyPath -Name $valueName -ErrorAction Stop
    $oldDNSSuffix = (Get-ItemProperty $keyPath -Name $valueName -ErrorAction Stop).$valueName

    if($oldDNSSuffix -eq '') {

        #Update primary DNS Suffix for FQDN
        Write-Verbose -Message ('Adding {0} to {1}' -f $DNSSuffix, $keyPath)
        Set-ItemProperty $keyPath -Name 'Domain' -Value $DNSSuffix
        Set-ItemProperty $keyPath -Name $valueName -Value $DNSSuffix

        # Update DNS Suffix Search List - Win8/2012 and above - if needed
        # Set-DnsClientGlobalSetting -SuffixSearchList $oldDNSSuffix,$DNSSuffix
    }
    else {
        Write-Host ("The computer DNS suffix is already set to {0}.`nThe script did not change the current configuration." -f $oldDNSSuffix)
    }

}
catch {
    $null = New-ItemProperty -Path $keyPath -Name $valueName -Value $DNSSuffix -Force
    $null = New-ItemProperty -Path $keyPath -Name 'Domain' -Value $DNSSuffix -Force
    Write-Host ("The computer DNS suffix is now set to {0}.`nPlease restart the computer." -f $DNSSuffix)
}


