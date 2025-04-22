# Abrufen aller Netzwerkschnittstellen und ihrer IP-Adressen
$networkAdapters = Get-NetIPAddress | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" }

foreach ($adapter in $networkAdapters) {
    $interfaceAlias = $adapter.InterfaceAlias
    $ipAddress = $adapter.IPAddress
    $addressFamily = $adapter.AddressFamily
    $dhcpEnabled = (Get-NetIPInterface -InterfaceAlias $interfaceAlias -AddressFamily $addressFamily).Dhcp

    # Ausgabe der Netzwerkkonfiguration je nach Adressfamilie
    if ($addressFamily -eq "IPv4") {
        Write-Output "Netzwerkadapter: $interfaceAlias (IPv4)"
        Write-Output "IP-Adresse     : $ipAddress"
        Write-Output "DHCP aktiviert : $dhcpEnabled"
        Write-Output "-----------------------------"

        # Überprüfen, ob DHCP deaktiviert ist (statische IP-Adresse)
        if ($dhcpEnabled -eq "Disabled") {
            Write-Output "> Der Adapter $interfaceAlias (IPv4) ist für eine fest zugewiesene IP-Adresse konfiguriert."
        } else {
            Write-Output "> Der Adapter $interfaceAlias (IPv4) verwendet DHCP."
        }
        Write-Output "============================="
    } elseif ($addressFamily -eq "IPv6") {
        Write-Output "Netzwerkadapter: $interfaceAlias (IPv6)"
        Write-Output "IP-Adresse     : $ipAddress"
        Write-Output "DHCP aktiviert : $dhcpEnabled"
        Write-Output "-----------------------------"

        # Überprüfen, ob DHCP deaktiviert ist (statische IP-Adresse)
        if ($dhcpEnabled -eq "Disabled") {
            Write-Output "> Der Adapter $interfaceAlias (IPv6) ist für eine fest zugewiesene IP-Adresse konfiguriert."
        } else {
            Write-Output "> Der Adapter $interfaceAlias (IPv6) verwendet DHCP."
        }
        Write-Output "============================="
    }
}
