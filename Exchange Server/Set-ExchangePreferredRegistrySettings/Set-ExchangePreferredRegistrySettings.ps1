write-host-Message "Setting preferred registry settings for Exchange Server 2016/2019"

# Set KeepAliveTime to 2 hours
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -Name 'KeepAliveTime' -Value '001b7740' -PropertyType 'DWord' -Force | Out-Null

# Set KeepAliveInterval to 1 second
New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Rpc' -Name 'EnableTcpPortScaling' -Value '00000001' -PropertyType 'DWord' -Force | Out-Null

# Set MinimumConnectionTimeout to 120 seconds
New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Rpc' -Name 'MinimumConnectionTimeout' -Value '00000078' -PropertyType 'DWord' -Force | Out-Null

# Set IVv6 to disabled
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' -Name 'DisabledComponents' -Value '000000ff' -PropertyType 'DWord' -Force | Out-Null

Write-Host 'Preferred registry settings have been applied. PLease restart the server to apply changes.'