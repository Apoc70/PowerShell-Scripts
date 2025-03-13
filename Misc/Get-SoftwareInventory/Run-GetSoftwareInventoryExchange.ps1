<#
    Script to fetch software inventory across all Exchange servers and some additional servers amd send the result via email.
    Must be executed in Exchange Management Shell (EMS)

    Not fancy script, but Copyright (c) 2025 Thomas Stensitzki
#>

# Define initial server list
$server = @('SERVER01','server02')

# Fetch Exchange server list and add to server list
# Requires Exchange Management Shell
Get-ExchangeServer | Select-Object Name | ForEach-Object{$server += $_.Name} | Sort-Object

# Fetch software inventory for all servers
# Requires Get-SoftwareInventory.ps1 in the same folder as this script
.\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmail.de -MailTo it@mcsmail.de -MailServer mobile.mcsmail.de -RemoteComputer $server