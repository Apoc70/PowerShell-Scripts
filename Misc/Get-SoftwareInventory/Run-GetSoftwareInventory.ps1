<#
    Predefined script to fetch software inventory and send email report

    Not a real script, but Copyright (c) 2025 Thomas Stensitzki
#>

# Define  server list
$server = @('SERVER01','server02')

.\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmemail.de -MailTo it@mcsmemail.de -MailServer myserver@mcsmemail.de -RemoteComputer $server