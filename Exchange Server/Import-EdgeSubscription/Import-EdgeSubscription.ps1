<#
.SYNOPSIS
This script imports an Edge Subscription file for a specific Active Directory site.

Thomas Stensitzki

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Version 1.0, 2024-02-15

Please send ideas, comments and suggestions to support@granikos.eu

.LINK
https://scripts.granikos.eu

.DESCRIPTION
The script takes two parameters: the path to the Edge Subscription file and the name of the Active Directory site.
It then imports the Edge Subscription file for the specified Active Directory site and does not create an Internet send connector.
This script is primarily used for hybrid Exchange deployments where a different internet send connector already exists.

.PARAMETER edgeSubscriptionFile
The full path to the Edge Subscription file.

.PARAMETER activeDirectorySite
The name of the Active Directory site for subscribign the Edge Transport Server to.

.EXAMPLE
.\Import-EdgeSubscription.ps1 -edgeSubscriptionFile "C:\Import\EdgeSubscription.xml" -activeDirectorySite "ADSiteName"

#>
param(
    [string]$edgeSubscriptionFile = "C:\Import\EdgeSubscription.xml",
    [string]$activeDirectorySite = "ADSiteName"
)

Import-EdgeSubscription -FileData ([byte[]]$(Get-Content -Path $edgeSubscriptionFile -Encoding Byte -ReadCount 0)) -Site $activeDirectorySite -CreateInternetSendConnector $false

Write-Output @(
    ('Imported Edge Subscription from {0} for Active Directory site {1}' -f $edgeSubscriptionFile, $activeDirectorySite),
    'One of the Exchange Servers in the Active Directory site will use the imported BootStrap account to initiate the Edge Subscription process.',
    'Use Test-EdgeSynchronization to verify that Exchange mailbox servers communicate with subscribed Edge Transport Server.',
    'More information: https://learn.microsoft.com/powershell/module/exchange/test-edgesynchronization?view=exchange-ps'
)


