[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$GroupId = ''
)

# Verify that the Microsoft.Graph module is loaded, import if not
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph

# Connect to Microsoft Graph with Group.Read.All and User.Read.All permissions
$scopes = @("Group.Read.All", "User.Read.All")
Connect-MgGraph -Scopes $scopes -ContextScope CurrentUser -NoWelcome -TenantId ''

$members = Get-MgGroupMember -GroupId $GroupId -All 

$memberEmailAddresses = @()

foreach ($member in $members) {

    $user = Get-MgUser -UserId $member.UserId
    Write-Verbose $user

    $memberEmailAddresses += $user.Mail
    
}

$memberEmailAddresses