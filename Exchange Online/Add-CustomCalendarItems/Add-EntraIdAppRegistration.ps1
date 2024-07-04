<#
    .SYNOPSIS
    Add a custom app registration to Entra ID for Microsoft Graph access

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, 2024-06-27

    Based on the work of Andres Bohren
    https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell

    .NOTES
    Requirements
    - PowerShell 7.1+
    - Account executing this script must be a role member of the Application Administrator or Global Administrator role

#>
[CmdletBinding()]
param(
    [string]$AppName = 'CustomCalendarItems-Application',
    [string]$AppSecretName = 'AppClientSecret',
    [string]$AppOwnerEmailAddress = 'Admin@TENANT.onmicrosoft.com'
)

if ($null -ne (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable).Version) {
    Import-Module -Name Microsoft.Graph.Authentication
}
else {
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Authentication PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Import-Module Microsoft.Graph'
    exit
}
if ($null -ne (Get-Module -Name Microsoft.Graph.Applications -ListAvailable).Version) {
    Import-Module -Name Microsoft.Graph.Applications
}
else {
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Applications PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Import-Module Microsoft.Graph'
    exit
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All", "Application.ReadWrite.All", "User.Read.All" -NoWelcome

# Create a new application
$newApp= New-MgApplication -DisplayName $AppName -Notes 'Application for adding custom events to user calendars. This app is used by the Add-CustomCalendarItems.ps1 script.'
$appObjectId = $newApp.Id

# Set the owner of the application
$User = Get-MgUser -UserId $AppOwnerEmailAddress
$ObjectId = $User.ID
$NewOwner = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/{$ObjectId}"
}
$null = New-MgApplicationOwnerByRef -ApplicationId $appObjectId -BodyParameter $NewOwner

# Create a new application client secret with a validity of 12 months
$newAppSecret = @{
    "displayName" = $AppSecretName
    "endDateTime" = (Get-Date).AddMonths(+12)
}
$appSecret = Add-MgApplicationPassword -ApplicationId $appObjectId -PasswordCredential $newAppSecret

Write-Host 'Copy the following information to your settings file' -ForegroundColor Green
Write-Host ('ClientSecret: {0}' -f $appSecret.SecretText) -ForegroundColor Green

<#
    All permissions and IDs
    https://learn.microsoft.com/graph/permissions-reference#all-permissions-and-ids

    Calendars.ReadWrite	Application	ef54d2bf-783f-4e0f-bca1-3210c0444d99
    GroupMember.Read.All	Application	98830695-27a2-44f7-8c18-0c3ebc9698f6
    User.ReadBasic.All	Application	97235f07-e226-4f63-ace3-39588e11d3a1
    User.Read	Delegated	e1fe6dd8-ba31-4d61-89e7-88639da4683d
#>

$params = @{
    RequiredResourceAccess = @(
        @{
            ResourceAppId  = "00000003-0000-0000-c000-000000000000"
            ResourceAccess = @(
                @{
                    Id   = "98830695-27a2-44f7-8c18-0c3ebc9698f6"
                    Type = "Role"
                },
                @{
                    Id   = "ef54d2bf-783f-4e0f-bca1-3210c0444d99"
                    Type = "Role"
                },
                @{
                    Id   = "97235f07-e226-4f63-ace3-39588e11d3a1"
                    Type = "Role"
                },
                @{
                    Id   = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"
                    Type = "Scope"
                }
            )
        }
    )
}

# Add permissions to the application
$null = Update-MgApplication -ApplicationId $appObjectId -BodyParameter $params

# Return the application ID for the settings file
Write-Host ('ClientId (App ID): {0}' -f $newApp.AppId) -ForegroundColor Green

# Set the application as a public client with a redirect URI
$RedirectURI = @()
$RedirectURI += "https://login.microsoftonline.com/common/oauth2/nativeclient"

$params = @{
    RedirectUris = @($RedirectURI)
}

$null = Update-MgApplication -ApplicationId $appObjectId -IsFallbackPublicClient -PublicClient $params

# Open browser to grant admin consent
$URL = ('https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/{0}' -f $newApp.AppId)

# Open the browser to grant admin consent
# Wait for 30 seconds to allow Entra ID to provision the application
Write-Host 'Browser will open in 30 seconds.'
Start-Sleep -Seconds 30
Start-Process $URL