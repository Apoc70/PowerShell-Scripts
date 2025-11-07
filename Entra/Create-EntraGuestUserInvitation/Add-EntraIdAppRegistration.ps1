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

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release

    .PARAMETER AppName

    The display name of the application in Entra ID

    .PARAMETER AppSecretName

    The name of the client secret in Entra ID

    .PARAMETER AppOwnerEmailAddress

    The email address of the application owner

    .PARAMETER AppSecretValidityInMonths

    The validity of the client secret in months

#>
[CmdletBinding()]
param(
    [string]$AppName = 'EntraGuestUserInvitation-Application',
    [string]$AppSecretName = 'AppClientSecret',
    [string]$AppOwnerEmailAddress = 'Admin@TENANT.onmicrosoft.com', #Adjust to your tenant and your admin user
    [int]$AppSecretValidityInMonths = 12 
)

if ($null -ne (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable).Version) {
    Import-Module -Name Microsoft.Graph.Authentication
}
else {
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Authentication PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Microsoft.Graph'
    exit
}
if ($null -ne (Get-Module -Name Microsoft.Graph.Applications -ListAvailable).Version) {
    Import-Module -Name Microsoft.Graph.Applications
}
else {
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Applications PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Microsoft.Graph'
    exit
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All", "Application.ReadWrite.All", "User.Read.All" -NoWelcome

# Create a new application
$newApp= New-MgApplication -DisplayName $AppName -Notes 'Application for creating new guest user invites. This app is used by the Create-EntraGuestUserInvitation.ps1 script.'
$appObjectId = $newApp.Id

# Set the owner of the application
$User = Get-MgUser -UserId $AppOwnerEmailAddress
$ObjectId = $User.ID
$NewOwner = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/{$ObjectId}"
}
$null = New-MgApplicationOwnerByRef -ApplicationId $appObjectId -BodyParameter $NewOwner

# Create a new client secret for the application
$newAppSecret = @{
    "displayName" = $AppSecretName
    "endDateTime" = (Get-Date).AddMonths($AppSecretValidityInMonths)
}
$appSecret = Add-MgApplicationPassword -ApplicationId $appObjectId -PasswordCredential $newAppSecret

Write-Host 'Copy the following information to your settings file' -ForegroundColor Green
Write-Host ('ClientSecret: {0}' -f $appSecret.SecretText) -ForegroundColor Green

<#
    All permissions and IDs
    https://learn.microsoft.com/graph/permissions-reference#all-permissions-and-ids

    User.Invite.All	Application	09850681-111b-4a89-9bed-3f2cae46d706
    User.Read	Delegated	e1fe6dd8-ba31-4d61-89e7-88639da4683d
#>

$params = @{
    RequiredResourceAccess = @(
        @{
            ResourceAppId  = "00000003-0000-0000-c000-000000000000"
            ResourceAccess = @(
                @{
                    Id   = "09850681-111b-4a89-9bed-3f2cae46d706"
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