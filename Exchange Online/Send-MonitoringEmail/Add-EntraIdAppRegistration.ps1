<#
    .SYNOPSIS
    Add a custom app registration to Entra ID for Microsoft Graph access

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, 2024-07-22

    Based on the work of Andres Bohren
    https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell

    .NOTES
    Requirements
    - PowerShell 7.1+
    - Account executing this script must be a role member of the Application Administrator or Global Administrator role

    Use Connect-MgGraph -Scopes "Directory.ReadWrite.All" -NoWelcome to connect to Microsoft Graph when creating the application.

    When using application permissions for Microsoft Grapg, consider restricting access to the application to specific users or groups:
    https://bit.ly/LimitExoAppAccess

#>
[CmdletBinding()]
param(
    [string]$AppName = 'CustomSendEmail-Application',
    [string]$AppSecretName = 'AppClientSecret',
    # Update the email address to the email address of the user that should be the owner of the application
    [string]$AppOwnerEmailAddress = 'admin@egxde.onmicrosoft.com',
    [int]$AppSecretLifetime = 12
)

# Load required modules

if ($null -ne (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable).Version) {
    # Import the Microsoft.Graph.Authentication module
    Import-Module -Name Microsoft.Graph.Authentication
}
else {
    # Display a warning and exit the script
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Authentication PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Import-Module Microsoft.Graph'
    exit
}
if ($null -ne (Get-Module -Name Microsoft.Graph.Applications -ListAvailable).Version) {
    # Import the Microsoft.Graph.Applications module
    Import-Module -Name Microsoft.Graph.Applications
}
else {
    # Display a warning and exit the script
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Applications PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Import-Module Microsoft.Graph'
    exit
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All", "Application.ReadWrite.All", "User.Read.All" -NoWelcome

# Create a new application
$newApp= New-MgApplication -DisplayName $AppName -Notes 'Application for sending monitoring emails. This app is used by the Send-MonitoringEmail.ps1 script.'
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
    "endDateTime" = (Get-Date).AddMonths($AppSecretLifetime)
}
$appSecret = Add-MgApplicationPassword -ApplicationId $appObjectId -PasswordCredential $newAppSecret

Write-Host 'Copy the following information to your settings file' -ForegroundColor Green
Write-Host ('ClientSecret: {0}' -f $appSecret.SecretText) -ForegroundColor Green

<#
    All permissions and IDs
    https://learn.microsoft.com/graph/permissions-reference#all-permissions-and-ids

    Mail.ReadWrite Application e2a3a72e-5f79-4c64-b1b1-878b674786c9
    Mail.Send      Application b633e1c5-b582-4048-a93e-9f11b44c7e96
    User.Read	Delegated	e1fe6dd8-ba31-4d61-89e7-88639da4683d
    User.Read.All    Application	df021288-bdef-4463-88db-98f22de89214
#>

$params = @{
    RequiredResourceAccess = @(
        @{
            ResourceAppId  = "00000003-0000-0000-c000-000000000000"
            ResourceAccess = @(
                @{
                    Id   = "df021288-bdef-4463-88db-98f22de89214"
                    Type = "Role"
                },
                @{
                    Id   = "e2a3a72e-5f79-4c64-b1b1-878b674786c9"
                    Type = "Role"
                },
                @{
                    Id   = "b633e1c5-b582-4048-a93e-9f11b44c7e96"
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