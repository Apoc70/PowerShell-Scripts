<#
    .SYNOPSIS
    This Script creates a new B2B guest user invitation and optionally sends the invitation or prints the redemption URL to the screen

    .DESCRIPTION
    The script uses the Graph API to create the invitation and send it to the user. The script
    can be used to create a new guest user in Entra ID and send them an invitation to join the tenant.

    .NOTES
    Version: 1.0, 2025-04-25

    Credits to Sean McAvinue (https://seanmcavinue.net) for the original script

    .LINK
    https://github.com/smcavinue/AdminSeanMc/tree/master/Graph%20Scripts

    .LINK
    https://github.com/Apoc70

    .PARAMETER UserDisplayName
    Display name for the new guest account

    .PARAMETER UserEmail
    Email address of the requested external user

    .PARAMETER UserMessage
    Custom message to present to the user, can be used when sending intivation automatically.

    .PARAMETER clientSecret
    Entra app registration client secret

    .PARAMETER clientID
    Entra app clientID

    .PARAMETER tenantID
    Entra ID of the M365 tenant

    .PARAMETER RedirectURL
    A URL to redrect to after the invitation is redeemed.
    By default Entra ID will redirect to https://myapps.microsoft.com which is not the best option for an external user.
    Recommendation: Redirect the invited user to a dedicated landing page.

    .PARAMETER SendInvite
    Boolean to control whether the invite is sent to the user or not. $True will send the invite and $False will only create the invitation and return the redemption URL.

    .EXAMPLE
    To send a B2B guest user  invitation and trigger an email
    .\Create-EntraGuestUserInvitation.ps1 -UserDisplayName 'John Doe' -UserEmail jd@varunagroup.de -ClientSecret $clientSecret -TenantID $tenantID -ClientID $clientID -SendInvite -UserMessage "This is your Guest user invitation to the Varunagroup Tenant, please contact your account manager or check out our information on https://varunagroup.de"

    .EXAMPLE
    To create an invitation and return the redemption URL
     .\Create-EntraGuestUserInvitation.ps1 -UserDisplayName 'John Doe' -UserEmail jd@varunagroup.de -ClientSecret $clientSecret -TenantID $tenantID -ClientID $clientID
    #>

[CmdletBinding()]
Param(
    [parameter(Mandatory = $true)]
    [String]$UserDisplayName,
    [parameter(Mandatory = $true)]
    [String]$UserEmail,
    [String]$UserMessage = '',
    [parameter(Mandatory = $true)]
    [String]$ClientSecret,
    [parameter(Mandatory = $true)]
    [String]$ClientID,
    [parameter(Mandatory = $true)]
    [String]$TenantID,
    [String]$RedirectURL = "https://myapps.microsoft.com",
    [ValidateSet("nl-NL", "en-US", "fr-FR", "de-DE", "es-ES", "sv-SE", "fi-FI", "da-DK", "no-NO", "it-IT", "pt-PT", "ru-RU", "ja-JP", "zh-CN", "cs-CZ", "hu-HU", "pl-PL", "tr-TR", "ar-SA", "he-IL", "th-TH", "ko-KR", "el-GR", "id-ID", "ms-MY", "tl-PH", "vi-VN")]
    [String]$MessageLanguage = "de-DE",
    [Boolean]$SendInvite
)

<#
Some links to the documentation:
https://learn.microsoft.com/en-us/entra/external-id/invitation-email-elements
https://learn.microsoft.com/en-us/graph/api/resources/invitedusermessageinfo?view=graph-rest-1.0
https://learn.microsoft.com/en-us/powershell/module/microsoft.entra/new-entrainvitation?view=entra-powershell
https://practical365.com/creating-custom-b2b-guest-user-invitations-with-graph-api/
#>

function Get-MSGraphToken {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $true)]
        [String]$ClientSecret,
        [parameter(Mandatory = $true)]
        [String]$ClientID,
        [parameter(Mandatory = $true)]
        [String]$TenantID
    )

    # Construct Invoke-WebRequest URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Invoke-WebMethod Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    # Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

    # Fetch access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    return $token
}

function New-UserInvitation {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $true)]
        [String]$UserDisplayName,
        [parameter(Mandatory = $true)]
        [String]$UserEmail,
        [parameter(Mandatory = $true)]
        [String]$RedirectURL,
        [parameter(Mandatory = $true)]
        [bool]$SendInvite,
        [parameter(Mandatory = $true)]
        [String]$UserMessage,
        [parameter(Mandatory = $true)]
        [String]$MessageLanguage

    )

    Write-Host "Creating User Invitation with the following settings:"
    #nl-NL, en-US, fr-FR, de-DE, es-ES, sv-SE, fi-FI, da-DK, no-NO, it-IT, pt-PT, ru-RU, ja-JP, zh-CN, cs-CZ, hu-HU, pl-PL, tr-TR, ar-SA, he-IL, th-TH, ko-KR, el-GR, id-ID, ms-MY, tl-PH, vi-VN
    $InvitationObject = @"
    {
        "invitedUserDisplayName": "$UserDisplayName",
        "invitedUserEmailAddress": "$UserEmail",
        "sendInvitationMessage": "$SendInvite",
        "inviteRedirectUrl": "$RedirectURL",
        "invitedUserType": "Guest",
        "invitedUserMessageInfo": {
            "messageLanguage": "$MessageLanguage",
            "customizedMessageBody": "$UserMessage"
        }
    }
"@
    $InvitationObject | Out-Host
    return $InvitationObject
}

# Get MS Graph token
$token = Get-MSGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID
# MS Graph API URI
$apiUri = 'https://graph.microsoft.com/beta/invitations/'

# Create the invitation
$body = New-UserInvitation -UserDisplayName $UserDisplayName -UserEmail $UserEmail -RedirectURL $RedirectURL -SendInvite $SendInvite -UserMessage $UserMessage -MessageLanguage $MessageLanguage

try {
    $invitation = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)" } -Uri $apiUri -Method Post -ContentType 'application/json' -Body $body)

    if ($SendInvite) {
        Write-Host ('Invitation has been sent to {0}' -f $UserEmail)
    }
    else {

        $inviteObject = [PSCustomObject]@{
            DisplayName = $userDisplayName
            Email       = $UserEmail
            URL         = $invitation.inviteRedeemUrl
        }

        # Export the invitation URL to a CSV file
        $inviteObject | Export-Csv -Path "InvitationURLs.csv" -NoClobber -NoTypeInformation -Append -Encoding UTF8 -Force
        # Display the invitation URL
        Write-Host ('Invitation Redemption URL is: {0} and has been exported to InvitationURLs.csv file' -f $invitation.inviteRedeemUrl)
    }
}
catch {
    Write-Host ('Error creating invitation: {0}' -f $_.Exception.Message)
}
