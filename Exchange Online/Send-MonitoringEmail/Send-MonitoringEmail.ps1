<#
    .SYNOPSIS
    Send a single monitoring email to a recipient

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, 2024-07-22

    Please post ideas, comments, and suggestions at GitHub.

    .LINK
    https://scripts.Granikos.eu

    .DESCRIPTION

    Use the Add-EntraIdAppRegistration.ps1 script to create a custom application registration in Entra ID.

    Adjust the settings in the Settings.xml file to match your environment.

    .NOTES
    Requirements
    - PowerShell 7.1+
    - GlobalFunctions PowerShell module
    - Registered Entra ID application with access to Microsoft Graph

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release

    .PARAMETER SettingsFileName
    The file name of the settings file located in the script directory.

    .EXAMPLE
    Create calendar items for users based on the JSON file CustomEvents.json and the settings file CustomSettings.xml

    .\Add-CustomCalendarItems.ps1 -EventFileName CustomEvents.json SettingsFileName CustomSettings.xml
#>
[CmdletBinding()]
param(
    [string]$SettingsFileName = 'Settings.xml'
)

# Some general variables
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path

# Import GlobalFunctions
if ($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
    Import-Module -Name GlobalFunctions
}
else {
    Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module GlobalFunctions'
    Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further information.'
    exit
}
if ($null -ne (Get-Module -Name Microsoft.Graph.Users.Actions -ListAvailable).Version) {
    # Import the Microsoft.Graph.Users.Actions module
    Import-Module -Name Microsoft.Graph.Users.Actions
}
else {
    # Display a warning and exit the script
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Users.Actions PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Import-Module Microsoft.Graph'
    exit
}

# Load script settings
if (Test-Path -Path (Join-Path -Path $ScriptDir -ChildPath $SettingsFileName) ) {

    # Load settings from XML file
    [xml]$Config = Get-Content -Path (Join-Path -Path $ScriptDir -ChildPath $SettingsFileName)

    Write-Verbose -Message 'Loading script settings'

    #
    $tenantId = $Config.Settings.TenantId
    $clientId = $Config.Settings.ClientId
    $clientSecret = $Config.Settings.ClientSecret
    $senderEmailAddress = $Config.Settings.SenderEmailAddress
    $recipientEmailAddress = $Config.Settings.RecipientEmailAddress
    $messageSubject = $Config.Settings.MessageSubject
    $attachmentFilename = $Config.Settings.AttachmentFilename
    $messageType = $Config.Settings.MessageType
    $saveToSentItems = $Config.Settings.SaveToSentItems

    Write-Verbose -Message 'Script settings loaded'
}
else {
    # Ooops, settings XML is missing
    Write-Error -Message 'Script settings file settings.xml missing'
    exit 99
}

# DO SOME STUFF
# Create a new logger and delete log files older than 14 days
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

# Prepare client secret credential
$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass

# Connect to Microsoft Graph
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

# Load the attachment content
if ($attachmentFilename -ne '') {
    Write-Verbose -Message ('Loading attachment content "{0}"' -f $attachmentFilename)
    $attachmentPath = Join-Path -Path $ScriptDir -ChildPath $attachmentFilename
    $attachmentBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($attachmentPath))
    $attachmentName = (Get-Item -Path $attachmentPath).Name
}
else {
    $attachmentBase64 = ''
    $attachmentName = ''
}

# Prepare the monitoring email body
$messageBody = @"
This is a monitoring email sent by the monitoring script.
GUID
$(New-Guid)
TICKS
$((Get-Date).Ticks)
"@

if ($attachmentFilename -ne '') {

    Write-Verbose -Message ('Preparing message parameters with attachment "{0}"' -f $attachmentFilename)

    $params = @{
        Message         = @{
            Subject      = $messageSubject
            Body         = @{
                ContentType = $messageType
                Content     = $messageBody
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $recipientEmailAddress
                    }
                }
            )
            Attachments  = @(
                @{
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    Name          = $attachmentName
                    ContentType   = "application/pdf"
                    ContentBytes  = $attachmentBase64
                }
            )
        }
        SaveToSentItems = $saveToSentItems
    }
}
else {

    Write-Verbose -Message 'Preparing message parameters without attachment'

    $params = @{
        Message         = @{
            Subject      = $messageSubject
            Body         = @{
                ContentType = $messageType
                Content     = $messageBody
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $recipientEmailAddress
                    }
                }
            )
        }
        SaveToSentItems = $saveToSentItems
    }
}

# Send the monitoring email
Write-Verbose -Message ('Sending monitoring email to {0}' -f $recipientEmailAddress)
Send-MgUserMail -UserId $senderEmailAddress -BodyParameter $params

# Disconnect from Microsoft Graph
$null = Disconnect-MgGraph

# Log the end of the script
$logger.Write('Script finished')