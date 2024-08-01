<#
    .SYNOPSIS
    Add calendar items to the default calendar of users in a security group

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, 2024-06-27

    Please post ideas, comments, and suggestions at GitHub.

    .LINK
    https://scripts.Granikos.eu

    .DESCRIPTION

    The script reads a JSON file with event data and creates or deletes calendar events in the default calendar of users in a security group.

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

    .PARAMETER EventFileName
    The name of the JSON file containing event data and action located in the script directory.

    .PARAMETER SettingsFileName
    The file name of the settings file located in the script directory.

    .EXAMPLE
    Create calendar items for users based on the JSON file CustomEvents.json and the settings file CustomSettings.xml

    .\Add-CustomCalendarEvents.ps1 -EventFileName CustomEvents.json SettingsFileName CustomSettings.xml
#>
[CmdletBinding()]
param(
    [string]$EventFileName = 'events.json',
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

# Load script settings
if (Test-Path -Path (Join-Path -Path $ScriptDir -ChildPath $SettingsFileName) ) {

    # Load settings from XML file
    [xml]$Config = Get-Content -Path (Join-Path -Path $ScriptDir -ChildPath $SettingsFileName)

    Write-Verbose -Message 'Loading script settings'

    #
    $SecurityGroupId = $Config.Settings.SecurityGroupId
    $tenantId = $Config.Settings.TenantId
    $clientId = $Config.Settings.ClientId
    $clientSecret = $Config.Settings.ClientSecret

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

#Connect to GRAPH API
<#
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}
#>
# Get the access token
# $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody

# Set the authorization headers
<#
$headers = @{
    "Authorization" = "Bearer $($tokenResponse.access_token)"
    "Content-type"  = "application/json"
}
    #>

# Do some stuff

# Prepare client secret credential
$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass

# Connect to Microsoft Graph
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

# Read JSON input file
if (Test-Path -Path (Join-Path -Path $ScriptDir -ChildPath $EventFileName) ) {

    # Read the JSON file with events
    $jsonFilePath = Join-Path -Path $ScriptDir -ChildPath $EventFileName
    $jsonData = Get-Content -Path $jsonFilePath | ConvertFrom-Json

    Write-Host -Message ('Found {0} events in the JSON file' -f $jsonData.Count)
    $logger.Write('Found {0} events in the JSON file' -f $jsonData.Count)

    # Check if we have events to process
    if ($jsonData.Count -gt 0) {

        # Fetch security group members
        $securityGroupMembers = Get-MgGroupMember -GroupId $SecurityGroupId -All

        # Check if we have members in the security group
        if ($securityGroupMembers.Count -gt 0) {
            # we have members in the security group
            Write-Host -Message ('Found {0} members in the security group' -f $securityGroupMembers.Count)
            $logger.Write( ('Found {0} members in the security group {1}' -f ($securityGroupMembers.Count), $SecurityGroupId) )

            # Loop through each member of the security group
            foreach ($member in $securityGroupMembers) {

                $user = Get-MgUser -UserId $member.Id

                Write-Host -Message ('Processing member {0}' -f $user.UserPrincipalName)
                $logger.Write( ('Processing member {0}' -f $user.UserPrincipalName) )

                # Check if the user has an existing calendar named 'Calendar'
                $userCalendars = Get-MgUserCalendar -UserId $user.Id
                $calendar = $userCalendars | Where-Object { $_.IsDefaultCalendar -eq $true } # { ($_.Name -eq 'Calendar') -or ($_.Name -eq 'Kalender')}

                if ($null -ne $calendar) {
                    # We have a calendar named 'Calendar'
                    Write-Host -Message ('Found calendar named "Calendar" for user {0}' -f $user.UserPrincipalName)

                    # Initialize counters
                    $calendarEventsAdded = 0
                    $calendarEventsSkipped = 0
                    $calendarEventsDeleted = 0

                    # Loop through each event in the JSON data and create calendar items as needed
                    foreach ($event in $jsonData) {

                        Write-Verbose ('- Processing event "{0}" [Start: {1} | Action: {2}]' -f $event.eventData.Subject, $event.eventData.Start.DateTime, $event.eventAction)

                        # Parse event action from JSON data
                        $eventAction = ([string]$event.eventAction).ToUpper()

                        # Check if the event already exists
                        $start = ([datetime]([string]$event.eventData.Start.DateTime))
                        $existingEvent = Get-MgUserCalendarEvent -UserId $user.Id -CalendarId $calendar.Id | Where-Object { $_.Subject -eq ([string]$event.eventData.Subject) -and ( ([DateTime]$_.Start.DateTime) -eq $start ) }

                        if (($null -eq $existingEvent) -and ($eventAction -eq 'CREATE')) {
                            # Event does not exist, and eventAction is CREATE, create the event

                            # Write event creation information to the console and log file
                            $message = ('- Creating event {0} [Start {2}] for user {1}' -f $event.eventData.Subject, $user.UserPrincipalName, $event.eventData.Start.DateTime)
                            Write-Host -Message $message
                            $logger.Write($message)

                            # Convert the event to JSON
                            $eventJson = $event.eventData | ConvertTo-Json -Depth 10

                            # Create the event
                            $null = New-MgUserCalendarEvent -CalendarId $calendar.Id -UserId $user.Id-BodyParameter $eventJson

                            # Increment the counter
                            $calendarEventsAdded++
                        }
                        elseif (($null -ne $existingEvent) -and ($eventAction -eq 'DELETE')) {
                            # Event exists, and eventAction is DELETE, delete the event

                            # Write event deletion information to the console and log file
                            $message = ('- DELETING event {0} [Start {2}] for user {1}' -f $event.eventData.Subject, $user.UserPrincipalName, $event.eventData.Start.DateTime)
                            Write-Host -Message $message
                            $logger.Write($message)

                            # Delete the event
                            $null = Remove-MgUserEvent -UserId $user.Id -EventId $existingEvent.Id

                            # Increment the counter
                            $calendarEventsDeleted++
                        }
                        else {
                            # Event already exists
                            Write-Host -Message ('- Event "{0}" [Start {1}] already exists and is skipped' -f $event.eventData.Subject, $event.eventData.Start.DateTime)

                            $calendarEventsSkipped++
                        }
                    }
                }
                else {
                    # We have a calendar named 'Calendar'

                    Write-Host -Message ('No calendar named "Calendar" found for user {0}' -f $user.UserPrincipalName)
                }

                $logger.Write( ('Processed {0} events for user {1} - Added: {2} - Skipped: {3} - Deleted: {4}' -f $jsonData.Count, $user.UserPrincipalName, $calendarEventsAdded, $calendarEventsSkipped, $calendarEventsDeleted) )
            }
        }
        else {
            # we have no members in the security group
            Write-Error -Message ('No members found in the security group {0}' -f $securityGroupId)
            exit 1
        }
    }
    else {
        Write-Information ('No events found in the JSON file {0}' -f $EventFileName)
        exit 1
    }

}
else {
    Write-Error -Message ('File {0} not found' -f $EventFileName)
    exit 1
}

# Disconnect from Microsoft Graph
$null = Disconnect-MgGraph

# Log the end of the script
$logger.Write('Script finished')