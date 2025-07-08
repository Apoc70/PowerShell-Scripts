<#
    .SYNOPSIS

    Short description
   
    .DESCRIPTION

    Long description

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .NOTES 

    Requirements 

    - Windows Server 2019+
    - Exchange Online PowerShell V3 module
    - GlobalFunctions PowerShell module
    - Entra ID application with certificate-based authentication
  
    Version 1.0, 2025-07-08
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial release 

    .LINK

    https://scripts.granikos.eu

    .PARAMETER SettingsFile

    Path to the Entra settings file for app-based authentication.


#>

# Parameter section with examples
# Additional information parameters: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters
[CmdletBinding()]
param(
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Path to the Entra settings file for app-based authentication")]
    [ValidateNotNull()]
    [string]$SettingsFile
)

#region Initialize Script 

# Measure script running time
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

$script:ScriptPath = Split-Path $script:MyInvocation.MyCommand.Path
$script:ScriptName = $MyInvocation.MyCommand.Name

# Load required module for logging
if ($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
    Import-Module -Name GlobalFunctions
}
else {
    Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Import-Module GlobalFunctions'
    Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
    exit
}

# Create a logging oobject
$logger = New-Logger -ScriptRoot $script:ScriptPath -ScriptName $script:ScriptName -LogFileRetention 14
# purge old log files
$logger.Purge()
$logger.Write('Script started')

#endregion

#region Functions

<#
    Load script settings from dedicated settings.jsan 
    The settings file should contain the following properties:
    - tenantid: The tenant ID for the Entra ID application
    - clientid: The client ID of the Entra ID application
    - certThumbprint: The thumbprint of the certificate used for authentication
#>
function LoadScriptSettings {

    $configFilePath = Join-Path -Path (Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath $SettingsFile

    if (Test-Path -Path $configFilePath) {
        Write-Verbose ('Loading configuration from {0}' -f $configFilePath)
        try {
            $config = Get-Content -Path $configFilePath | ConvertFrom-Json
            # Extract configuration values
            $script:tenantId = $config.tenantid
            $script:organizationName = $config.organizationname
            $script:clientid = $config.ClientId
            $script:certThumbprint = $config.CertThumbprint
        }
        catch {
            Write-Error ('Failed to load configuration from {0}: {1}' -f $configFilePath, $_.Exception.Message)
            exit 98
        }
    }
    else {
        Write-Error ('Configuration file not found: {0}' -f $configFilePath)
        exit 99
    }
}

<#
    Function to prompt user for yes/no

    Depending on the choice order, the functions returns
    0 = YES
    1 = NO

    Example: Interactive Y/N query before running additional code

    if((Request-Choice -Caption ('Do you want to apply the settings to {0}?' -f $SomeVariable)) -eq 0) {
        # Yes Option
    }
    else {
        # No Option
    }
#>
function Request-Choice {
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Provide a caption for the Y/N question.")]
        [string]$Caption
    )
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@('&Yes', '&No')
    [int]$defaultChoice = 1

    $choiceReturn = $Host.UI.PromptForChoice($Caption, '', $choices, $defaultChoice)

    return $choiceReturn   
}

#endregion

#region MAIN

# 1. Load script settings
LoadScriptSettings

# 2. Connect to Exchange Online using the Entra ID application
Write-Host ('Connecting to Exchange Online tenant {0}' -f $script:organizationName)

Connect-ExchangeOnline -Organization $script:organizationName -AppId $script:clientid -CertificateThumbprint $script:certThumbprint
Write-Verbose -Message ('Connected to Exchange Online tenant {0}' -f $script:organizationName)

# 3. Get all unified groups
Write-Host 'Retrieving all unified groups...'  
$unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited
Write-Verbose -Message ('Retrieved {0} unified groups' -f $unifiedGroups.Count)

# 4. Set address book visibility for each unified group
Write-Host 'Setting address book visibility for each unified group...'  
foreach ($group in $unifiedGroups) {
    Write-Verbose -Message ('Processing group {0} ({1})' -f $group.DisplayName, $group.PrimarySmtpAddress)

    # Check if the group is already set to hidden from address book
    if ($group.HiddenFromAddressListsEnabled -eq $false) {
        Write-Host ('Setting address book visibility for group {0} ({1}) to hidden' -f $group.DisplayName, $group.PrimarySmtpAddress)
        Set-UnifiedGroup -Identity $group.Identity -HiddenFromAddressListsEnabled $true
        # Log the change
        $logger.Write(('Set address book visibility for group {0} ({1}) to hidden' -f $group.DisplayName, $group.PrimarySmtpAddress))
        Write-Verbose -Message ('Set address book visibility for group {0} ({1}) to hidden' -f $group.DisplayName, $group.PrimarySmtpAddress)
    }
    else {
        Write-Host ('Group {0} ({1}) is already hidden from address book' -f $group.DisplayName, $group.PrimarySmtpAddress)
    }
}


#endregion

#region End Script

# Stop watch
$StopWatch.Stop()

# Write script runtime
Write-Verbose -Message ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)
$logger.Write( ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds) )
$logger.Write('Script finished')

return 0

#endregion