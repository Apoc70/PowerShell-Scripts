<#
    .SYNOPSIS

    Short description

    Remove any comment section not used, e.g., LINK, INPUTS, or OUTPUTS
    Additonal information on comment based help: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_comment_based_help

    .DESCRIPTION

    Long description

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .NOTES

    Requirements

    - Windows Server 2019+

    Revision History
    --------------------------------------------------------------------------------
    1.0 Initial release
    1.1 xxxx

    .LINK

    https://scripts.granikos.eu

    .PARAMETER ExportFileWithoutDefaultValue

    Madatory export file name

    .PARAMETER ExportFileWithDefaultValue

    Export file name example using a default value of 'ExportFile.csv''

    .EXAMPLE

    Get-SomeCmdlet.ps1 -SomeParameter1 'ExportToUTF8.csv'

    Executes the script and exports the gathered information to a CSV file named ExportToUTF8.csv

#>

# Parameter section with examples
# Additional information parameters: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters
[CmdletBinding()]
param(
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Full path of the output file to be generated. If only filename is specified, then the output file will be generated in the current directory.")]
    [ValidateNotNull()]
    [string]$ExportFileWithoutDefaultValue,
    [string]$ExportFileWithDefaultValue = 'ExportFile.csv'
)

#region Initialize Script

# Measure script running time
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

$script:ScriptPath = Split-Path $script:MyInvocation.MyCommand.Path
$script:ScriptName = $MyInvocation.MyCommand.Name

# Load required module for logging
if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
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
$logger.Purge()
$logger.Write('Script started')

#endregion

#region Functions

<#
    Load script settings from dedicated settings.xml
    The following comment block contains an XML example

    <?xml version="1.0"?>
    <Settings>
        <Group>
            <SomeValue>4711</SomeValue>
        </Group>
    </Settings>

#>
function LoadScriptSettings {
    if (Test-Path -Path ('{0}\Settings.xml' -f $script:ScriptPath)) {
        # Load Script settings
        [xml]$Config = Get-Content -Path ('{0}\Settings.xml' -f $script:ScriptPath)

        Write-Verbose -Message 'Loading script settings'

        # Group settings
        $someValue = $Config.Settings.Group.SomeValue

        Write-Verbose -Message 'Script settings loaded'
    }
    else {
        Write-Error -Message 'Script settings file settings.xml missing. Please check documentation.'
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
            Mandatory=$true,
            HelpMessage = "Provide a caption for the Y/N question.")]
        [string]$Caption
    )
    $choices =  [System.Management.Automation.Host.ChoiceDescription[]]@('&Yes','&No')
    [int]$defaultChoice = 1

    $choiceReturn = $Host.UI.PromptForChoice($Caption, '', $choices, $defaultChoice)

    return $choiceReturn
}

#endregion

#region MAIN

# 1. Load script settings
LoadScriptSettings

<#
    The main code
#>

#endregion

#region End Script

# Stop watch
$StopWatch.Stop()

# Write script runtime
Write-Verbose -Message ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)
$logger.Write( ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds) )
$logger.Write('Script finished')

#endregion