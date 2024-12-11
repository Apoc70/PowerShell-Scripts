$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

# Use current date as file timestamp
$now = Get-Date -Format 'yyyy-MM-dd'

$cssFile = Join-Path -Path $ScriptDir -ChildPath styles.css

$reportFileName = "Public Folder Permissions - $($now).csv"

$pf = Get-PublicFolder -ResultSize Unlimited -Recurse

$pf | .\Report-PFPermissions.ps1 | Export-Csv -Path (Join-Path -Path $ScriptDir -ChildPath $reportFileName) -NoTypeInformation -Encoding UTF8 -Force