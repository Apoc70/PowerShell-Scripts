# Name of Entra Connect AD connect
$adConnector  = "varunagroup.de"

# Name ENtra Connect Entra connector
$aadConnector = "varunagroup.onmicrosoft.com - AAD"


Import-Module adsync
$c = Get-ADSyncConnector -Name $adConnector
$p = New-Object Microsoft.IdentityManagement.PowerShell.ObjectModel.ConfigurationParameter "Microsoft.Synchronize.ForceFullPasswordSync", String, ConnectorGlobal, $null, $null, $null
$p.Value = 1
$c.GlobalParameters.Remove($p.Name)
$c.GlobalParameters.Add($p)
$c = Add-ADSyncConnector -Connector $c


Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $adConnector -TargetConnector $aadConnector -Enable $false
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $adConnector -TargetConnector $aadConnector -Enable $true


Get-Mailbox -ResultSize Unlimited | Get-MailboxFolderStatistics -IncludeAnalysis -FolderScope All | Where-Object {(($_.TopSubjectSize -Match "MB") -and ($_.TopSubjectSize -GE 50.0)) -or ($_.TopSubjectSize -Match "GB")} | Select-Object Identity, TopSubject, TopSubjectSize | Export-CSV -path "C:\report.csv" -notype