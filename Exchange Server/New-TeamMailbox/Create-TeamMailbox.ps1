<#
    Create-TeamMailbox.ps1

    Helper script to simplify the use of New-TeamMailbox.ps1
#>

$teamMailboxName = 'TM-Exchange Admin'
$teamMailboxDisplayName = 'Exchange Admins'
$teamMailboxAlias = 'TM-ExchangeAdmin'
$teamMailboxSmtpAddress = 'ExchangeAdmins@mcsmemails.de'
$departmentPrefix = 'IT'
$groupFullAccessMembers = @('exAdmin1','exAdmin2')
$groupSendAsMember = @('exAdmin1','exAdmin2')

.\New-TeamMailbox.ps1 -TeamMailboxName $teamMailboxName -TeamMailboxDisplayName $teamMailboxDisplayName -TeamMailboxAlias $teamMailboxAlias -TeamMailboxSmtpAddress $teamMailboxSmtpAddress -DepartmentPrefix $departmentPrefix -GroupFullAccessMembers $groupFullAccessMembers -GroupSendAsMember $groupSendAsMember -Verbose