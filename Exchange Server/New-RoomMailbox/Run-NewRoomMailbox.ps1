<#
    Run-NewRoomMailbox.ps1

    Helper script to simplify the use of New-RoomMailbox.ps1
#>

$roomMailboxName = 'MB-Conference Room'
$roomMailboxDisplayName = 'Board Conference Room'
$roomMailboxAlias = 'MB-ConferenceRoom'
$roomMailboxSmtpAddress = 'ConferenceRoom@mcsmemail.de'
$departmentPrefix = 'C'
$groupFullAccessMembers = @('JohnDoe','JaneDoe')  # Empty = @()
$groupSendAsMembers = @()
$groupCalendarBookingMembers = @()
$RoomCapacity = 0
$RoomList = 'AllRoomsHQ'
$Language = 'de-DE' # en-US, de-DE, fr-FR, ja-JP, pt-BR, zh-CN, zh-TW, ...

.\New-RoomMailbox.ps1 -RoomMailboxName $roomMailboxName -RoomMailboxDisplayName $roomMailboxDisplayName -RoomMailboxAlias $roomMailboxAlias -RoomMailboxSmtpAddress $roomMailboxSmtpAddress -DepartmentPrefix $departmentPrefix -GroupFullAccessMembers $groupFullAccessMembers -GroupSendAsMembers $groupSendAsMembers -RoomCapacity $RoomCapacity -AutoAccept -RoomList $RoomList -Language $Language

if ($roomMailboxSmtpAddress -ne '') {
    # Use the provided room mailbox SMTP address
    .\New-RoomMailbox.ps1 -RoomMailboxName $roomMailboxName -RoomMailboxDisplayName $roomMailboxDisplayName -RoomMailboxAlias $roomMailboxAlias -RoomMailboxSmtpAddress $roomMailboxSmtpAddress -GroupFullAccessMembers $groupFullAccessMembers -GroupSendAsMember $groupSendAsMembers -RoomCapacity $RoomCapacity -AutoAccept -RoomList $RoomList -Language $Language
}
else {
    # Generate the room mailbox SMTP address automatically
    .\New-RoomMailbox.ps1 -roomMailboxName $roomMailboxName -RoomMailboxDisplayName $roomMailboxDisplayName -roomMailboxAlias $roomMailboxAlias -GroupFullAccessMembers $groupFullAccessMembers -GroupSendAsMember $groupSendAsMembers -RoomCapacity $RoomCapacity -AutoAccept -RoomList $RoomList -Language $Language
}