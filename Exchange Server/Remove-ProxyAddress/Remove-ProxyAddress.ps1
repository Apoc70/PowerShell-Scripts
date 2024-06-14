# Description: Remove email addresses from users in Active Directory

$DomainFilter = "*mail.onmicrosoft.com"

# Get all users with email addresses in scope
Get-ADUser -Properties proxyaddresses -Filter {ProxyAddresses -like $DomainFilter} |
    ForEach-Object {
        # Remove the email addresses
        ForEach ($proxyAddress in $_.proxyAddresses) {
            # Check if the email address is in scope
            If ($proxyAddress -like $DomainFilter) {
                # Output the action to the console
                Write-Verbose -Message ('Removing $proxyAddress from {0}' -f $_.SamAccountName)
                # Remove the email address
                Set-ADUser $_.SamAccountName -Remove @{ProxyAddresses=$proxyAddress}
            }
        }
    }