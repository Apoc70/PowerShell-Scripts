# This script calls a dedicated PowerShell script to create a B2B guest user invitation in Entra ID.
# The script must reside in the very same folder as this script

# The display name of the person requesting the B2B user invite
$InvitingPerson = 'Max Mustermann'

# Message to be sent to the invited user as part of the invitation
$UserMessage = ('Willkommen bei Varunagroup. Die Einladung erfolgt auf Wunsch von {0}.' -f $InvitingPerson )

# URL to redirect to after the invitation is redeemed
$RedirectURL = 'https://teams.microsoft.com'

# Call the Create-EntraGuestUserInvitation script with the parameters
.\Create-EntraGuestUserInvitation.ps1 `
-ClientSecret 'XXXXXX' `
-TenantID '' `
-ClientID '' `
-UserDisplayName 'John Doe' `
-UserEmail 'john.doe@example.com' `
-MessageLanguage 'de-DE' `
-SendInvite $true `
-UserMessage $UserMessage `
-RedirectURL $RedirectURL