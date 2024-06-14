param (
  [int]$UserCount = 5,
  [string]$Company = 'Varunagroup',
  [string]$UserNameCsv = '',
  [switch]$RandomPassword,
  [string]$TestUserPrefix = 'TestUser',
  [string]$PreferredLanguage = 'de-DE',
  [string]$TargetOU = 'OU=IT,dc=varunagroup,dc=de',
  [string]$TestUserOU = 'Test User',
  [string]$UpnDomain = 'varunagroup.de'
)

$DefaultPassword = 'Pa55w.rd'

# User name prefix
# New user object will be named TestUser1, TestUser2, ...

# User object properties
$GivenName = 'Test'
$Surname = 'User'
$JobTitle = @('Junior Consultant','Senior Consultant','Technical Consultant','Business Consultant','Sales Professional','Team Lead')

# Import Active Directory PowerShell Module
Import-Module -Name ActiveDirectory

# Build OU Path
$TestUserOUPath = ('OU={0},{1}' -f $TestUserOU, $TargetOU)

# Check if OU exists
$OUExists = $false

try {
   $OUExists = [adsi]::Exists("LDAP://$TestUserOUPath")
}
catch {
   $OUExists =$true
}

if(-not $OUExists) {
   # Create new organizational unit for test users
   New-ADOrganizationalUnit -Name $TestUserOU -Path $TargetOU -ProtectedFromAccidentalDeletion:$false -Confirm:$false
}
else {
   Write-Warning -Message ('OU {0} exists please delete the OU and user objects manually, before running this script.' -f $TestUserOUPath)
   Exit
}

Write-Output -InputObject ('Creating {0} user object in {1}' -f $UserCount, $TestUserOUPath)

# Create new user objects
1..$UserCount | ForEach-Object {

   # Get a random number for selecting a job title
   $random = Get-Random -Minimum 0 -Maximum (($JobTitle | Measure-Object). Count - 1)

   # Set user password
   if($RandomPassword) {
      # Create a random password
      $UserPassword = ConvertTo-SecureString -String (-join ((33..93) + (97..125) | Get-Random -Count 25 | % {[char]$_})) -AsPlainText -Force
   }
   else {
      # Use a fixed password
      $UserPassword = ConvertTo-SecureString -String $DefaultPassword -AsPlainText -Force
   }

   # Create a new user object
   # Adjust user name template and other attributes as needed
   New-ADUser -Name ('{0}{1}' -f $TestUserPrefix, $_) `
   -DisplayName ('{0} {1}' -f $TestUserPrefix, $_) `
   -GivenName $GivenName `
   -Surname (('{0}{1}' -f $Surname, $_)) `
   -OtherAttributes @{title=$JobTitle[$random];company=$Company;preferredLanguage=$PreferredLanguage} `
   -Path $TestUserOUPath `
   -AccountPassword $UserPassword `
   -UserPrincipalName ('{0}.{1]@{2}' -f $GivenName, (('{0}{1}' -f $Surname, $_)), $UpnDomain) `
   -Enabled:$True `
   -Confirm:$false
}