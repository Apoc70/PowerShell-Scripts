# This script retrieves the version of the Active Directory schema used by Exchange.
# It uses the ADSI provider to access the Active Directory schema and retrieves the version information.

# schema and object version strings
# The version strings are used to map the schema and object versions to the corresponding Exchange version.
# The mapping is done using a hashtable, where the keys are the schema and object version strings, and the values are the corresponding Exchange version strings.
# The hashtable is used to simplify the mapping process and make it easier to retrieve the Exchange version based on the schema and object versions.
$VersionStrings = @{
    '17003.16763.13243' = @{Long = 'Exchange Server 2019 CU15'; Short = 'Ex19CU15' }
    '17003.16762.13243' = @{Long = 'Exchange Server 2019 CU14'; Short = 'Ex19CU14' }
    '17003.16761.13243' = @{Long = 'Exchange Server 2019 CU13'; Short = 'Ex19CU13' }
    '17003.16760.13243' = @{Long = 'Exchange Server 2019 CU12'; Short = 'Ex19CU12' }
    '17003.16759.13242' = @{Long = 'Exchange Server 2019 CU11'; Short = 'Ex19CU11' }

    '15334.16223.13243' = @{Long = 'Exchange Server 2016 CU23'; Short = 'Ex16CU23' }
    '15334.16222.13242' = @{Long = 'Exchange Server 2016 CU22'; Short = 'Ex16CU22' }

    '15312.16133.13237' = @{Long = 'Exchange Server 2013 CU23'; Short = 'Ex13CU23' }
}


$exchangeVersion = $null

# Forest version
# This script retrieves the version of the Active Directory schema used by Exchange.
$NetBiosDomainName = 'varunagroup'

$RootDSE = ([ADSI]"").distinguishedName
$schemaVersion = ([ADSI]"LDAP://CN=ms-Exch-Schema-Version-Pt,CN=Schema,CN=Configuration,$RootDSE").rangeUpper.ToString()
$adObjectVersion = ([ADSI]"LDAP://cn=$($NetBiosDomainName),cn=Microsoft Exchange,cn=Services,cn=Configuration,$RootDSE").objectVersion.ToString()
$NetBiosDomainNameObjectVersion = ([ADSI]"LDAP://CN=Microsoft Exchange System Objects,$RootDSE").objectVersion.ToString()

Write-Host ('Schema rangeUpper   : {0}' -f $schemaVersion)
Write-Host ('Forest objectVersion: {0}' -f $adObjectVersion )

# Domain version

$RootDSE = ([ADSI]"").distinguishedName
Write-Host ('Domain objectVersion: {0}' -f $NetBiosDomainNameObjectVersion)

# Exchange version
$exchangeVersion = $VersionStrings[('{0}.{1}.{2}' -f $schemaVersion, $adObjectVersion, $NetBiosDomainNameObjectVersion)].Long
if ($null -eq $exchangeVersion) {
    Write-Host 'No matching Exchange schema/object version found!'
}
else {
    Write-Host ('Exchange Version    : {0}' -f $exchangeVersion)
}