# Send Connector Parameters

# The name of the connector
$sendConnectorName = "E-Mail to Internet via ExchangeOnline"

# Use the same Fqdn as the one used for your hybrid send connector created by HCW
$sendConnectorFqdn = "smtpo.varunagroup.de"

# Source transpport server(s) for the send connector
# Might be an Edge Transport Server
$sendConnectorSourceTransportServers = "EX01"

# Maximum message size for the send connector
$sendConnectorMaxMessageSize = 100MB

# Use the MX host provided in the Microsoft 365 Admin Center for your primary custom domain
$sendConnectorTargetSmartHost = "varunagroup-de.mail.protection.outlook.com"

# Create new send connector
# The new send connector is not enabled by default!
# Check additional settings, e.g., MaxMessageSize, before enabling the send connector
New-SendConnector -Name $sendConnectorName -AddressSpaces * -CloudServicesMailEnabled $true `
-Fqdn $sendConnectorFqdn `
-SmartHosts $sendConnectorTargetSmartHost `
-SourceTransportServers $sendConnectorSourceTransportServers`
-MaxMessageSize $sendConnectorMaxMessageSize `
-RequireTLS $true -DNSRoutingEnabled $false `
-TlsAuthLevel CertificateValidation -Enabled:$false