# This script runs the Get-AppsPermissionsReport.ps1 script with the specified parameters to generate an application permissions report for a tenant.
# Note: You need to replace the $tenantId, $clientId, and $CertificateThumbprint variables with your own values before running the script.

$tenantId = "" # Your tenant ID here
$clientId = "" # Your app client ID here
$CertificateThumbprint = "" # Your certificate thumbprint here

# Run the Get-AppsPermissionsReport.ps1 script with the specified parameters
.\Get-AppsPermissionsReport.ps1 -AuthMode AppCertificate -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -openhtmlreport -HighlightCategories HighPrivilege,Exchange,SharePoint