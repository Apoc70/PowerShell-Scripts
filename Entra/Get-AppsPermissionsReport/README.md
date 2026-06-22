# Get-AppsPermissionsReport

`Get-AppsPermissionsReport.ps1` generates a permissions inventory for Entra ID app registrations and exports the result to both HTML and CSV.

The report includes:
- Application and delegated permissions
- Resource API/service principal names
- EWS-related permission detection
- Optional highlighting for high-privilege, Exchange, and SharePoint permissions
- Optional delivery by local files, email, and Teams webhook

## Requirements

- PowerShell 7+ (Windows PowerShell 5.1 also works in many environments)
- Microsoft Graph PowerShell SDK

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

The script uses Microsoft Graph cmdlets for:
- Applications
- Service principals
- Organization details

## Authentication Modes

Use `-AuthMode` to select one of:

- `Interactive` (default)
- `AppCertificate`
- `AppSecret`

### Interactive (default)
Connects with delegated scopes:
- `Application.Read.All`
- `Directory.Read.All`

### AppCertificate
Requires:
- `-TenantId`
- `-ClientId`
- `-CertificateThumbprint`

### AppSecret
Requires:
- `-TenantId`
- `-ClientId`
- `-ClientSecret` (`SecureString`)

## Parameters

- `-PredefinedSets` (`Exchange`, `SharePoint`): filter by built-in permission groups
- `-CustomPermissions`: additional permission names to include in filter
- `-HighlightCategories` (`HighPrivilege`, `Exchange`, `SharePoint`): controls HTML highlighting/badges
- `-EwsFlagColor`: HTML color for app name cell when EWS permissions are detected (default `#FFD700`)
- `-DeliveryOptions` (`FileSystem`, `Email`, `Teams`): output destinations (default `FileSystem`)
- `-OpenHtmlReport`: opens the generated HTML report after completion

Delivery-specific parameters:
- Email: `-EmailTo`, `-EmailFrom`, `-SmtpServer`, optional `-SmtpPort` (default `587`) and `-SmtpCredential`
- Teams: `-TeamsWebhookUrl`

## Output

Reports are saved under the local `Reports` folder next to the script.

Filename format:
- `<TenantName>_<yyyyMMdd_HHmmss>.html`
- `<TenantName>_<yyyyMMdd_HHmmss>.csv`

CSV columns:
- `AppName`
- `AppId`
- `ResourceName`
- `Permission`
- `PermissionType` (`Application` or `Delegated`)
- `IsHighPrivilege`
- `IsExchangePermission`
- `IsSharePointPermission`
- `HasEWS`

The HTML report includes:
- grouped rows per app
- search box filtering
- legend and badges
- summary stats (apps count, EWS app count, visible rows)


## NOTE

The Teams channel notification requires some code changes, because the direct webhook delivery has been deprecated.

## Usage Examples

### 1) All permissions, save to file system

```powershell
.\Get-AppsPermissionsReport.ps1
```

### 2) Filter to Exchange and SharePoint permissions

```powershell
.\Get-AppsPermissionsReport.ps1 -PredefinedSets Exchange,SharePoint
```

### 3) Filter by custom permissions

```powershell
.\Get-AppsPermissionsReport.ps1 -CustomPermissions "User.Export.All","Directory.Read.All"
```

### 4) Highlight all categories and open HTML automatically

```powershell
.\Get-AppsPermissionsReport.ps1 -HighlightCategories HighPrivilege,Exchange,SharePoint -OpenHtmlReport
```

### 5) Send by email

```powershell
.\Get-AppsPermissionsReport.ps1 \
	-PredefinedSets Exchange \
	-DeliveryOptions FileSystem,Email \
	-EmailTo admin@contoso.com \
	-EmailFrom noreply@contoso.com \
	-SmtpServer smtp.contoso.com
```

### 6) Use app-only auth with certificate

```powershell
.\Get-AppsPermissionsReport.ps1 \
	-AuthMode AppCertificate \
	-TenantId "contoso.onmicrosoft.com" \
	-ClientId "00000000-0000-0000-0000-000000000000" \
	-CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
```

### 7) Use app-only auth with client secret

```powershell
$secret = Read-Host "Client Secret" -AsSecureString
.\Get-AppsPermissionsReport.ps1 \
	-AuthMode AppSecret \
	-TenantId "contoso.onmicrosoft.com" \
	-ClientId "00000000-0000-0000-0000-000000000000" \
	-ClientSecret $secret
```

## Notes

- If `-PredefinedSets` and `-CustomPermissions` are both omitted, the script reports all discovered app permissions.
- If email or Teams settings are incomplete, the script skips that delivery path and continues.
- EWS-related permissions currently tracked: `full_access_as_app`, `full_access_as_user`, `EWS.AccessAsUser.All`, `Exchange.ManageAsApp`.
