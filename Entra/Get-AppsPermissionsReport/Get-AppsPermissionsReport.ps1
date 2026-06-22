<#
.SYNOPSIS
    Generates a permissions report for all Entra ID (Azure AD) app registrations.

.DESCRIPTION
    Fetches application permissions (AppRoles) and delegated permissions (OAuth2PermissionScopes)
    for all app registrations in the tenant. Permissions can be filtered using predefined sets
    ("Exchange", "SharePoint") or custom permission strings.
    Generates an HTML report and a CSV report stored in a "Reports" subfolder.
    Apps containing EWS (Exchange Web Services) permissions are flagged in a configurable color.
    Delivery options: local filesystem, email, or Microsoft Teams channel webhook.

.PARAMETER PredefinedSets
    Specify predefined permission sets to include: "Exchange", "SharePoint", or both.
    If omitted, all permissions are reported.

.PARAMETER CustomPermissions
    Array of custom permission strings to filter on (e.g. "User.Export.All").
    If omitted together with PredefinedSets, all permissions are reported.

.PARAMETER EwsFlagColor
    HTML color used to highlight apps with EWS permissions. Default: "#FFD700" (gold).

.PARAMETER DeliveryOptions
    Array of delivery methods: "FileSystem", "Email", "Teams". Default: "FileSystem".

.PARAMETER EmailTo
    Recipient email address(es) when DeliveryOptions includes "Email".

.PARAMETER EmailFrom
    Sender email address when DeliveryOptions includes "Email".

.PARAMETER SmtpServer
    SMTP server hostname when DeliveryOptions includes "Email".

.PARAMETER SmtpPort
    SMTP server port. Default: 587.

.PARAMETER SmtpCredential
    PSCredential for SMTP authentication.

.PARAMETER TeamsWebhookUrl
    Incoming Webhook URL for the Teams channel when DeliveryOptions includes "Teams".

.PARAMETER TenantId
    Tenant ID to use when connecting to Microsoft Graph.

.PARAMETER AuthMode
    Authentication mode for Microsoft Graph connection.
    Options: Interactive, AppCertificate, AppSecret.

.PARAMETER ClientId
    Entra ID application (client) ID used for app-based authentication.

.PARAMETER CertificateThumbprint
    Certificate thumbprint from CurrentUser/LocalMachine cert store used with AuthMode AppCertificate.

.PARAMETER ClientSecret
    Client secret (as SecureString) used with AuthMode AppSecret.

.PARAMETER HighlightCategories
    Select permission highlight categories for the HTML report.
    Supports any combination of: HighPrivilege, Exchange, SharePoint.

.PARAMETER OpenHtmlReport
    Opens the generated HTML report in the default browser after script completion.

.EXAMPLE
    .\Get-AppsPermissionsReport.ps1 -PredefinedSets Exchange,SharePoint -DeliveryOptions FileSystem,Email -EmailTo admin@contoso.com -EmailFrom noreply@contoso.com -SmtpServer smtp.contoso.com

.EXAMPLE
    .\Get-AppsPermissionsReport.ps1 -CustomPermissions "User.Export.All","Directory.Read.All" -EwsFlagColor "#FF6347" -DeliveryOptions Teams -TeamsWebhookUrl "https://outlook.office.com/webhook/..."

.EXAMPLE
    .\Get-AppsPermissionsReport.ps1 -PredefinedSets Exchange -OpenHtmlReport

.EXAMPLE
    .\Get-AppsPermissionsReport.ps1 -AuthMode AppCertificate -TenantId "contoso.onmicrosoft.com" -ClientId "00000000-0000-0000-0000-000000000000" -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12"

.EXAMPLE
    $secret = Read-Host "Client Secret" -AsSecureString
    .\Get-AppsPermissionsReport.ps1 -AuthMode AppSecret -TenantId "contoso.onmicrosoft.com" -ClientId "00000000-0000-0000-0000-000000000000" -ClientSecret $secret

.EXAMPLE
    .\Get-AppsPermissionsReport.ps1 -HighlightCategories HighPrivilege,Exchange,SharePoint

.NOTES
    Requires: Microsoft.Graph PowerShell SDK (modules: Microsoft.Graph.Applications, Microsoft.Graph.Identity.DirectoryManagement)
    Install : Install-Module Microsoft.Graph -Scope CurrentUser
#>

[CmdletBinding()]
param (
    [ValidateSet("Exchange", "SharePoint")]
    [string[]]$PredefinedSets = @(),

    [string[]]$CustomPermissions = @(),

    [string]$EwsFlagColor = "#FFD700",

    [ValidateSet("FileSystem", "Email", "Teams")]
    [string[]]$DeliveryOptions = @("FileSystem"),

    [string[]]$EmailTo,
    [string]$EmailFrom,
    [string]$SmtpServer,
    [int]$SmtpPort = 587,
    [System.Management.Automation.PSCredential]$SmtpCredential,

    [string]$TeamsWebhookUrl,

    [string]$TenantId,

    [ValidateSet("Interactive", "AppCertificate", "AppSecret")]
    [string]$AuthMode = "Interactive",

    [string]$ClientId,

    [string]$CertificateThumbprint,

    [System.Security.SecureString]$ClientSecret,

    [ValidateSet("HighPrivilege", "Exchange", "SharePoint")]
    [string[]]$HighlightCategories = @("HighPrivilege"),

    [switch]$OpenHtmlReport
)

#region --- Permission Definitions ---

$PredefinedPermissionSets = @{
    Exchange = @(
        # Mail permissions
        "Mail.Read", "Mail.ReadBasic", "Mail.ReadBasic.All", "Mail.ReadWrite",
        "Mail.Send", "Mail.Send.Shared", "Mail.ReadWrite.Shared",
        # Calendar
        "Calendars.Read", "Calendars.ReadWrite", "Calendars.Read.Shared", "Calendars.ReadWrite.Shared",
        # Contacts
        "Contacts.Read", "Contacts.ReadWrite", "Contacts.Read.Shared", "Contacts.ReadWrite.Shared",
        # Groups / Distribution Lists
        "Group.Read.All", "Group.ReadWrite.All", "GroupMember.Read.All", "GroupMember.ReadWrite.All",
        # MailboxSettings
        "MailboxSettings.Read", "MailboxSettings.ReadWrite",
        # EWS
        "full_access_as_app", "full_access_as_user", "EWS.AccessAsUser.All",
        # Exchange Admin
        "Exchange.ManageAsApp"
    )
    SharePoint = @(
        "Sites.Read.All", "Sites.ReadWrite.All", "Sites.Manage.All", "Sites.FullControl.All",
        "Sites.Selected",
        "Files.Read", "Files.ReadWrite", "Files.Read.All", "Files.ReadWrite.All",
        "Files.ReadWrite.AppFolder", "Files.SelectedOperations.Selected",
        "TermStore.Read.All", "TermStore.ReadWrite.All",
        "User.Read.All", "User.ReadWrite.All",
        "AllSites.Read", "AllSites.Write", "AllSites.Manage", "AllSites.FullControl",
        "MyFiles.Read", "MyFiles.Write"
    )
}

$EwsPermissions = @(
    "full_access_as_app", "full_access_as_user", "EWS.AccessAsUser.All", "Exchange.ManageAsApp"
)

$HighPrivilegePermissions = @(
    "Directory.ReadWrite.All", "Directory.AccessAsUser.All",
    "RoleManagement.ReadWrite.Directory",
    "Application.ReadWrite.All", "Application.ReadWrite.OwnedBy",
    "AppRoleAssignment.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All",
    "Policy.ReadWrite.ConditionalAccess", "Policy.ReadWrite.PermissionGrant",
    "User.ReadWrite.All",
    "Group.ReadWrite.All", "GroupMember.ReadWrite.All",
    "Sites.FullControl.All", "Files.ReadWrite.All",
    "Mail.ReadWrite", "Mail.ReadWrite.Shared", "MailboxSettings.ReadWrite",
    "full_access_as_app", "Exchange.ManageAsApp"
)

$HighPrivilegePermissionSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($perm in $HighPrivilegePermissions) {
    $null = $HighPrivilegePermissionSet.Add($perm)
}

$ExchangePermissionSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($perm in $PredefinedPermissionSets.Exchange) {
    $null = $ExchangePermissionSet.Add($perm)
}

$SharePointPermissionSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($perm in $PredefinedPermissionSets.SharePoint) {
    $null = $SharePointPermissionSet.Add($perm)
}

$HighlightCategorySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($category in $HighlightCategories) {
    $null = $HighlightCategorySet.Add($category)
}

#endregion

#region --- Helper: Build permission filter list ---

function Get-PermissionFilter {
    param (
        [string[]]$Sets,
        [string[]]$Custom
    )
    $filter = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($set in $Sets) {
        foreach ($perm in $PredefinedPermissionSets[$set]) {
            $null = $filter.Add($perm)
        }
    }
    foreach ($perm in $Custom) {
        $null = $filter.Add($perm)
    }
    return $filter
}

#endregion

#region --- Connect to Microsoft Graph ---

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    switch ($AuthMode) {
        "Interactive" {
            $connectParams = @{
                Scopes      = @("Application.Read.All", "Directory.Read.All")
                NoWelcome   = $true
                ErrorAction = 'Stop'
            }
            if ($TenantId) {
                $connectParams.TenantId = $TenantId
            }
            Connect-MgGraph @connectParams
        }

        "AppCertificate" {
            if (-not $TenantId) {
                throw "AuthMode 'AppCertificate' requires -TenantId."
            }
            if (-not $ClientId) {
                throw "AuthMode 'AppCertificate' requires -ClientId."
            }
            if (-not $CertificateThumbprint) {
                throw "AuthMode 'AppCertificate' requires -CertificateThumbprint."
            }

            $connectParams = @{
                TenantId             = $TenantId
                ClientId             = $ClientId
                CertificateThumbprint = $CertificateThumbprint
                NoWelcome            = $true
                ErrorAction          = 'Stop'
            }
            Connect-MgGraph @connectParams
        }

        "AppSecret" {
            if (-not $TenantId) {
                throw "AuthMode 'AppSecret' requires -TenantId."
            }
            if (-not $ClientId) {
                throw "AuthMode 'AppSecret' requires -ClientId."
            }
            if (-not $ClientSecret) {
                throw "AuthMode 'AppSecret' requires -ClientSecret."
            }

            $clientSecretCredential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
            $connectParams = @{
                TenantId              = $TenantId
                ClientSecretCredential = $clientSecretCredential
                NoWelcome             = $true
                ErrorAction           = 'Stop'
            }
            Connect-MgGraph @connectParams
        }
    }
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

#endregion

#region --- Fetch Tenant Info ---

Write-Host "Fetching tenant information..." -ForegroundColor Cyan
try {
    $tenantDetails = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    $tenantDisplayName = $tenantDetails.DisplayName
    if (-not $tenantDisplayName) {
        $tenantDisplayName = ($tenantDetails.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1).Name
    }
    if (-not $tenantDisplayName) {
        $tenantDisplayName = "UnknownTenant"
    }
    $tenantFileName = ($tenantDisplayName -replace '[\\/:*?"<>|]', '_').Trim()
    if (-not $tenantFileName) {
        $tenantFileName = "UnknownTenant"
    }
}
catch {
    Write-Warning "Could not retrieve tenant name. Using 'UnknownTenant'."
    $tenantDisplayName = "UnknownTenant"
    $tenantFileName = "UnknownTenant"
}
Write-Host "Tenant: $tenantDisplayName" -ForegroundColor Green

#endregion

#region --- Fetch Service Principals (resource apps) for permission name resolution ---

Write-Host "Fetching service principals for permission name resolution..." -ForegroundColor Cyan
$allServicePrincipals = Get-MgServicePrincipal -All -Property "AppId,DisplayName,AppRoles,Oauth2PermissionScopes" -ErrorAction SilentlyContinue

# Build lookup: AppId -> ServicePrincipal
$spLookup = @{}
foreach ($sp in $allServicePrincipals) {
    if (-not $spLookup.ContainsKey($sp.AppId)) {
        $spLookup[$sp.AppId] = $sp
    }
}

function Resolve-PermissionName {
    param(
        [string]$ResourceAppId,
        [string]$PermissionId,
        [ValidateSet("Role","Scope")]
        [string]$Type
    )
    $sp = $spLookup[$ResourceAppId]
    if (-not $sp) { return $PermissionId }
    if ($Type -eq "Role") {
        $match = $sp.AppRoles | Where-Object { $_.Id -eq $PermissionId } | Select-Object -First 1
        if ($match) { return $match.Value }
    } else {
        $match = $sp.Oauth2PermissionScopes | Where-Object { $_.Id -eq $PermissionId } | Select-Object -First 1
        if ($match) { return $match.Value }
    }
    return $PermissionId
}

function Resolve-ResourceName {
    param([string]$ResourceAppId)
    $sp = $spLookup[$ResourceAppId]
    if ($sp) { return $sp.DisplayName }
    return $ResourceAppId
}

#endregion

#region --- Fetch App Registrations ---

Write-Host "Fetching app registrations..." -ForegroundColor Cyan
$apps = Get-MgApplication -All -Property "Id,AppId,DisplayName,RequiredResourceAccess,SignInAudience,CreatedDateTime" -ErrorAction Stop
Write-Host "Found $($apps.Count) app registrations." -ForegroundColor Green

#endregion

#region --- Build permission filter ---

$permissionFilter = Get-PermissionFilter -Sets $PredefinedSets -Custom $CustomPermissions
$useFilter        = $permissionFilter.Count -gt 0
Write-Host "Permission filter active: $useFilter$(if ($useFilter) { " ($($permissionFilter.Count) permissions)" })" -ForegroundColor Cyan

#endregion

#region --- Process Apps ---

Write-Host "Processing app permissions..." -ForegroundColor Cyan

$reportData = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($app in ($apps | Sort-Object DisplayName)) {

    $appPermissions = [System.Collections.Generic.List[PSCustomObject]]::new()
    $hasEws         = $false

    foreach ($resourceAccess in $app.RequiredResourceAccess) {
        $resourceName = Resolve-ResourceName -ResourceAppId $resourceAccess.ResourceAppId

        foreach ($access in $resourceAccess.ResourceAccess) {
            $permType = switch ($access.Type) {
                "Role"  { "Application" }
                "Scope" { "Delegated" }
                default { $access.Type }
            }
            $resolvedName = Resolve-PermissionName -ResourceAppId $resourceAccess.ResourceAppId `
                                                    -PermissionId $access.Id `
                                                    -Type $access.Type

            # Apply filter
            if ($useFilter -and -not $permissionFilter.Contains($resolvedName)) { continue }

            if ($EwsPermissions -contains $resolvedName) { $hasEws = $true }
            $isHighPrivilege = $HighPrivilegePermissionSet.Contains($resolvedName)
            $isExchangePermission = $ExchangePermissionSet.Contains($resolvedName)
            $isSharePointPermission = $SharePointPermissionSet.Contains($resolvedName)

            $appPermissions.Add([PSCustomObject]@{
                AppName        = $app.DisplayName
                AppId          = $app.AppId
                ResourceName   = $resourceName
                Permission     = $resolvedName
                PermissionType = $permType
                IsHighPrivilege = $isHighPrivilege
                IsExchangePermission = $isExchangePermission
                IsSharePointPermission = $isSharePointPermission
                HasEWS         = $false  # set per-app below
            })
        }
    }

    # If filtering is active and this app has no matching permissions, skip it
    if ($useFilter -and $appPermissions.Count -eq 0) { continue }

    # Update HasEWS flag on all rows for this app
    foreach ($row in $appPermissions) { $row.HasEWS = $hasEws }

    $reportData.AddRange($appPermissions)
}

Write-Host "Report data rows: $($reportData.Count)" -ForegroundColor Green

#endregion

#region --- Prepare output folder and filenames ---

$scriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportsDir  = Join-Path $scriptRoot "Reports"
if (-not (Test-Path $reportsDir)) {
    New-Item -ItemType Directory -Path $reportsDir | Out-Null
}

$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$baseName    = "${tenantFileName}_${timestamp}"
$htmlPath    = Join-Path $reportsDir "${baseName}.html"
$csvPath     = Join-Path $reportsDir "${baseName}.csv"

#endregion

#region --- Generate CSV ---

Write-Host "Generating CSV report..." -ForegroundColor Cyan
$reportData | Select-Object AppName, AppId, ResourceName, Permission, PermissionType, IsHighPrivilege, IsExchangePermission, IsSharePointPermission, HasEWS |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV saved: $csvPath" -ForegroundColor Green

#endregion

#region --- Generate HTML ---

Write-Host "Generating HTML report..." -ForegroundColor Cyan

# Group rows by app for the HTML table
$groupedApps = $reportData | Group-Object AppName | Sort-Object Name
$highlightHighPrivilege = $HighlightCategorySet.Contains("HighPrivilege")
$highlightExchange = $HighlightCategorySet.Contains("Exchange")
$highlightSharePoint = $HighlightCategorySet.Contains("SharePoint")

$htmlRows = [System.Text.StringBuilder]::new()
$isFirstAppGroup = $true

foreach ($group in $groupedApps) {
    $isEws     = ($group.Group | Select-Object -First 1).HasEWS
    $appCellStyle = if ($isEws) { " style=`"background-color:$EwsFlagColor;font-weight:600;`"" } else { "" }
    $rowCount  = $group.Group.Count
    $firstRow  = $true
    $ewsBadge  = if ($isEws) { " <span class='ews-badge'>EWS</span>" } else { "" }

    foreach ($row in ($group.Group | Sort-Object ResourceName, Permission)) {
        if ($firstRow) {
            $rowSeparatorClass = if ($isFirstAppGroup) { "" } else { " class='app-separator'" }
            $null = $htmlRows.Append("<tr$rowSeparatorClass>")
            $null = $htmlRows.Append("<td rowspan=`"$rowCount`"$appCellStyle>$([System.Web.HttpUtility]::HtmlEncode($row.AppName))$ewsBadge</td>")
            $null = $htmlRows.Append("<td rowspan=`"$rowCount`"><code>$([System.Web.HttpUtility]::HtmlEncode($row.AppId))</code></td>")
            $firstRow = $false
        } else {
            $null = $htmlRows.Append("<tr>")
        }
        $null = $htmlRows.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($row.ResourceName))</td>")
        $permissionBackgroundClass = ""
        $permissionBadges = [System.Text.StringBuilder]::new()

        if ($highlightHighPrivilege -and $row.IsHighPrivilege) {
            $permissionBackgroundClass = "high-priv-perm"
            $null = $permissionBadges.Append(" <span class='high-priv-badge'>HIGH</span>")
        } elseif ($highlightExchange -and $row.IsExchangePermission) {
            $permissionBackgroundClass = "exchange-perm"
        } elseif ($highlightSharePoint -and $row.IsSharePointPermission) {
            $permissionBackgroundClass = "sharepoint-perm"
        }

        if ($highlightExchange -and $row.IsExchangePermission) {
            $null = $permissionBadges.Append(" <span class='exchange-badge'>EXCHANGE</span>")
        }
        if ($highlightSharePoint -and $row.IsSharePointPermission) {
            $null = $permissionBadges.Append(" <span class='sharepoint-badge'>SHAREPOINT</span>")
        }

        $permissionCellClass = if ($permissionBackgroundClass) { " class='$permissionBackgroundClass'" } else { "" }
        $null = $htmlRows.Append("<td$permissionCellClass><strong>$([System.Web.HttpUtility]::HtmlEncode($row.Permission))</strong>$($permissionBadges.ToString())</td>")
        $null = $htmlRows.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($row.PermissionType))</td>")
        $null = $htmlRows.Append("</tr>`n")
    }

    $isFirstAppGroup = $false
}

$filterDesc = if ($useFilter) {
    "Predefined Sets: <strong>$(($PredefinedSets -join ', ') -replace '^$','none')</strong> &nbsp;|&nbsp; Custom: <strong>$(($CustomPermissions -join ', ') -replace '^$','none')</strong>"
} else {
    "No filter applied &mdash; showing <strong>all</strong> permissions."
}

$generatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$totalAppCount = @($reportData | Select-Object -ExpandProperty AppId -Unique).Count
$ewsAppCount   = @($reportData | Where-Object HasEWS | Select-Object -ExpandProperty AppId -Unique).Count

$legendItems = [System.Text.StringBuilder]::new()
$null = $legendItems.Append("<div class='legend-item'><div class='legend-box'></div><span>App contains EWS permission(s)</span></div>")
if ($highlightHighPrivilege) {
    $null = $legendItems.Append("<div class='legend-item'><div class='legend-box high-priv-box'></div><span>High-privilege permission</span></div>")
}
if ($highlightExchange) {
    $null = $legendItems.Append("<div class='legend-item'><div class='legend-box exchange-box'></div><span>Exchange permission</span></div>")
}
if ($highlightSharePoint) {
    $null = $legendItems.Append("<div class='legend-item'><div class='legend-box sharepoint-box'></div><span>SharePoint permission</span></div>")
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>App Permissions Report &ndash; $tenantDisplayName</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; background: #f4f6f9; color: #222; margin: 0; padding: 20px; }
  h1   { color: #0078d4; margin-bottom: 4px; }
  .meta { color: #555; margin-bottom: 16px; font-size: 12px; }
  .filter-info { background: #e8f0fe; border-left: 4px solid #0078d4; padding: 8px 12px; margin-bottom: 16px; border-radius: 4px; }
    .legend { display: flex; align-items: center; gap: 16px; margin-bottom: 12px; font-size: 12px; flex-wrap: wrap; }
    .legend-item { display: inline-flex; align-items: center; gap: 8px; }
  .legend-box { width: 20px; height: 20px; border: 1px solid #aaa; border-radius: 3px; background: $EwsFlagColor; }
    .high-priv-box { background: #a80000; }
        .exchange-box { background: #0a64c5; }
        .sharepoint-box { background: #107c10; }
  table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
  th { background: #0078d4; color: #fff; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
  td { padding: 8px 12px; border-bottom: 1px solid #e5e9f0; vertical-align: top; }
    tr.app-separator td { border-top: 2px solid #b9c7d8; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { filter: brightness(0.96); }
  code { font-size: 11px; background: #f0f0f0; padding: 1px 4px; border-radius: 3px; }
  .ews-badge { background: #c50000; color: #fff; font-size: 10px; font-weight: bold; padding: 2px 6px; border-radius: 10px; }
    .high-priv-perm { background: #ffe2e2; }
    .high-priv-badge { background: #a80000; color: #fff; font-size: 10px; font-weight: bold; padding: 2px 6px; border-radius: 10px; margin-left: 6px; }
        .exchange-perm { background: #e8f3ff; }
        .exchange-badge { background: #0a64c5; color: #fff; font-size: 10px; font-weight: bold; padding: 2px 6px; border-radius: 10px; margin-left: 6px; }
        .sharepoint-perm { background: #e8f7ec; }
        .sharepoint-badge { background: #107c10; color: #fff; font-size: 10px; font-weight: bold; padding: 2px 6px; border-radius: 10px; margin-left: 6px; }
  input[type=search] { padding: 7px 12px; width: 320px; border: 1px solid #ccc; border-radius: 20px; font-size: 13px; margin-bottom: 12px; outline: none; }
  input[type=search]:focus { border-color: #0078d4; box-shadow: 0 0 0 2px #c7e0f4; }
  .count { color: #555; font-size: 12px; margin-left: 8px; }
</style>
</head>
<body>
<h1>&#128274; App Permissions Report</h1>
<div class="meta">Tenant: <strong>$tenantDisplayName</strong> &nbsp;|&nbsp; Generated: $generatedOn &nbsp;|&nbsp; Apps: <strong>$totalAppCount</strong> &nbsp;|&nbsp; EWS Apps: <strong>$ewsAppCount</strong> &nbsp;|&nbsp; Visible rows: <span id="visibleRowCount">$($reportData.Count)</span></div>
<div class="filter-info">$filterDesc</div>
<div class="legend">
    $($legendItems.ToString())
</div>
<input type="search" id="searchBox" placeholder="Search apps, permissions&#x2026;" oninput="filterTable()" />
<table id="reportTable">
<thead>
  <tr>
    <th>App Name</th>
    <th>App ID</th>
    <th>Resource</th>
    <th>Permission</th>
    <th>Type</th>
  </tr>
</thead>
<tbody id="tableBody">
$($htmlRows.ToString())
</tbody>
</table>
<script>
function filterTable() {
    var q = document.getElementById('searchBox').value.toLowerCase();
    var rows = document.getElementById('tableBody').querySelectorAll('tr');
    var visible = 0;
    rows.forEach(function(row) {
        var text = row.innerText.toLowerCase();
        var show = text.indexOf(q) > -1;
        row.style.display = show ? '' : 'none';
        if (show) visible++;
    });
    document.getElementById('visibleRowCount').innerText = visible;
}
</script>
</body>
</html>
"@

$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML saved: $htmlPath" -ForegroundColor Green

#endregion

#region --- Delivery ---

foreach ($delivery in $DeliveryOptions) {

    switch ($delivery) {

        "FileSystem" {
            # Already saved above
            Write-Host "FileSystem delivery: reports saved in $reportsDir" -ForegroundColor Green
        }

        "Email" {
            Write-Host "Sending report via email..." -ForegroundColor Cyan
            if (-not $EmailTo -or -not $EmailFrom -or -not $SmtpServer) {
                Write-Warning "Email delivery requires -EmailTo, -EmailFrom, and -SmtpServer parameters. Skipping."
                continue
            }
            try {
                $mailParams = @{
                    To          = $EmailTo
                    From        = $EmailFrom
                    Subject     = "App Permissions Report - $tenantDisplayName - $timestamp"
                    Body        = "Please find the App Permissions Report attached.<br><br>Tenant: <b>$tenantDisplayName</b><br>Generated: $generatedOn"
                    BodyAsHtml  = $true
                    Attachments = @($htmlPath, $csvPath)
                    SmtpServer  = $SmtpServer
                    Port        = $SmtpPort
                    UseSsl      = $true
                    
                }
                if ($SmtpCredential) { $mailParams.Credential = $SmtpCredential }
                Send-MailMessage @mailParams -ErrorAction Stop
                Write-Host "Email sent to: $($EmailTo -join ', ')" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to send email: $_"
            }
        }

        "Teams" {
            Write-Host "Sending notification to Teams channel..." -ForegroundColor Cyan
            if (-not $TeamsWebhookUrl) {
                Write-Warning "Teams delivery requires -TeamsWebhookUrl parameter. Skipping."
                continue
            }
            try {
                $ewsApps = ($reportData | Where-Object HasEWS | Select-Object -ExpandProperty AppName -Unique | Sort-Object) -join ", "
                $teamsBody = @{
                    "@type"      = "MessageCard"
                    "@context"   = "https://schema.org/extensions"
                    summary      = "App Permissions Report - $tenantDisplayName"
                    themeColor   = "0078D4"
                    title        = "&#128274; App Permissions Report"
                    sections     = @(
                        @{
                            facts = @(
                                @{ name = "Tenant";       value = $tenantDisplayName }
                                @{ name = "Generated";    value = $generatedOn }
                                @{ name = "Total Rows";   value = "$($reportData.Count)" }
                                @{ name = "EWS Apps";     value = if ($ewsApps) { $ewsApps } else { "None" } }
                                @{ name = "CSV Report";   value = $csvPath }
                                @{ name = "HTML Report";  value = $htmlPath }
                            )
                        }
                    )
                } | ConvertTo-Json -Depth 10

                Invoke-RestMethod -Uri $TeamsWebhookUrl -Method Post -Body $teamsBody -ContentType "application/json" -ErrorAction Stop
                Write-Host "Teams notification sent." -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to send Teams notification: $_"
            }
        }
    }
}

#endregion

Write-Host "`nDone! Reports stored in: $reportsDir" -ForegroundColor Cyan

if ($OpenHtmlReport) {
    try {
        if (Test-Path -Path $htmlPath) {
            Write-Host "Opening HTML report in default browser..." -ForegroundColor Cyan
            Start-Process -FilePath $htmlPath
        } else {
            Write-Warning "HTML report not found at expected path: $htmlPath"
        }
    }
    catch {
        Write-Warning "Failed to open HTML report in browser: $_"
    }
}

# Disconnect-MgGraph | Out-Null
