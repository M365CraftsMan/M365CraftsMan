<#
.SYNOPSIS
    ONE-TIME, interactive SharePoint sharing-exposure report at FILE and FOLDER
    level (anonymous links, org-wide links, external/guest grants) for every
    document library in the tenant. Writes a CSV locally.

.DESCRIPTION
    How it scales: Graph has no tenant-wide "list shared items" API, so every
    library is walked with a delta query (GET /drives/{id}/root/delta, paged).
    With the "Prefer: hierarchicalsharing" header, Graph only returns sharing
    information on items that are the ROOT of a permission hierarchy (items
    with their own sharing/unique permissions). Items that just inherit come back
    without it. So in the default mode (-PermissionCheckMode SharedOnly) the
    per-item /permissions call is made ONLY for those items - not for every file.

    -PermissionCheckMode AllItems calls /permissions on every single item. Use
    that only on one site, to validate SharedOnly against it - tenant-wide it
    will be extremely slow and get throttled.

    Auth: APP-ONLY (same temporary app registration as the root-level script).
        Sites.Read.All          - required
        Organization.Read.All   - verified domains for the external check
    Microsoft's scanning guidance says apps processing permissions "correctly"
    should hold Sites.FullControl.All. If the validation run (see EXAMPLE 1)
    misses a file you know is shared, that's the first thing to add.

    Output:
      <OutputFolder>\SharePointItemSharing_<timestamp>_raw.csv
          written incrementally after every library, so a crash or Ctrl+C
          after 3 hours doesn't lose 3 hours
      <OutputFolder>\SharePointItemSharing_<timestamp>.csv
          final, ranked by risk (written at the end)

    By default only findings (non-OK rows) and errors are written. Inherited
    permissions are skipped (they're reported once, on the item they come from).

.EXAMPLE
    # 1) VALIDATE FIRST: one site where you know a file has an "Anyone" link
    .\Get-SharePointSharingExposure-ItemLevel.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "<app-id>" -SiteUrlFilter "*/sites/Finance"

.EXAMPLE
    # 2) Same site, brute-force mode - results should match run 1
    .\Get-SharePointSharingExposure-ItemLevel.ps1 -TenantId "..." -ClientId "..." -SiteUrlFilter "*/sites/Finance" -PermissionCheckMode AllItems

.EXAMPLE
    # 3) Full tenant incl. OneDrive (run off-peak: US night / weekend)
    .\Get-SharePointSharingExposure-ItemLevel.ps1 -TenantId "..." -ClientId "..." -IncludeOneDrive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$TenantId,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$OutputFolder = (Join-Path $env:USERPROFILE "Downloads"),
    [switch]$IncludeOneDrive,
    [string[]]$SiteUrlFilter,                                          # wildcard(s) - only scan matching sites, e.g. "*/sites/Finance*"
    [string[]]$ExcludedUrlPatterns = @("*/sites/appcatalog*"),
    [string[]]$ExcludedDriveNames  = @("Preservation Hold Library", "Form Templates", "Style Library"),
    [int]$MaxSites = 0,
    [ValidateSet("SharedOnly", "AllItems")][string]$PermissionCheckMode = "SharedOnly",
    [switch]$IncludeOkRows,                                             # also write internal/OK grants (much bigger file)
    [switch]$IncludeInherited,                                          # also write permissions inherited from a parent
    [int]$MaxNamesPerCell = 25
)

$ErrorActionPreference = "Stop"

trap {
    Write-Host ""
    Write-Host "FAILED at script line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.GetType().FullName): $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    Write-Host "PowerShell $($PSVersionTable.PSVersion) | Graph.Authentication $((Get-Module Microsoft.Graph.Authentication).Version)" -ForegroundColor DarkGray
    if ($script:rawFile) { Write-Host "Partial results so far: $script:rawFile" -ForegroundColor Yellow }
    break
}

# ---- Module ------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "Installing Microsoft.Graph.Authentication for current user..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication

# ---- Helpers -----------------------------------------------------------------
function ConvertTo-JoinedNames {
    param([string[]]$Names, [int]$MaxNames)
    $unique = @($Names | Where-Object { $_ } | Sort-Object -Unique)
    if ($unique.Count -eq 0) { return "" }
    if ($unique.Count -le $MaxNames) { return ($unique -join "; ") }
    return ("{0}; +{1} more" -f (($unique | Select-Object -First $MaxNames) -join "; "), ($unique.Count - $MaxNames))
}

function Test-HasKey {
    param($Object, [string]$Key)
    if ($null -eq $Object) { return $false }
    if ($Object -is [System.Collections.IDictionary]) { return $Object.Contains($Key) }
    return $null -ne $Object.PSObject.Properties[$Key]
}

function Invoke-GraphRequestWithRetry {
    param([string]$Uri, [hashtable]$Headers = @{}, [int]$MaxRetries = 8)
    $attempt = 0
    while ($true) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            if ($msg -match '429|TooManyRequests|503|504|ServiceUnavailable|GatewayTimeout' -and $attempt -le $MaxRetries) {
                # Honor Retry-After when Graph sends it (SharePoint blocks apps that ignore it)
                $wait = $null
                try { $wait = [int]$_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds } catch { }
                if (-not $wait -or $wait -le 0) { $wait = [int][math]::Min(120, [math]::Pow(2, $attempt)) }
                Write-Warning "Throttled (attempt $attempt/$MaxRetries) - pausing $wait s"
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphCollection {
    param([string]$Uri, [hashtable]$Headers = @{})
    $results = New-Object System.Collections.ArrayList
    $nextUri = $Uri
    while ($nextUri) {
        $resp = Invoke-GraphRequestWithRetry -Uri $nextUri -Headers $Headers
        foreach ($item in @($resp.value)) { if ($null -ne $item) { [void]$results.Add($item) } }
        $nextUri = $resp.'@odata.nextLink'
    }
    return $results.ToArray()
}

function New-Row {
    param(
        $Site, [string]$DriveName = "", [string]$ItemType = "", [string]$ItemName = "", [string]$ItemUrl = "",
        [string]$PermissionType = "", [string]$GrantedTo = "", [string]$GrantedToGroups = "",
        [string]$LinkScope = "", [string]$LinkType = "", [string]$Roles = "",
        [bool]$IsExternal = $false, [string]$ExternalGrantees = "", [bool]$HasPassword = $false,
        [string]$ExpirationDateTime = "", [string]$InheritedFrom = "",
        [string]$Status = "OK", [string]$Action = "OK", [string]$ErrorMessage = ""
    )
    [PSCustomObject]@{
        RecommendedAction  = $Action
        SiteUrl            = $Site.webUrl
        SiteName           = $Site.displayName
        DriveName          = $DriveName
        ItemType           = $ItemType
        ItemName           = $ItemName
        ItemUrl            = $ItemUrl
        PermissionType     = $PermissionType
        LinkScope          = $LinkScope
        LinkType           = $LinkType
        Roles              = $Roles
        GrantedTo          = $GrantedTo
        GrantedToGroups    = $GrantedToGroups
        IsExternal         = $IsExternal
        ExternalGrantees   = $ExternalGrantees
        HasPassword        = $HasPassword
        ExpirationDateTime = $ExpirationDateTime
        InheritedFrom      = $InheritedFrom
        Status             = $Status
        Error              = $ErrorMessage
    }
}

# ---- Connect (app-only) ---------------------------------------------------------
if ($CertificateThumbprint) {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
else {
    $cred = Get-Credential -UserName $ClientId -Message "Paste the CLIENT SECRET VALUE as the password"
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
    Remove-Variable cred
}
$ctx = Get-MgContext
if ($ctx.AuthType -ne "AppOnly") { throw "Connected as '$($ctx.AuthType)', not AppOnly. getAllSites needs application permissions." }
Write-Host "Connected app-only to tenant $($ctx.TenantId) | roles: $($ctx.Scopes -join ', ')" -ForegroundColor Green
if ($ctx.Scopes -notcontains "Sites.FullControl.All") {
    Write-Host "Note: running without Sites.FullControl.All - validate on a known-shared site first (see EXAMPLE 1)." -ForegroundColor DarkYellow
}

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$timestamp      = Get-Date -Format "yyyy-MM-dd_HHmm"
$script:rawFile = Join-Path $OutputFolder "SharePointItemSharing_${timestamp}_raw.csv"
$finalFile      = Join-Path $OutputFolder "SharePointItemSharing_${timestamp}.csv"

# ---- Verified domains -------------------------------------------------------------
$verifiedDomains = @()
try {
    $org = Invoke-GraphRequestWithRetry -Uri "https://graph.microsoft.com/v1.0/organization?`$select=verifiedDomains"
    $verifiedDomains = @($org.value[0].verifiedDomains | ForEach-Object { $_.name.ToLower() })
    Write-Host "Verified domains: $($verifiedDomains -join ', ')"
}
catch {
    Write-Warning "Could not read verified domains - IsExternal will be false everywhere. $($_.Exception.Message)"
}

# ---- Sites ------------------------------------------------------------------------------
Write-Host "Enumerating sites via /sites/getAllSites ..."
$allSites = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/getAllSites?`$select=id,name,displayName,webUrl,isPersonalSite&`$top=200")
Write-Host "Found $($allSites.Count) site(s) before filtering."

if (-not $IncludeOneDrive) {
    $allSites = @($allSites | Where-Object { -not $_.isPersonalSite -and $_.webUrl -notmatch '-my\.sharepoint\.com/personal/' })
}
foreach ($pattern in $ExcludedUrlPatterns) { $allSites = @($allSites | Where-Object { $_.webUrl -notlike $pattern }) }
if ($SiteUrlFilter) {
    $allSites = @($allSites | Where-Object {
        $url = $_.webUrl
        @($SiteUrlFilter | Where-Object { $url -like $_ }).Count -gt 0
    })
}
if ($MaxSites -gt 0 -and $allSites.Count -gt $MaxSites) { $allSites = @($allSites | Select-Object -First $MaxSites) }
if ($allSites.Count -eq 0) { throw "No sites left after filtering - check -SiteUrlFilter / -IncludeOneDrive." }
Write-Host "Scanning $($allSites.Count) site(s) | mode: $PermissionCheckMode | raw output: $script:rawFile" -ForegroundColor Cyan

# ---- Scan ---------------------------------------------------------------------------------
$deltaHeaders = @{ Prefer = "hierarchicalsharing" }
$stats = @{ Items = 0; PermChecks = 0; Rows = 0; Errors = 0 }
$actionCounts = @{}
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$siteIndex = 0

foreach ($site in $allSites) {
    $siteIndex++
    Write-Progress -Id 1 -Activity "Sites" -Status "$siteIndex / $($allSites.Count): $($site.webUrl)" -PercentComplete (($siteIndex / $allSites.Count) * 100)
    $siteRows = New-Object System.Collections.ArrayList

    try {
        $drives = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,webUrl")
    }
    catch {
        [void]$siteRows.Add((New-Row -Site $site -Status "Error" -Action "Check-Access" -ErrorMessage $_.Exception.Message))
        $drives = @()
    }

    foreach ($drive in $drives) {
        if ($ExcludedDriveNames -contains $drive.name) { continue }

        # 1) Walk the library with delta, keep only items that need a permission check
        $candidates = New-Object System.Collections.ArrayList
        $driveItems = 0
        try {
            $nextUri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/delta"
            while ($nextUri) {
                $page = Invoke-GraphRequestWithRetry -Uri $nextUri -Headers $deltaHeaders
                foreach ($item in @($page.value)) {
                    if ($null -eq $item -or (Test-HasKey $item 'deleted')) { continue }
                    $driveItems++
                    $isRoot = Test-HasKey $item 'root'
                    $needsCheck = ($PermissionCheckMode -eq "AllItems") -or $isRoot -or (Test-HasKey $item 'shared')
                    if ($needsCheck) {
                        $type = if ($isRoot) { "LibraryRoot" } elseif (Test-HasKey $item 'folder') { "Folder" } else { "File" }
                        [void]$candidates.Add([PSCustomObject]@{ Id = $item.id; Name = $item.name; WebUrl = $item.webUrl; Type = $type })
                    }
                }
                $nextUri = $page.'@odata.nextLink'      # the final page carries @odata.deltaLink instead - we stop there
                if ($driveItems -gt 0 -and $driveItems % 5000 -lt 200) {
                    Write-Progress -Id 2 -ParentId 1 -Activity "$($drive.name)" -Status "$driveItems items walked, $($candidates.Count) to check"
                }
            }
        }
        catch {
            [void]$siteRows.Add((New-Row -Site $site -DriveName $drive.name -Status "Error" -Action "Check-Access" -ErrorMessage "Delta walk failed: $($_.Exception.Message)"))
            continue
        }
        $stats.Items += $driveItems

        # 2) Permissions only for the candidates
        $checked = 0
        foreach ($c in $candidates) {
            $checked++
            if ($checked % 50 -eq 0) {
                Write-Progress -Id 2 -ParentId 1 -Activity "$($drive.name)" -Status "Permissions $checked / $($candidates.Count)" -PercentComplete (($checked / $candidates.Count) * 100)
            }
            $stats.PermChecks++
            try {
                $perms = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($c.Id)/permissions")
            }
            catch {
                [void]$siteRows.Add((New-Row -Site $site -DriveName $drive.name -ItemType $c.Type -ItemName $c.Name -ItemUrl $c.WebUrl -Status "Error" -Action "Check-Access" -ErrorMessage $_.Exception.Message))
                continue
            }

            foreach ($perm in $perms) {
                $inheritedFrom = ""
                if ($perm.inheritedFrom -and ($perm.inheritedFrom.id -or $perm.inheritedFrom.path)) {
                    $inheritedFrom = if ($perm.inheritedFrom.path) { $perm.inheritedFrom.path } else { $perm.inheritedFrom.id }
                    if (-not $IncludeInherited) { continue }
                }

                $isLink    = $null -ne $perm.link
                $linkScope = if ($isLink) { [string]$perm.link.scope } else { "" }
                $linkType  = if ($isLink) { [string]$perm.link.type }  else { "" }

                $emails = @()
                if ($perm.grantedToV2.user.email) { $emails += $perm.grantedToV2.user.email }
                foreach ($id in @($perm.grantedToIdentitiesV2)) { if ($id.user.email) { $emails += $id.user.email } }
                if ($perm.invitation.email) { $emails += $perm.invitation.email }
                $emails = @($emails | Where-Object { $_ } | Sort-Object -Unique)

                $groups = @()
                if ($perm.grantedToV2.siteGroup.displayName) { $groups += $perm.grantedToV2.siteGroup.displayName }
                if ($perm.grantedToV2.group.displayName)     { $groups += $perm.grantedToV2.group.displayName }
                foreach ($id in @($perm.grantedToIdentitiesV2)) {
                    if ($id.siteGroup.displayName) { $groups += $id.siteGroup.displayName }
                    if ($id.group.displayName)     { $groups += $id.group.displayName }
                }

                $externalEmails = @()
                if ($verifiedDomains.Count -gt 0) {
                    $externalEmails = @($emails | Where-Object {
                        $parts = $_ -split '@'
                        $parts.Count -ne 2 -or ($verifiedDomains -notcontains $parts[1].ToLower())
                    })
                }
                $isExternal = $externalEmails.Count -gt 0

                $permType =
                    if ($linkScope -eq "anonymous")        { "AnonymousLink" }
                    elseif ($linkScope -eq "organization") { "OrganizationLink" }
                    elseif ($isLink)                       { "OtherLink" }
                    else                                   { "Direct" }

                $action =
                    if ($permType -eq "AnonymousLink")        { "Review-Anonymous-Link" }
                    elseif ($isExternal)                      { "Review-External-Access" }
                    elseif ($permType -eq "OrganizationLink") { "Review-OrgWide-Link" }
                    else                                      { "OK" }

                if ($action -eq "OK" -and -not $IncludeOkRows) { continue }

                [void]$siteRows.Add((New-Row -Site $site -DriveName $drive.name -ItemType $c.Type -ItemName $c.Name -ItemUrl $c.WebUrl `
                    -PermissionType $permType -GrantedTo (ConvertTo-JoinedNames -Names $emails -MaxNames $MaxNamesPerCell) `
                    -GrantedToGroups (ConvertTo-JoinedNames -Names $groups -MaxNames $MaxNamesPerCell) `
                    -LinkScope $linkScope -LinkType $linkType -Roles (@($perm.roles) -join ",") `
                    -IsExternal $isExternal -ExternalGrantees (ConvertTo-JoinedNames -Names $externalEmails -MaxNames $MaxNamesPerCell) `
                    -HasPassword ([bool]$perm.hasPassword) -ExpirationDateTime ([string]$perm.expirationDateTime) `
                    -InheritedFrom $inheritedFrom -Action $action))
            }
        }
        Write-Progress -Id 2 -ParentId 1 -Activity "$($drive.name)" -Completed
        Write-Host ("  {0} | {1}: {2} items, {3} permission checks" -f $site.webUrl, $drive.name, $driveItems, $candidates.Count) -ForegroundColor DarkGray
    }

    # Checkpoint: append this site's rows to the raw CSV
    if ($siteRows.Count -gt 0) {
        $siteRows | Export-Csv -Path $script:rawFile -NoTypeInformation -Encoding UTF8 -Append
        foreach ($r in $siteRows) {
            $stats.Rows++
            if ($r.Status -eq "Error") { $stats.Errors++ }
            if ($actionCounts.ContainsKey($r.RecommendedAction)) { $actionCounts[$r.RecommendedAction]++ } else { $actionCounts[$r.RecommendedAction] = 1 }
        }
    }
}
Write-Progress -Id 1 -Activity "Sites" -Completed
$stopwatch.Stop()

# ---- Final ranked CSV ------------------------------------------------------------------------
if (Test-Path $script:rawFile) {
    $riskRank = @{ "Review-Anonymous-Link" = 0; "Review-External-Access" = 1; "Review-OrgWide-Link" = 2; "Check-Access" = 3; "OK" = 4 }
    Import-Csv $script:rawFile |
        Sort-Object @{ Expression = { $riskRank[$_.RecommendedAction] } }, SiteUrl, DriveName, ItemUrl |
        Export-Csv -Path $finalFile -NoTypeInformation -Encoding UTF8
}

function Get-Count([string]$k) { if ($actionCounts.ContainsKey($k)) { $actionCounts[$k] } else { 0 } }
Write-Host ""
Write-Host "===== SharePoint Sharing Exposure (file + folder level) =====" -ForegroundColor Cyan
Write-Host ("Sites scanned        : {0}" -f $allSites.Count)
Write-Host ("Items walked         : {0}" -f $stats.Items)
Write-Host ("Permission checks    : {0}" -f $stats.PermChecks)
Write-Host ("Anonymous links      : {0}" -f (Get-Count "Review-Anonymous-Link")) -ForegroundColor Red
Write-Host ("External grants      : {0}" -f (Get-Count "Review-External-Access")) -ForegroundColor Yellow
Write-Host ("Org-wide links       : {0}" -f (Get-Count "Review-OrgWide-Link"))
Write-Host ("Errors (check access): {0}" -f $stats.Errors)
Write-Host ("Run time             : {0:hh\:mm\:ss}" -f $stopwatch.Elapsed)
if (Test-Path $finalFile) { Write-Host ("CSV                  : {0}" -f $finalFile) -ForegroundColor Green }
else { Write-Host "No findings - no CSV written." -ForegroundColor Green }

Disconnect-MgGraph | Out-Null
