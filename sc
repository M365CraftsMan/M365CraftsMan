<#
.SYNOPSIS
    ONE-TIME, interactive run of the SharePoint sharing-exposure report
    (anonymous links, org-wide links, external/guest grants on every document
    library ROOT of every site in the tenant). Writes a CSV locally - no email,
    no Automation Account, no managed identity.

.DESCRIPTION
    Standalone copy of the Runbook 9 logic, adapted to run from a workstation
    (Cloud PC). Not part of the maintenance-automation runbook set.

    Auth is APP-ONLY (client secret or certificate) on a temporary app
    registration. Delegated sign-in will NOT work for this report: Microsoft
    Graph's /sites/getAllSites supports application permissions only, so a
    delegated session cannot enumerate every site in the tenant.

    Temporary app registration needs (APPLICATION permissions, admin consent):
        Sites.Read.All          - enumerate sites, drives, root permissions
        Organization.Read.All   - verified domains for the external check
    Delete the app registration when you're done.

    Differences vs Runbook 9:
      - App-only auth via TenantId/ClientId + secret (prompted, never stored) or cert thumbprint
      - CSV saved to -OutputFolder instead of emailed
      - IsExternal fixed: row is external if ANY grantee is outside the tenant's
        verified domains (Runbook 9 marked a row internal if ANY grantee was
        internal, so a link redeemed by 1 internal + 1 external user was missed)
      - Site-group / group grants now show the group name instead of a blank cell
      - Error rows keep the actual error message instead of always "AccessDenied"
      - OneDrive inclusion is a switch (-IncludeOneDrive)

    SCOPE GAP (same as Runbook 9): only the ROOT of each document library is
    checked. Links/grants on individual files and folders are NOT enumerated.
    A clean result here does not mean there are no anonymous links in the tenant.

.EXAMPLE
    .\Get-SharePointSharingExposure-OneTime.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "00000000-0000-0000-0000-000000000000"
    (prompts for the client secret)

.EXAMPLE
    .\Get-SharePointSharingExposure-OneTime.ps1 -TenantId "<tenant-guid>" -ClientId "<app-id>" -CertificateThumbprint "ABC123..." -IncludeOneDrive -MaxSites 25
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$TenantId,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [string]$CertificateThumbprint,                                     # if omitted, you'll be prompted for a client secret
    [string]$OutputFolder        = (Join-Path $env:USERPROFILE "Downloads"),
    [switch]$IncludeOneDrive,                                           # personal OneDrive sites - off by default
    [string[]]$ExcludedUrlPatterns = @("*/sites/appcatalog*"),
    [int]$MaxSites               = 0,                                   # 0 = no limit; try 25 first to time a run
    [int]$MaxNamesPerCell        = 25
)

$ErrorActionPreference = "Stop"

# ---- Module ------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "Microsoft.Graph.Authentication not found - installing for current user..." -ForegroundColor Yellow
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

function Invoke-GraphRequestWithRetry {
    param([string]$Uri, [string]$Method = "GET", [int]$MaxRetries = 5)
    $attempt = 0
    while ($true) {
        try {
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -ErrorAction Stop
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            if ($msg -match '429|TooManyRequests|503|504|ServiceUnavailable|GatewayTimeout' -and $attempt -le $MaxRetries) {
                $backoff = [math]::Min(60, [math]::Pow(2, $attempt))
                Write-Warning "Throttled/transient (attempt $attempt/$MaxRetries) - waiting $backoff s"
                Start-Sleep -Seconds $backoff
                continue
            }
            throw
        }
    }
}

function Get-GraphCollection {
    param([string]$Uri)
    $results = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri
    while ($nextUri) {
        $resp = Invoke-GraphRequestWithRetry -Uri $nextUri
        if ($resp.value) { $results.AddRange([object[]]$resp.value) }
        $nextUri = $resp.'@odata.nextLink'
    }
    return ,$results
}

function New-ResultRow {
    param($Site, [string]$DriveName = "", [string]$Status, [string]$Action, [string]$ErrorMessage = "")
    [PSCustomObject]@{
        SiteUrl = $Site.webUrl; SiteName = $Site.displayName; DriveName = $DriveName
        PermissionType = ""; GrantedTo = ""; GrantedToGroups = ""; LinkScope = ""; LinkType = ""; Roles = ""
        IsExternal = $false; ExternalGrantees = ""; HasPassword = $false; ExpirationDateTime = ""
        Status = $Status; RecommendedAction = $Action; Error = $ErrorMessage
    }
}

# ---- Connect (app-only) -------------------------------------------------------
if ($CertificateThumbprint) {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
else {
    $cred = Get-Credential -UserName $ClientId -Message "Paste the CLIENT SECRET VALUE as the password"
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
    Remove-Variable cred
}

$ctx = Get-MgContext
if ($ctx.AuthType -ne "AppOnly") {
    throw "Connected as '$($ctx.AuthType)', not AppOnly. /sites/getAllSites only works with application permissions."
}
if ($ctx.Scopes -notcontains "Sites.Read.All" -and $ctx.Scopes -notcontains "Sites.ReadWrite.All" -and $ctx.Scopes -notcontains "Sites.FullControl.All") {
    Write-Warning "Token roles: $($ctx.Scopes -join ', '). Sites.Read.All not visible - admin consent may be missing or still propagating."
}
Write-Host "Connected app-only to tenant $($ctx.TenantId)" -ForegroundColor Green

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$outFile   = Join-Path $OutputFolder "SharePointSharingExposure_$timestamp.csv"

# ---- Verified domains -----------------------------------------------------------
$verifiedDomains = @()
try {
    $org = Invoke-GraphRequestWithRetry -Uri "https://graph.microsoft.com/v1.0/organization?`$select=verifiedDomains"
    $verifiedDomains = @($org.value[0].verifiedDomains | ForEach-Object { $_.name.ToLower() })
    Write-Host "Verified domains: $($verifiedDomains -join ', ')"
}
catch {
    Write-Warning "Could not read verified domains (Organization.Read.All missing?) - IsExternal will be blank-false for every row. $($_.Exception.Message)"
}

# ---- Enumerate sites ----------------------------------------------------------------
Write-Host "Enumerating all sites via /sites/getAllSites ..."
try {
    $allSites = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/getAllSites?`$select=id,name,displayName,webUrl,isPersonalSite&`$top=200"
}
catch {
    throw "getAllSites failed - confirm Sites.Read.All (APPLICATION) is granted with admin consent. $($_.Exception.Message)"
}
$allSites = @($allSites)
Write-Host "Found $($allSites.Count) site(s) before filtering."

if (-not $IncludeOneDrive) {
    $allSites = @($allSites | Where-Object { -not $_.isPersonalSite -and $_.webUrl -notmatch '-my\.sharepoint\.com/personal/' })
}
foreach ($pattern in $ExcludedUrlPatterns) {
    $allSites = @($allSites | Where-Object { $_.webUrl -notlike $pattern })
}
if ($MaxSites -gt 0 -and $allSites.Count -gt $MaxSites) {
    Write-Host "MaxSites=$MaxSites - trimming from $($allSites.Count)." -ForegroundColor Yellow
    $allSites = @($allSites | Select-Object -First $MaxSites)
}
Write-Host "Processing $($allSites.Count) site(s)."

# ---- Walk drives + root permissions ---------------------------------------------------
$results   = New-Object System.Collections.Generic.List[object]
$processed = 0
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($site in $allSites) {
    $processed++
    Write-Progress -Activity "SharePoint sharing exposure" -Status "$processed / $($allSites.Count): $($site.webUrl)" -PercentComplete (($processed / [math]::Max(1, $allSites.Count)) * 100)

    try {
        $drives = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,webUrl")
    }
    catch {
        $results.Add((New-ResultRow -Site $site -Status "Error" -Action "Check-Access" -ErrorMessage $_.Exception.Message))
        continue
    }

    if ($drives.Count -eq 0) {
        $results.Add((New-ResultRow -Site $site -Status "NoDrives" -Action "OK"))
        continue
    }

    foreach ($drive in $drives) {
        try {
            $perms = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives/$($drive.id)/root/permissions")
        }
        catch {
            $results.Add((New-ResultRow -Site $site -DriveName $drive.name -Status "Error" -Action "Check-Access" -ErrorMessage $_.Exception.Message))
            continue
        }

        foreach ($perm in $perms) {
            $isLink    = $null -ne $perm.link
            $linkScope = if ($isLink) { $perm.link.scope } else { "" }
            $linkType  = if ($isLink) { $perm.link.type }  else { "" }

            # Users: direct grant, redeemed link users, pending invitation
            $emails = @()
            if ($perm.grantedToV2.user.email) { $emails += $perm.grantedToV2.user.email }
            foreach ($id in @($perm.grantedToIdentitiesV2)) {
                if ($id.user.email) { $emails += $id.user.email }
            }
            if ($perm.invitation.email) { $emails += $perm.invitation.email }
            $emails = @($emails | Where-Object { $_ } | Sort-Object -Unique)

            # Groups / SharePoint site groups (Owners/Members/Visitors show up here)
            $groups = @()
            if ($perm.grantedToV2.siteGroup.displayName) { $groups += $perm.grantedToV2.siteGroup.displayName }
            if ($perm.grantedToV2.group.displayName)     { $groups += $perm.grantedToV2.group.displayName }
            foreach ($id in @($perm.grantedToIdentitiesV2)) {
                if ($id.siteGroup.displayName) { $groups += $id.siteGroup.displayName }
                if ($id.group.displayName)     { $groups += $id.group.displayName }
            }

            # External = ANY grantee whose domain isn't a verified tenant domain
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

            $results.Add([PSCustomObject]@{
                SiteUrl            = $site.webUrl
                SiteName           = $site.displayName
                DriveName          = $drive.name
                PermissionType     = $permType
                GrantedTo          = ConvertTo-JoinedNames -Names $emails -MaxNames $MaxNamesPerCell
                GrantedToGroups    = ConvertTo-JoinedNames -Names $groups -MaxNames $MaxNamesPerCell
                LinkScope          = $linkScope
                LinkType           = $linkType
                Roles              = (@($perm.roles) -join ",")
                IsExternal         = $isExternal
                ExternalGrantees   = ConvertTo-JoinedNames -Names $externalEmails -MaxNames $MaxNamesPerCell
                HasPassword        = [bool]$perm.hasPassword
                ExpirationDateTime = $perm.expirationDateTime
                Status             = "OK"
                RecommendedAction  = $action
                Error              = ""
            })
        }
    }
}
Write-Progress -Activity "SharePoint sharing exposure" -Completed
$stopwatch.Stop()

# ---- Export, ranked by risk -------------------------------------------------------------
$riskRank = @{
    "Review-Anonymous-Link"  = 0
    "Review-External-Access" = 1
    "Review-OrgWide-Link"    = 2
    "Check-Access"           = 3
    "OK"                     = 4
}
$export = $results | Sort-Object @{ Expression = { $riskRank[$_.RecommendedAction] } }, SiteUrl, DriveName
$export | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

# ---- Summary ------------------------------------------------------------------------------
$count = { param($a) @($results | Where-Object { $_.RecommendedAction -eq $a }).Count }
Write-Host ""
Write-Host "===== SharePoint Sharing Exposure (drive ROOT only) =====" -ForegroundColor Cyan
Write-Host ("Sites processed      : {0}" -f $allSites.Count)
Write-Host ("Rows                 : {0}" -f $results.Count)
Write-Host ("Anonymous links      : {0}" -f (& $count "Review-Anonymous-Link")) -ForegroundColor Red
Write-Host ("External grants      : {0}" -f (& $count "Review-External-Access")) -ForegroundColor Yellow
Write-Host ("Org-wide links       : {0}" -f (& $count "Review-OrgWide-Link"))
Write-Host ("Errors (check access): {0}" -f (& $count "Check-Access"))
Write-Host ("Run time             : {0:hh\:mm\:ss}" -f $stopwatch.Elapsed)
Write-Host ("CSV                  : {0}" -f $outFile) -ForegroundColor Green
Write-Host "Reminder: file/folder-level links are NOT in this report." -ForegroundColor DarkYellow

Disconnect-MgGraph | Out-Null
