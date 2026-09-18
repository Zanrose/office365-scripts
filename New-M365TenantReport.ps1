#Requires -Version 5.1
<#
.SYNOPSIS
    Generates a general-purpose Microsoft 365 tenant report (identity, security,
    licensing, mail, and collaboration inventory) as a multi-tab Excel workbook.

.DESCRIPTION
    Connects to Microsoft Graph (PowerShell SDK) and inventories the tenant
    across four workloads:

        * Identity & licensing  - users, guests, license SKUs (+ unused seats),
                                   per-user licenses, MFA / auth-method
                                   registration, privileged (admin) role holders
        * Exchange / mailboxes  - mailbox usage (size, item count, archive),
                                   mailbox-level and inbox-rule mail forwarding
        * SharePoint & OneDrive - site storage, per-user OneDrive usage,
                                   tenant-wide external sharing settings
        * Teams & Groups        - M365 groups, Teams, owner/member counts

    Useful on its own for tenant documentation, security/hygiene reviews, and
    license audits, and equally as the source-tenant inventory for a
    tenant-to-tenant migration.

    Each data set is written to its own worksheet, with a "Summary" tab that
    rolls up the headline numbers an IT or security team cares about.

    Mailbox / OneDrive / SharePoint sizing comes from the Graph **usage reports**
    API (getMailboxUsageDetail, getOneDriveUsageAccountDetail,
    getSharePointSiteUsageDetail). This keeps the whole script on the Graph SDK
    (no Exchange Online module required).

.PARAMETER OutputPath
    Folder or full .xlsx path for the report. Defaults to the script folder with
    a timestamped filename.

.PARAMETER TenantId
    Optional tenant id / domain to target a specific tenant at sign-in.

.PARAMETER Workload
    Which workloads to collect. Defaults to all four. Accepts any combination of:
    Identity, Exchange, SharePoint, Teams.

.PARAMETER UsagePeriod
    Reporting period for the usage-report data: D7, D30, D90, or D180. Default D30.

.PARAMETER UseAppOnly
    Use app-only (unattended) auth instead of interactive. Requires -ClientId,
    -TenantId and either -CertificateThumbprint.

.PARAMETER ClientId
    App registration (client) id for app-only auth.

.PARAMETER CertificateThumbprint
    Thumbprint of a cert in the local cert store for app-only auth.

.PARAMETER SkipModuleInstall
    Don't attempt to install missing modules; fail instead.

.PARAMETER SkipPublicFolders
    Skip public folder discovery (folder tree, item counts/size, mail-enabled
    addresses). Requires the same Exchange Online connection as mailbox type /
    delegation, and runs even if -SkipMailboxType is passed.

.PARAMETER SkipPublicFolderPermissions
    Skip the public folder client-permissions sheet (one call per folder) while
    still collecting the folder inventory itself.

.PARAMETER SkipMailForwarding
    Skip the mail-forwarding sheet (mailbox-level ForwardingSmtpAddress plus a
    per-mailbox inbox-rule scan for forward/redirect rules). Flags forwarding
    to addresses outside the tenant's verified domains, so it's one of the more
    useful security checks in the report; skip it on very large tenants where
    the per-mailbox inbox-rule calls add meaningful runtime.

.EXAMPLE
    .\New-M365TenantReport.ps1
    Interactive sign-in, all workloads, report dropped next to the script.

.EXAMPLE
    .\New-M365TenantReport.ps1 -OutputPath C:\Reports -Workload Identity,Exchange -UsagePeriod D90

.EXAMPLE
    .\New-M365TenantReport.ps1 -UseAppOnly -TenantId contoso.onmicrosoft.com `
        -ClientId 1111-... -CertificateThumbprint ABCD... -OutputPath C:\Reports

.NOTES
    Required Graph permissions (delegated or application):
        Organization.Read.All, User.Read.All, Group.Read.All, GroupMember.Read.All,
        Directory.Read.All, Sites.Read.All, Reports.Read.All,
        AuditLog.Read.All (auth-method registration report),
        RoleManagement.Read.Directory (privileged role holders),
        SharePointTenantSettings.Read.All (tenant-wide external sharing settings)

    Usage-report user/site names are blank/anonymised if the tenant has
    "Reports: Display concealed user, group and site names" turned ON in
    M365 admin center > Settings > Org settings > Reports. Turn it off (or accept
    GUIDs) before running for full detail.
#>
[CmdletBinding()]
param(
    [string]$OutputPath,

    [string]$TenantId,

    [ValidateSet('Identity', 'Exchange', 'SharePoint', 'Teams')]
    [string[]]$Workload = @('Identity', 'Exchange', 'SharePoint', 'Teams'),

    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$UsagePeriod = 'D30',

    [switch]$UseAppOnly,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    [switch]$SkipModuleInstall,

    # Mailbox type (User/Shared/Room/Equipment) is not available from Graph, so by
    # default the Exchange workload also connects to Exchange Online PowerShell to
    # enrich the Mailboxes sheet. Use -SkipMailboxType to stay pure-Graph (no ExO
    # module / second sign-in); the MailboxType column then reads "Unknown".
    [switch]$SkipMailboxType,

    # Per-mailbox delegation (FullAccess / SendAs / SendOnBehalf) is collected via
    # Exchange Online when the Exchange workload runs. It makes a couple of calls
    # per mailbox, so use -SkipDelegation on very large tenants to skip that sheet.
    [switch]$SkipDelegation,

    # Mail forwarding: mailbox-level ForwardingSmtpAddress plus a per-mailbox
    # inbox-rule scan for forward/redirect rules, flagged when the target is
    # outside the tenant's verified domains. One extra call per mailbox (the
    # inbox-rule scan), so use -SkipMailForwarding on very large tenants.
    [switch]$SkipMailForwarding,

    # Public folder discovery (folder tree, item counts/size, mail-enabled
    # addresses) via Exchange Online. Runs even if -SkipMailboxType is used,
    # since it only needs the ExO connection, not the mailbox list.
    [switch]$SkipPublicFolders,

    # Public folder client permissions (Owner/Editor/Reviewer per folder) are
    # one call per folder, so use -SkipPublicFolderPermissions on large PF
    # trees to skip just that sheet while still getting the folder inventory.
    [switch]$SkipPublicFolderPermissions,

    # License friendly-name catalog. By default the script downloads Microsoft's
    # published "Product names and service plan identifiers" CSV. Use
    # -OfflineSkuNames to skip the download (built-in map only), or -SkuCatalogPath
    # to point at a previously saved copy of the CSV (air-gapped runs).
    [switch]$OfflineSkuNames,
    [string]$SkuCatalogPath,
    [string]$SkuCatalogUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv',

    # Self-update. On start the script checks GitHub for a newer version, shows the
    # changelog, and (after you confirm) downloads and replaces itself.
    #   -SkipUpdateCheck : don't check at all
    #   -AutoUpdate      : apply a newer version without prompting (unattended)
    #   -UpdateToken     : a GitHub PAT, only needed while the repo is private
    [switch]$SkipUpdateCheck,
    [switch]$AutoUpdate,
    [string]$UpdateToken
)

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

# Version + self-update source. Keep the $ScriptVersion line in this exact format;
# the updater parses it out of the remote copy to detect newer releases.
$ScriptVersion       = '2.0.0'
$script:RepoOwner    = 'Zanrose'
$script:RepoName     = 'office365-scripts'
$script:RepoBranch   = 'main'
$script:ScriptFileName = 'New-M365TenantReport.ps1'
# office365-scripts is a flat, multi-script collection repo, so this script's
# changelog lives under its own name rather than a repo-root CHANGELOG.md.
$script:ChangelogFileName = 'CHANGELOG-TenantReport.md'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Step {
    param([string]$Message, [string]$Status = 'INFO')
    $color = switch ($Status) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host ("[{0:HH:mm:ss}] " -f (Get-Date)) -NoNewline -ForegroundColor DarkGray
    Write-Host ("{0,-5} " -f $Status) -NoNewline -ForegroundColor $color
    Write-Host $Message
}

# --- Self-update ----------------------------------------------------------

# Fetch a file from the repo. Tries the gh CLI first (uses the caller's existing
# auth, so it works while the repo is private), then the anonymous/token REST API
# (works once the repo is public, or with -UpdateToken / $env:GITHUB_TOKEN).
function Get-RemoteRepoFile {
    param([string]$Path, [string]$Ref = $script:RepoBranch)
    $api = "repos/$($script:RepoOwner)/$($script:RepoName)/contents/$Path"
    $gh  = Get-Command gh -ErrorAction SilentlyContinue
    if ($gh) {
        try {
            $out = & $gh.Source api "$api`?ref=$Ref" -H 'Accept: application/vnd.github.raw' 2>$null
            if ($LASTEXITCODE -eq 0 -and $out) { return ($out -join "`n") }
        } catch { }
    }
    $uri     = "https://api.github.com/$api`?ref=$Ref"
    $headers = @{ 'Accept' = 'application/vnd.github.raw'; 'User-Agent' = 'M365TenantReport-Updater' }
    $token   = if ($UpdateToken) { $UpdateToken } elseif ($env:GITHUB_TOKEN) { $env:GITHUB_TOKEN } else { $null }
    if ($token) { $headers['Authorization'] = "token $token" }
    return Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 30 -ErrorAction Stop
}

# Call a GitHub REST API path and return parsed JSON (gh CLI -> token/anon REST).
function Invoke-GitHubApi {
    param([string]$ApiPath)
    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if ($gh) {
        try {
            $out = & $gh.Source api $ApiPath 2>$null
            if ($LASTEXITCODE -eq 0 -and $out) { return (($out -join "`n") | ConvertFrom-Json) }
        } catch { }
    }
    $uri     = "https://api.github.com/$ApiPath"
    $headers = @{ 'Accept' = 'application/vnd.github+json'; 'User-Agent' = 'M365TenantReport-Updater' }
    $token   = if ($UpdateToken) { $UpdateToken } elseif ($env:GITHUB_TOKEN) { $env:GITHUB_TOKEN } else { $null }
    if ($token) { $headers['Authorization'] = "token $token" }
    return Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 30 -ErrorAction Stop
}

# Print only the changelog sections newer than the installed version.
function Show-ChangelogSince {
    param([string]$Changelog, [version]$Since)
    $show = $false; $any = $false
    foreach ($line in ($Changelog -split "`r?`n")) {
        $h = [regex]::Match($line, '^##+\s*\[?v?([0-9]+\.[0-9]+\.[0-9]+)\]?')
        if ($h.Success) {
            $show = ([version]$h.Groups[1].Value -gt $Since)
            if ($show) { $any = $true }
        }
        if ($show) { Write-Host "   $line" }
    }
    if (-not $any) { Write-Host $Changelog }   # fallback: unparseable, show all
}

# Validate the downloaded script, back up the current file, and replace it.
function Install-ScriptUpdate {
    param([string]$NewContent, [version]$NewVersion)
    $target = $PSCommandPath
    if (-not $target) {
        Write-Step "Cannot self-update: script path unknown (not running from a .ps1 file)." 'ERROR'
        return
    }
    $perr = $null
    [System.Management.Automation.Language.Parser]::ParseInput($NewContent, [ref]$null, [ref]$perr) | Out-Null
    if ($perr -and $perr.Count) {
        Write-Step "Downloaded update failed syntax validation; keeping current version." 'ERROR'
        return
    }
    try {
        $backup = "$target.v$ScriptVersion.bak"
        Copy-Item -LiteralPath $target -Destination $backup -Force
        Set-Content -LiteralPath $target -Value $NewContent -Encoding UTF8
        Write-Step "Updated to v$NewVersion. Backup: $(Split-Path $backup -Leaf)" 'OK'
        Write-Step "Please re-run the script to use the new version." 'OK'
        exit 0
    } catch {
        Write-Step "Update failed while writing the file: $($_.Exception.Message)" 'ERROR'
    }
}

# Orchestrates the check: find the latest version, show notes, prompt, apply.
# Primary source is the latest published GitHub Release; if none exist (or the
# call fails) it falls back to parsing $ScriptVersion off the main branch.
function Invoke-UpdateCheck {
    if ($SkipUpdateCheck) { return }
    Write-Step "Checking for updates (current v$ScriptVersion)..." 'INFO'
    $localVer  = [version]$ScriptVersion
    $remoteVer = $null
    $notes     = $null
    $ref       = $script:RepoBranch

    # Primary: latest published release
    try {
        $rel = Invoke-GitHubApi -ApiPath "repos/$($script:RepoOwner)/$($script:RepoName)/releases/latest"
        if ($rel -and $rel.tag_name) {
            $vm = [regex]::Match([string]$rel.tag_name, '([0-9]+\.[0-9]+\.[0-9]+)')
            if ($vm.Success) {
                $remoteVer = [version]$vm.Groups[1].Value
                $ref       = $rel.tag_name
                $notes     = $rel.body
            }
        }
    } catch { }

    # Fallback: parse the version marker off the main branch
    if (-not $remoteVer) {
        try {
            $remoteScript = Get-RemoteRepoFile -Path $script:ScriptFileName
            $m = [regex]::Match($remoteScript, '\$ScriptVersion\s*=\s*''([\d.]+)''')
            if ($m.Success) { $remoteVer = [version]$m.Groups[1].Value; $ref = $script:RepoBranch }
        } catch {
            Write-Step "  -> update check skipped ($($_.Exception.Message))" 'WARN'
            return
        }
    }
    if (-not $remoteVer) { Write-Step "  -> couldn't determine latest version; skipping" 'WARN'; return }
    if ($remoteVer -le $localVer) { Write-Step "  -> up to date" 'OK'; return }

    Write-Step "  -> update available: v$localVer -> v$remoteVer" 'WARN'
    Write-Host ""
    Write-Host "  ----------------- What's new -----------------" -ForegroundColor Cyan
    if ($notes) {
        foreach ($l in ($notes -split "`r?`n")) { Write-Host "   $l" }
    } else {
        $changelog = $null
        try { $changelog = Get-RemoteRepoFile -Path $script:ChangelogFileName } catch { }
        if ($changelog) { Show-ChangelogSince -Changelog $changelog -Since $localVer }
        else { Write-Host "   (no release notes found)" }
    }
    Write-Host "  ----------------------------------------------" -ForegroundColor Cyan
    Write-Host ""

    $accept = $false
    if ($AutoUpdate) {
        $accept = $true
    } elseif (-not [Environment]::UserInteractive) {
        Write-Step "Non-interactive session; skipping. Use -AutoUpdate to apply unattended." 'WARN'
        return
    } else {
        $answer = Read-Host "  Download and apply update v$remoteVer now? (Y/N)"
        $accept = ($answer -match '^(y|yes)$')
    }
    if (-not $accept) { Write-Step "Update declined; continuing with v$localVer." 'INFO'; return }

    # Download the script at the resolved ref (release tag, or main for the fallback)
    try {
        $newContent = Get-RemoteRepoFile -Path $script:ScriptFileName -Ref $ref
    } catch {
        Write-Step "Failed to download update: $($_.Exception.Message)" 'ERROR'
        return
    }
    Install-ScriptUpdate -NewContent $newContent -NewVersion $remoteVer
}

function Initialize-Module {
    param([string]$Name, [string]$MinimumVersion)
    $existing = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending | Select-Object -First 1
    if ($existing -and (-not $MinimumVersion -or $existing.Version -ge [version]$MinimumVersion)) {
        return
    }
    if ($SkipModuleInstall) {
        throw "Required module '$Name' is missing and -SkipModuleInstall was specified."
    }
    Write-Step "Installing module $Name ..." 'INFO'
    $scope = if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { 'AllUsers' } else { 'CurrentUser' }
    Install-Module -Name $Name -Scope $scope -Force -AllowClobber -Repository PSGallery
}

# Safely run a collection block; on failure log and return an empty array so one
# broken workload never kills the whole report.
function Invoke-Collection {
    param([string]$Name, [scriptblock]$Script)
    Write-Step "Collecting: $Name" 'INFO'
    try {
        $data = & $Script
        $count = @($data).Count
        Write-Step "  -> $count record(s)" 'OK'
        return , @($data)
    } catch {
        Write-Step "  -> FAILED: $($_.Exception.Message)" 'ERROR'
        return , @()
    }
}

# Download a Graph usage report CSV and return it as objects.
function Get-GraphUsageReport {
    param([string]$ReportName, [string]$Period)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("graphrpt_{0}_{1}.csv" -f $ReportName, [guid]::NewGuid())
    $uri = "https://graph.microsoft.com/v1.0/reports/$ReportName(period='$Period')"
    try {
        Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $tmp | Out-Null
        if (Test-Path $tmp) {
            return Import-Csv -Path $tmp
        }
        return @()
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Format-Gb { param($Bytes) if ($null -eq $Bytes -or $Bytes -eq '') { 0 } else { [math]::Round([double]$Bytes / 1GB, 2) } }

# Get-PublicFolder / Get-PublicFolderStatistics (Exchange Online V3 module) return
# FolderPath as an array of path segments rather than a single delimited string
# like the classic module did. Normalize to "\Segment1\Segment2" either way, since
# using the raw array as an export column or a dictionary key silently breaks both
# (ArrayList has no useful ToString() and no value-based equality).
function Get-PfPathString {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    return '\' + ((@($Value) | Where-Object { $_ }) -join '\')
}

# Download (or load) Microsoft's licensing CSV and return skuPartNumber/skuId ->
# friendly-name lookups. The CSV has one row per service plan, so we de-dupe to
# the first product name per SKU. Returns $null if it can't be obtained.
function Get-SkuCatalog {
    param([string]$Url, [string]$LocalPath)
    $byPart = @{}; $byGuid = @{}
    $tmp = $null
    try {
        $src = $LocalPath
        if (-not $src) {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("m365skucatalog_{0}.csv" -f [guid]::NewGuid())
            Invoke-WebRequest -Uri $Url -OutFile $tmp -UseBasicParsing -TimeoutSec 60
            $src = $tmp
        }
        if (-not (Test-Path $src)) { return $null }
        Import-Csv -Path $src | ForEach-Object {
            if ($_.String_Id -and -not $byPart.ContainsKey($_.String_Id)) { $byPart[$_.String_Id] = $_.Product_Display_Name }
            if ($_.GUID      -and -not $byGuid.ContainsKey($_.GUID))      { $byGuid[$_.GUID]      = $_.Product_Display_Name }
        }
        if ($byPart.Count -eq 0) { return $null }
        return [pscustomobject]@{ ByPartNumber = $byPart; ByGuid = $byGuid; Count = $byPart.Count }
    } catch {
        return $null
    } finally {
        if ($tmp -and (Test-Path $tmp)) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

# Classify an MX host to a mail-security provider (best effort, by DNS suffix).
function Get-MxProvider {
    param([string]$Exchange)
    $e = ("$Exchange").ToLower().TrimEnd('.')
    switch -Wildcard ($e) {
        '*.mail.protection.outlook.com' { 'Microsoft 365 (Exchange Online Protection)'; break }
        '*.pphosted.com'                { 'Proofpoint'; break }
        '*.ppe-hosted.com'              { 'Proofpoint Essentials'; break }
        '*.mimecast.com'                { 'Mimecast'; break }
        '*.mimecast.co*'                { 'Mimecast'; break }
        '*.barracudanetworks.com'       { 'Barracuda'; break }
        '*.cudaops.com'                 { 'Barracuda'; break }
        '*.mailcontrol.com'             { 'Forcepoint (MailControl)'; break }
        '*.iphmx.com'                   { 'Cisco IronPort / Trend Micro'; break }
        '*.trendmicro.com'              { 'Trend Micro'; break }
        '*.messagelabs.com'             { 'Broadcom / Symantec MessageLabs'; break }
        '*aspmx.l.google.com'           { 'Google Workspace'; break }
        '*.googlemail.com'              { 'Google Workspace'; break }
        '*.sophos.com'                  { 'Sophos'; break }
        '*.antispamcloud.com'           { 'SpamExperts'; break }
        '*.hostedemail.com'             { 'Hosted email filter'; break }
        '*.securence.com'               { 'Securence'; break }
        '*fortimail*'                   { 'Fortinet FortiMail'; break }
        '*.dnsmadeeasy.com'             { 'DNS Made Easy (relay)'; break }
        default                         { 'Other / self-hosted / unknown' }
    }
}

# Resolve a license to its friendly name: Microsoft catalog (by GUID, then part
# number) -> built-in static map -> raw part number -> raw GUID.
function Resolve-SkuName {
    param($PartNumber, $SkuId, $Catalog, $StaticMap)
    if ($Catalog) {
        if ($SkuId      -and $Catalog.ByGuid.ContainsKey($SkuId))            { return $Catalog.ByGuid[$SkuId] }
        if ($PartNumber -and $Catalog.ByPartNumber.ContainsKey($PartNumber)) { return $Catalog.ByPartNumber[$PartNumber] }
    }
    if ($PartNumber -and $StaticMap -and $StaticMap.ContainsKey($PartNumber)) { return $StaticMap[$PartNumber] }
    if ($PartNumber) { return $PartNumber }
    return $SkuId
}

# ---------------------------------------------------------------------------
# 0. Modules & connection
# ---------------------------------------------------------------------------

Write-Step "M365 Tenant Report (v$ScriptVersion)" 'INFO'
Write-Step "Workloads: $($Workload -join ', ')  |  Usage period: $UsagePeriod" 'INFO'

# Check for a newer version before doing any work (may prompt / self-update / exit).
Invoke-UpdateCheck

Initialize-Module -Name 'Microsoft.Graph.Authentication'
Initialize-Module -Name 'ImportExcel'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

$scopes = @(
    'Organization.Read.All', 'User.Read.All', 'Group.Read.All', 'GroupMember.Read.All',
    'Directory.Read.All', 'Sites.Read.All', 'Reports.Read.All', 'AuditLog.Read.All',
    'Domain.Read.All', 'RoleManagement.Read.Directory', 'SharePointTenantSettings.Read.All'
)

Write-Step "Connecting to Microsoft Graph..." 'INFO'
if ($UseAppOnly) {
    if (-not ($ClientId -and $TenantId -and $CertificateThumbprint)) {
        throw "App-only auth requires -ClientId, -TenantId and -CertificateThumbprint."
    }
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId `
        -CertificateThumbprint $CertificateThumbprint -NoWelcome
} else {
    $connectParams = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams
}

$ctx = Get-MgContext
if (-not $ctx) { throw "Failed to establish a Graph connection." }
Write-Step "Connected as '$($ctx.Account)' to tenant '$($ctx.TenantId)'" 'OK'

# Resolve output location. Default is the user's Documents folder — resolved via
# GetFolderPath('MyDocuments'), which returns the *redirected* path when OneDrive
# Known Folder Move backs up Documents. We drop reports into a tidy subfolder.
if (-not $OutputPath) {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($docs)) { $docs = Join-Path $env:USERPROFILE 'Documents' }
    $OutputPath = Join-Path $docs 'M365 Tenant Reports'
}
# Treat a path ending in .xlsx as an explicit file; anything else is a folder
# (and may not exist yet, so we don't rely on Test-Path to classify it).
if ($OutputPath -match '\.xlsx$') {
    $script:OutputDir  = Split-Path $OutputPath -Parent
    $script:OutputFile = $OutputPath
} else {
    $script:OutputDir  = $OutputPath
    $script:OutputFile = $null
}
if ($script:OutputDir -and -not (Test-Path $script:OutputDir)) {
    New-Item -ItemType Directory -Path $script:OutputDir -Force | Out-Null
}

$collected = [ordered]@{}   # worksheet name -> object[]
$summary   = [System.Collections.Generic.List[object]]::new()
$script:ExoConnected = $false
function Add-Summary { param($Metric, $Value) $summary.Add([pscustomobject]@{ Metric = $Metric; Value = $Value }) }

# ---------------------------------------------------------------------------
# 1. Tenant / organization overview (always)
# ---------------------------------------------------------------------------
$org = Invoke-Collection 'Tenant overview' {
    $resp    = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization'
    $o       = @($resp.value)[0]
    $domains = @($o.verifiedDomains)
    # Split verified domains into onmicrosoft.com (tenant-issued, carries no
    # branding) vs vanity/custom domains (the ones that carry the org's identity,
    # and that a tenant-to-tenant migration would need to re-verify and re-point).
    $onMicrosoftDomains = @($domains | Where-Object { $_.name -like '*.onmicrosoft.com' } | Select-Object -ExpandProperty name)
    $vanityDomains      = @($domains | Where-Object { $_.name -notlike '*.onmicrosoft.com' } | Select-Object -ExpandProperty name)
    [pscustomobject]@{
        DisplayName        = $o.displayName
        TenantId           = $o.id
        DefaultDomain      = ($domains | Where-Object { $_.isDefault }).name
        InitialDomain      = ($domains | Where-Object { $_.name -like '*.onmicrosoft.com' } | Select-Object -First 1).name
        OnMicrosoftDomains = ($onMicrosoftDomains -join '; ')
        VanityDomains      = ($vanityDomains -join '; ')
        Country            = $o.countryLetterCode
        CreatedDateTime    = $o.createdDateTime
        OnPremisesSyncEnabled = $o.onPremisesSyncEnabled
    }
}
$collected['Tenant'] = $org
if ($org) {
    Add-Summary 'Tenant'            $org[0].DisplayName
    Add-Summary 'Default domain'    $org[0].DefaultDomain
    Add-Summary 'Directory sync (AD Connect)' $org[0].OnPremisesSyncEnabled
    Add-Summary 'Report generated'  (Get-Date)
    Add-Summary 'Source account'    $ctx.Account
}

# Now that we know the tenant, build the filename: <Tenant>-TenantReport-<date>.xlsx
# Prefer the org display name; fall back to the initial onmicrosoft.com prefix,
# then the default domain prefix, then the tenant id.
if (-not $script:OutputFile) {
    $tenantLabel = $null
    if ($org) {
        $tenantLabel = $org[0].DisplayName
        if ([string]::IsNullOrWhiteSpace($tenantLabel)) {
            if     ($org[0].InitialDomain) { $tenantLabel = ($org[0].InitialDomain -split '\.')[0] }
            elseif ($org[0].DefaultDomain) { $tenantLabel = ($org[0].DefaultDomain -split '\.')[0] }
            elseif ($org[0].TenantId)      { $tenantLabel = $org[0].TenantId }
        }
    }
    if ([string]::IsNullOrWhiteSpace($tenantLabel)) { $tenantLabel = $ctx.TenantId }
    if ([string]::IsNullOrWhiteSpace($tenantLabel)) { $tenantLabel = 'M365' }

    $safe  = (($tenantLabel -replace '[\\/:*?"<>|]', '_').Trim()) -replace '\s+', '-'
    $stamp = Get-Date -Format 'yyyy-MM-dd'
    $script:OutputFile = Join-Path $script:OutputDir "$safe-TenantReport-$stamp.xlsx"
}
$OutputPath = $script:OutputFile
Write-Step "Report will be saved as: $(Split-Path $OutputPath -Leaf)" 'OK'

# ---------------------------------------------------------------------------
# 2. Identity & licensing
# ---------------------------------------------------------------------------
if ($Workload -contains 'Identity') {

    # SKU friendly-name lookup (subset of common SKUs; falls back to part number)
    $skuMap = @{
        'O365_BUSINESS_ESSENTIALS'='M365 Business Basic'; 'O365_BUSINESS_PREMIUM'='M365 Business Standard'
        'SPB'='M365 Business Premium'; 'ENTERPRISEPACK'='Office 365 E3'; 'ENTERPRISEPREMIUM'='Office 365 E5'
        'SPE_E3'='Microsoft 365 E3'; 'SPE_E5'='Microsoft 365 E5'; 'SPE_F1'='Microsoft 365 F3'
        'EXCHANGESTANDARD'='Exchange Online Plan 1'; 'EXCHANGEENTERPRISE'='Exchange Online Plan 2'
        'POWER_BI_STANDARD'='Power BI Free'; 'FLOW_FREE'='Power Automate Free'
        'TEAMS_EXPLORATORY'='Teams Exploratory'; 'STANDARDPACK'='Office 365 E1'
    }

    # Microsoft's published product-name catalog (friendly names). Falls back to
    # the static $skuMap above if the download is unavailable or -OfflineSkuNames.
    $skuCatalog = $null
    if ($OfflineSkuNames) {
        Write-Step "Using built-in license name map (-OfflineSkuNames)." 'INFO'
    } else {
        Write-Step "Fetching Microsoft license name catalog..." 'INFO'
        $skuCatalog = Get-SkuCatalog -Url $SkuCatalogUrl -LocalPath $SkuCatalogPath
        if ($skuCatalog) { Write-Step "  -> $($skuCatalog.Count) products in catalog" 'OK' }
        else { Write-Step "  -> catalog unavailable; using built-in name map" 'WARN' }
    }

    # Pull subscribed SKUs once; build the display rows AND the skuId -> friendly
    # name lookup from the same response so every assigned license resolves.
    $skuIdToName = @{}
    $skus = Invoke-Collection 'License SKUs' {
        $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus'
        $raw  = @($resp.value)
        foreach ($r in $raw) {
            $part = $r.skuPartNumber
            $name = Resolve-SkuName -PartNumber $part -SkuId $r.skuId -Catalog $skuCatalog -StaticMap $skuMap
            $skuIdToName[$r.skuId] = $name
            [pscustomobject]@{
                Product       = $name
                SkuPartNumber = $part
                Enabled       = $r.prepaidUnits.enabled
                Assigned      = $r.consumedUnits
                Available     = [int]$r.prepaidUnits.enabled - [int]$r.consumedUnits
                Warning       = $r.prepaidUnits.warning
                Suspended     = $r.prepaidUnits.suspended
            }
        }
    }
    $collected['Licenses'] = $skus
    if ($skus) {
        Add-Summary 'Licensed seats assigned' (($skus | Measure-Object Assigned -Sum).Sum)
        Add-Summary '  Unused/available seats (cost review)' (($skus | Measure-Object Available -Sum).Sum)
    }

    $users = Invoke-Collection 'Users' {
        $select = 'id,displayName,userPrincipalName,mail,userType,accountEnabled,createdDateTime,assignedLicenses,onPremisesSyncEnabled,usageLocation,department,jobTitle'
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=$select&`$top=999"
        $all = [System.Collections.Generic.List[object]]::new()
        do {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($u in $page.value) {
                $licNames = @()
                foreach ($al in $u.assignedLicenses) {
                    $nm = $skuIdToName[$al.skuId]
                    if (-not $nm) { $nm = Resolve-SkuName -PartNumber $null -SkuId $al.skuId -Catalog $skuCatalog -StaticMap $skuMap }
                    $licNames += $nm
                }
                $all.Add([pscustomobject]@{
                    DisplayName        = $u.displayName
                    UserPrincipalName  = $u.userPrincipalName
                    Mail               = $u.mail
                    UserType           = $u.userType
                    AccountEnabled     = $u.accountEnabled
                    DirSynced          = [bool]$u.onPremisesSyncEnabled
                    UsageLocation      = $u.usageLocation
                    Department         = $u.department
                    JobTitle           = $u.jobTitle
                    LicenseCount       = @($u.assignedLicenses).Count
                    Licenses           = ($licNames -join '; ')
                    CreatedDateTime    = $u.createdDateTime
                })
            }
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        $all
    }
    $collected['Users'] = $users
    if ($users) {
        $members = @($users | Where-Object UserType -eq 'Member')
        $guests  = @($users | Where-Object UserType -eq 'Guest')
        Add-Summary 'Total users'            $users.Count
        Add-Summary '  Members'              $members.Count
        Add-Summary '  Guests'               $guests.Count
        Add-Summary '  Disabled accounts'    (@($users | Where-Object { -not $_.AccountEnabled }).Count)
        Add-Summary '  Licensed users'       (@($users | Where-Object { $_.LicenseCount -gt 0 }).Count)
        Add-Summary '  Directory-synced'     (@($users | Where-Object DirSynced).Count)
    }

    $mfa = Invoke-Collection 'MFA / auth-method registration' {
        $uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?`$top=999"
        $all = [System.Collections.Generic.List[object]]::new()
        do {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($r in $page.value) {
                $all.Add([pscustomobject]@{
                    UserPrincipalName    = $r.userPrincipalName
                    DisplayName          = $r.userDisplayName
                    MfaCapable           = $r.isMfaCapable
                    MfaRegistered        = $r.isMfaRegistered
                    SsprCapable          = $r.isSsprCapable
                    DefaultMethod        = $r.defaultMfaMethod
                    Methods              = ($r.methodsRegistered -join '; ')
                    IsAdmin              = $r.isAdmin
                })
            }
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        $all
    }
    $collected['MFA-Registration'] = $mfa
    if ($mfa) {
        Add-Summary 'MFA-registered users' (@($mfa | Where-Object MfaRegistered).Count)
        Add-Summary 'MFA-capable users'    (@($mfa | Where-Object MfaCapable).Count)
        Add-Summary '  Users with NO MFA method registered' (@($mfa | Where-Object { -not $_.MfaRegistered }).Count)
    }

    # Domains and their authentication type. Federated domains (ADFS / third-party
    # IdP) are worth flagging on their own merits (sign-in depends on an external
    # service being up), and doubly so before a tenant-to-tenant move, since they
    # don't carry over and must be converted to managed or re-federated.
    $domains = Invoke-Collection 'Domains (federation / managed)' {
        $rows = [System.Collections.Generic.List[object]]::new()
        $uri = 'https://graph.microsoft.com/v1.0/domains'
        do {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($d in @($page.value)) {
                $issuer = $null; $passiveUri = $null
                if ($d.authenticationType -eq 'Federated') {
                    try {
                        $fc  = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains/$($d.id)/federationConfiguration"
                        $cfg = @($fc.value)[0]
                        $issuer     = $cfg.issuerUri
                        $passiveUri = $cfg.passiveSignInUri
                    } catch { }
                }
                $rows.Add([pscustomobject]@{
                    Domain              = $d.id
                    AuthenticationType  = $d.authenticationType   # Managed | Federated
                    IsFederated         = ($d.authenticationType -eq 'Federated')
                    IsDefault           = $d.isDefault
                    IsInitial           = $d.isInitial
                    IsVerified          = $d.isVerified
                    IsAdminManaged      = $d.isAdminManaged
                    SupportedServices   = ($d.supportedServices -join '; ')
                    FederationIssuerUri = $issuer
                    PassiveSignInUri    = $passiveUri
                })
            }
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        $rows
    }
    $collected['Domains'] = $domains
    if ($domains) {
        $fed = @($domains | Where-Object IsFederated)
        Add-Summary 'Verified domains'  (@($domains | Where-Object IsVerified).Count)
        Add-Summary 'Federated domains' ($(if ($fed.Count) { "$($fed.Count): " + ($fed.Domain -join ', ') } else { 'None (all managed)' }))
    }

    # Privileged (admin) role holders. /directoryRoles only returns roles that
    # have actually been activated in the tenant, which is exactly the set worth
    # reporting - nobody needs to see the ~100 roles nobody has ever assigned.
    # Members can be users, role-assignable groups, or service principals; the
    # @odata.type Graph returns on each member tells them apart.
    $privRoles = Invoke-Collection 'Privileged role assignments' {
        $rows  = [System.Collections.Generic.List[object]]::new()
        $roles = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName'
        foreach ($role in @($roles.value)) {
            $uri = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$top=999"
            do {
                $page = Invoke-MgGraphRequest -Method GET -Uri $uri
                foreach ($m in @($page.value)) {
                    $rows.Add([pscustomobject]@{
                        RoleName          = $role.displayName
                        DisplayName       = $m.displayName
                        UserPrincipalName = $m.userPrincipalName
                        PrincipalType     = ("$($m.'@odata.type')" -replace '^#microsoft\.graph\.', '')
                        AccountEnabled    = $m.accountEnabled
                    })
                }
                $uri = $page.'@odata.nextLink'
            } while ($uri)
        }
        $rows
    }
    $collected['Privileged-Roles'] = $privRoles
    if ($privRoles) {
        $gaCount     = @($privRoles | Where-Object { $_.RoleName -eq 'Global Administrator' }).Count
        $multiRole   = @($privRoles | Where-Object UserPrincipalName | Group-Object UserPrincipalName | Where-Object Count -ge 2)
        Add-Summary 'Privileged role assignments'       $privRoles.Count
        Add-Summary '  Global Administrators'           $gaCount
        Add-Summary '  Principals holding 2+ admin roles' $multiRole.Count
    }
}

# ---------------------------------------------------------------------------
# 3. Exchange / mailboxes (usage report)
# ---------------------------------------------------------------------------
if ($Workload -contains 'Exchange') {

    # Mailbox type (User/Shared/Room/Equipment) isn't exposed by Graph, so connect
    # to Exchange Online and build a UPN/SMTP -> RecipientTypeDetails lookup.
    # Best-effort: any failure leaves the type as "Unknown" rather than aborting.
    # The same connection also feeds Delegation and Public Folders below, so it's
    # established whenever any of those three are wanted, not just mailbox type.
    $mbxTypeMap = @{}
    $exoMailboxes = @()
    if (-not $SkipMailboxType -or -not $SkipPublicFolders) {
        try {
            Initialize-Module -Name 'ExchangeOnlineManagement'
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Write-Step "Connecting to Exchange Online..." 'INFO'
            if ($UseAppOnly) {
                $exoOrg = if ($org -and $org[0].InitialDomain) { $org[0].InitialDomain }
                          elseif ($TenantId) { $TenantId } else { $ctx.TenantId }
                Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint `
                    -Organization $exoOrg -ShowBanner:$false -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
            $script:ExoConnected = $true

            if (-not $SkipMailboxType) {
                Write-Step "  -> retrieving mailboxes..." 'INFO'
                $exoMailboxes = @(Get-EXOMailbox -ResultSize Unlimited `
                    -Properties RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingAddress, `
                        ForwardingSmtpAddress, DeliverToMailboxAndForward -ErrorAction Stop)
                foreach ($m in $exoMailboxes) {
                    if ($m.UserPrincipalName)  { $mbxTypeMap[("$($m.UserPrincipalName)").ToLower()]  = $m.RecipientTypeDetails }
                    if ($m.PrimarySmtpAddress) { $mbxTypeMap[("$($m.PrimarySmtpAddress)").ToLower()] = $m.RecipientTypeDetails }
                }
                Write-Step "  -> $($exoMailboxes.Count) mailboxes, $($mbxTypeMap.Count) identities mapped" 'OK'
            }
        } catch {
            Write-Step "  -> Exchange Online unavailable: $($_.Exception.Message)" 'WARN'
            Write-Step "     MailboxType/Delegation/PublicFolders will be limited. Re-run once ExO is reachable." 'WARN'
        }
    }

    $mbx = Invoke-Collection 'Mailbox usage' {
        Get-GraphUsageReport -ReportName 'getMailboxUsageDetail' -Period $UsagePeriod | ForEach-Object {
            $upn   = $_.'User Principal Name'
            $mtype = if ($upn -and $mbxTypeMap.ContainsKey(("$upn").ToLower())) { $mbxTypeMap[("$upn").ToLower()] }
                     elseif ($SkipMailboxType) { 'Unknown (skipped)' } else { 'Unknown' }
            [pscustomobject]@{
                DisplayName       = $_.'Display Name'
                UserPrincipalName = $upn
                MailboxType       = $mtype
                IsShared          = ($mtype -eq 'SharedMailbox')
                IsDeleted         = $_.'Is Deleted'
                ItemCount         = [int64]($_.'Item Count' -as [int64])
                UsedGB            = Format-Gb $_.'Storage Used (Byte)'
                IssueWarningGB    = Format-Gb $_.'Issue Warning Quota (Byte)'
                ProhibitSendGB    = Format-Gb $_.'Prohibit Send Quota (Byte)'
                ProhibitSendRecvGB= Format-Gb $_.'Prohibit Send/Receive Quota (Byte)'
                LastActivityDate  = $_.'Last Activity Date'
                CreatedDate       = $_.'Created Date'
            }
        }
    }
    $collected['Mailboxes'] = $mbx
    if ($mbx) {
        Add-Summary 'Mailboxes'              $mbx.Count
        Add-Summary '  Shared mailboxes'     (@($mbx | Where-Object { $_.MailboxType -eq 'SharedMailbox' }).Count)
        Add-Summary '  Room mailboxes'       (@($mbx | Where-Object { $_.MailboxType -eq 'RoomMailbox' }).Count)
        Add-Summary '  Equipment mailboxes'  (@($mbx | Where-Object { $_.MailboxType -eq 'EquipmentMailbox' }).Count)
        Add-Summary 'Mailbox storage (GB)'   ([math]::Round((($mbx | Measure-Object UsedGB -Sum).Sum), 2))
        Add-Summary 'Largest mailbox (GB)'   (($mbx | Measure-Object UsedGB -Maximum).Maximum)
    }

    # Mailbox delegation: FullAccess, SendAs and SendOnBehalf. Worth documenting
    # in its own right (it's rarely tracked anywhere else), and essential before
    # a tenant-to-tenant move, since none of it carries over automatically and
    # has to be re-applied on the target. Requires the ExO connection above.
    if (-not $SkipDelegation -and $script:ExoConnected -and $exoMailboxes.Count) {
        $delegation = Invoke-Collection 'Mailbox delegation (FullAccess / SendAs / SendOnBehalf)' {
            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($m in $exoMailboxes) {
                $id = "$($m.PrimarySmtpAddress)"
                # FullAccess (explicit grants only; skip inherited and NT AUTHORITY\SELF)
                try {
                    Get-EXOMailboxPermission -Identity $id -ErrorAction Stop | Where-Object {
                        $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and
                        $_.User -notlike 'NT AUTHORITY\*' -and $_.User -ne $m.UserPrincipalName
                    } | ForEach-Object {
                        $rows.Add([pscustomobject]@{
                            Mailbox = $id; MailboxType = $m.RecipientTypeDetails
                            Permission = 'FullAccess'; Delegate = $_.User
                            # EXO doesn't expose the AutoMapping flag for existing grants
                            # (it lives in msExchDelegateListLink, not surfaced by cmdlets).
                            # Per MS docs, individual FullAccess grants automap by default
                            # unless -AutoMapping $false was used at grant time (undetectable).
                            AutoMapping = 'On (default)'
                        })
                    }
                } catch { }
                # SendAs
                try {
                    Get-EXORecipientPermission -Identity $id -ErrorAction Stop | Where-Object {
                        $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\*'
                    } | ForEach-Object {
                        $rows.Add([pscustomobject]@{
                            Mailbox = $id; MailboxType = $m.RecipientTypeDetails
                            Permission = 'SendAs'; Delegate = $_.Trustee
                            AutoMapping = 'N/A'
                        })
                    }
                } catch { }
                # SendOnBehalf (from the mailbox object)
                foreach ($sob in @($m.GrantSendOnBehalfTo)) {
                    if ($sob) {
                        $rows.Add([pscustomobject]@{
                            Mailbox = $id; MailboxType = $m.RecipientTypeDetails
                            Permission = 'SendOnBehalf'; Delegate = "$sob"
                            AutoMapping = 'N/A'
                        })
                    }
                }
            }
            $rows
        }
        $collected['Delegation'] = $delegation
        if ($delegation) {
            Add-Summary 'Delegation grants (Full/SendAs/OnBehalf)' $delegation.Count
            Add-Summary '  FullAccess grants (automap On by default)' (@($delegation | Where-Object { $_.Permission -eq 'FullAccess' }).Count)
            Add-Summary '  Mailboxes with delegation' (@($delegation | Select-Object -ExpandProperty Mailbox -Unique).Count)
        }
    }

    # Mail forwarding: mailbox-level ForwardingSmtpAddress/ForwardingAddress plus
    # a per-mailbox scan of inbox rules that forward or redirect. External
    # auto-forwarding is one of the most common indicators of a compromised
    # mailbox (BEC), and it's also easy to lose track of legitimately (someone
    # forwarding to a personal address before leaving), so it earns a dedicated
    # sheet rather than living inside the Mailboxes sheet.
    if (-not $SkipMailForwarding -and $script:ExoConnected -and $exoMailboxes.Count) {
        $verifiedDomains = @()
        try {
            $dResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/domains?$select=id,isVerified'
            $verifiedDomains = @($dResp.value | Where-Object isVerified | ForEach-Object { ("$($_.id)").ToLower() })
        } catch { }

        # Get-InboxRule's ForwardTo/RedirectTo/ForwardAsAttachmentTo resolve to
        # "Display Name [SMTP:address@domain.com]" for recipients inside the
        # tenant's address book, but come back as a bare address for anything
        # external/unresolved. Pull whichever form is present; if neither an
        # embedded SMTP address nor a bare address can be found, the target is
        # an internal recipient string we can't turn into a domain, so it's
        # reported with External = $null (verify manually) rather than guessed.
        function Get-ForwardTargetAddress {
            param([string]$Raw)
            if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
            $m = [regex]::Match($Raw, '\[SMTP:([^\]]+)\]', 'IgnoreCase')
            if ($m.Success) { return $m.Groups[1].Value }
            if ($Raw -match '^[^@\s]+@[^@\s]+$') { return $Raw }
            return $null
        }

        $forwarding = Invoke-Collection 'Mail forwarding (mailbox-level + inbox rules)' {
            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($m in $exoMailboxes) {
                $id = "$($m.PrimarySmtpAddress)"

                # Mailbox-level forwarding (Set-Mailbox -ForwardingSmtpAddress / -ForwardingAddress)
                if ($m.ForwardingSmtpAddress -or $m.ForwardingAddress) {
                    $target = if ($m.ForwardingSmtpAddress) { "$($m.ForwardingSmtpAddress)" -replace '^smtp:', '' }
                              else { "$($m.ForwardingAddress)" }
                    $domain = ($target -split '@')[-1].ToLower()
                    $rows.Add([pscustomobject]@{
                        Mailbox     = $id
                        MailboxType = $m.RecipientTypeDetails
                        Source      = 'Mailbox-level forwarding'
                        ForwardTo   = $target
                        KeepCopy    = [bool]$m.DeliverToMailboxAndForward
                        External    = ($domain -and $verifiedDomains -notcontains $domain)
                    })
                }

                # Inbox-rule forward / forward-as-attachment / redirect
                try {
                    Get-InboxRule -Mailbox $id -ErrorAction Stop |
                        Where-Object { $_.Enabled -and ($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo) } |
                        ForEach-Object {
                            $rule    = $_
                            $targets = @($rule.ForwardTo) + @($rule.ForwardAsAttachmentTo) + @($rule.RedirectTo) |
                                Where-Object { $_ } | ForEach-Object { "$_" } | Select-Object -Unique
                            foreach ($t in $targets) {
                                $addr   = Get-ForwardTargetAddress -Raw $t
                                $domain = if ($addr) { ($addr -split '@')[-1].ToLower() } else { $null }
                                $rows.Add([pscustomobject]@{
                                    Mailbox     = $id
                                    MailboxType = $m.RecipientTypeDetails
                                    Source      = "Inbox rule: $($rule.Name)"
                                    ForwardTo   = if ($addr) { $addr } else { $t }
                                    KeepCopy    = $null
                                    External    = if ($addr) { ($verifiedDomains -notcontains $domain) } else { $null }
                                })
                            }
                        }
                } catch { }
            }
            $rows
        }
        $collected['Mail-Forwarding'] = $forwarding
        if ($forwarding) {
            $extFwd = @($forwarding | Where-Object { $_.External -eq $true })
            $unk    = @($forwarding | Where-Object { $null -eq $_.External })
            Add-Summary 'Mail forwarding rules (mailbox + inbox rule)' $forwarding.Count
            Add-Summary '  Forwarding to EXTERNAL addresses' $extFwd.Count
            if ($unk.Count) { Add-Summary '  Forwarding targets needing manual review' $unk.Count }
        }
    }

    # Public folders: an inventory step most tenant documentation skips entirely,
    # since the tree, sizing, and mail-enabled addresses live outside the mailbox
    # list. That gap matters most before a tenant-to-tenant move, where public
    # folders have no native migration path and the mail-enabled ones need their
    # SMTP/proxy addresses re-created on the target before mail flow works there.
    if (-not $SkipPublicFolders -and $script:ExoConnected) {
        $pfEnabled = $null
        try { $pfEnabled = "$((Get-OrganizationConfig -ErrorAction Stop).PublicFoldersEnabled)" } catch { }

        if ($pfEnabled -eq 'None' -or [string]::IsNullOrWhiteSpace($pfEnabled)) {
            Write-Step "Public folders: not enabled for this tenant" 'INFO'
            Add-Summary 'Public folders enabled' 'No'
        } else {
            Add-Summary 'Public folders enabled' "Yes ($pfEnabled)"

            # Mail-enabled folders, keyed by the GUID that links them back to the
            # folder object (Get-PublicFolder.MailRecipientGuid).
            $mailPfMap = @{}
            try {
                Get-MailPublicFolder -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
                    if ($_.Guid) { $mailPfMap[$_.Guid.ToString()] = $_ }
                }
            } catch { }

            $statsMap = @{}
            try {
                Get-PublicFolderStatistics -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
                    $statsMap[(Get-PfPathString $_.FolderPath)] = $_
                }
            } catch { }

            $pfFolders = Invoke-Collection 'Public folders' {
                Get-PublicFolder -Identity '\' -Recurse -ResultSize Unlimited -ErrorAction Stop |
                    ForEach-Object {
                        $folderPath = Get-PfPathString $_.FolderPath
                        $stat   = $statsMap[$folderPath]
                        $mailPf = if ($_.MailEnabled -and $_.MailRecipientGuid) { $mailPfMap[$_.MailRecipientGuid.ToString()] } else { $null }
                        $sizeGb = 0
                        if ($stat -and $stat.TotalItemSize) {
                            try { $sizeGb = Format-Gb $stat.TotalItemSize.Value.ToBytes() } catch { }
                        }
                        [pscustomobject]@{
                            FolderPath         = $folderPath
                            MailEnabled        = [bool]$_.MailEnabled
                            HasSubfolders      = [bool]$_.HasSubfolders
                            ItemCount          = if ($stat) { [int64]$stat.ItemCount } else { 0 }
                            SizeGB             = $sizeGb
                            LastModified       = if ($stat) { $stat.LastModificationTime } else { $null }
                            PrimarySmtpAddress = if ($mailPf) { $mailPf.PrimarySmtpAddress } else { $null }
                            EmailAddresses     = if ($mailPf) { ($mailPf.EmailAddresses -join '; ') } else { $null }
                        }
                    }
            }
            $collected['PublicFolders'] = $pfFolders
            if ($pfFolders) {
                $mailEnabledPf = @($pfFolders | Where-Object MailEnabled)
                Add-Summary 'Public folders'                $pfFolders.Count
                Add-Summary '  Mail-enabled public folders' $mailEnabledPf.Count
                Add-Summary 'Public folder storage (GB)'    ([math]::Round((($pfFolders | Measure-Object SizeGB -Sum).Sum), 2))
            }

            # Client permissions (Owner/Editor/Reviewer/etc.) are an access-review
            # item on their own, and don't transfer in a tenant-to-tenant move
            # either, so they'd need to be re-applied on the target there too.
            if (-not $SkipPublicFolderPermissions -and $pfFolders.Count) {
                $pfPerms = Invoke-Collection 'Public folder client permissions' {
                    $rows = [System.Collections.Generic.List[object]]::new()
                    foreach ($f in $pfFolders) {
                        try {
                            Get-PublicFolderClientPermission -Identity $f.FolderPath -ErrorAction Stop | ForEach-Object {
                                foreach ($right in @($_.AccessRights)) {
                                    $rows.Add([pscustomobject]@{
                                        FolderPath  = $f.FolderPath
                                        User        = "$($_.User)"
                                        AccessRight = $right
                                    })
                                }
                            }
                        } catch { }
                    }
                    $rows
                }
                $collected['PublicFolder-Permissions'] = $pfPerms
                if ($pfPerms) { Add-Summary 'Public folder permission grants' $pfPerms.Count }
            }
        }
    } elseif (-not $SkipPublicFolders) {
        Write-Step "Public folders: skipped (Exchange Online not connected)" 'WARN'
    }

    # Distribution lists / mail-enabled groups (Graph).
    $distGroups = Invoke-Collection 'Mail-enabled groups' {
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=mailEnabled eq true&`$select=displayName,mail,mailNickname,groupTypes,securityEnabled&`$top=999"
        $all = [System.Collections.Generic.List[object]]::new()
        do {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($g in $page.value) {
                $type = if ($g.groupTypes -contains 'Unified') { 'M365 Group' }
                        elseif ($g.securityEnabled) { 'Mail-enabled security' } else { 'Distribution list' }
                $all.Add([pscustomobject]@{
                    DisplayName = $g.displayName; Mail = $g.mail
                    Alias = $g.mailNickname; Type = $type
                })
            }
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        $all
    }
    $collected['MailGroups'] = $distGroups
    if ($distGroups) { Add-Summary 'Mail-enabled groups' $distGroups.Count }

    # MX lookup: is inbound mail delivered straight to M365/EOP, or through a
    # third-party filtering gateway (Mimecast/Proofpoint/Barracuda/etc.)? Useful
    # for documenting the current mail path, and essential before a
    # tenant-to-tenant move, since that gateway has to be re-pointed at cutover.
    $mxRows = Invoke-Collection 'MX / inbound mail routing' {
        $rows = [System.Collections.Generic.List[object]]::new()
        $domainList = @()
        if ($org -and $org[0].VanityDomains) { $domainList = $org[0].VanityDomains -split '\s*;\s*' }
        $domainList = $domainList | Where-Object { $_ } | Sort-Object -Unique
        foreach ($d in $domainList) {
            try {
                $recs = @(Resolve-DnsName -Name $d -Type MX -ErrorAction Stop |
                    Where-Object { $_.QueryType -eq 'MX' } | Sort-Object Preference)
                if (-not $recs) {
                    $rows.Add([pscustomobject]@{
                        Domain = $d; Preference = $null; MailExchange = '(no MX record)'
                        Provider = 'None'; ThirdPartyInbound = $false
                    })
                    continue
                }
                foreach ($r in $recs) {
                    $isM365 = (("$($r.NameExchange)").ToLower().TrimEnd('.')) -like '*.mail.protection.outlook.com'
                    $rows.Add([pscustomobject]@{
                        Domain = $d; Preference = $r.Preference; MailExchange = $r.NameExchange
                        Provider = (Get-MxProvider -Exchange $r.NameExchange)
                        ThirdPartyInbound = (-not $isM365)
                    })
                }
            } catch {
                $rows.Add([pscustomobject]@{
                    Domain = $d; Preference = $null
                    MailExchange = "(lookup failed: $($_.Exception.Message))"
                    Provider = 'Unknown'; ThirdPartyInbound = $null
                })
            }
        }
        $rows
    }
    $collected['MX-Records'] = $mxRows
    if ($mxRows) {
        $tpDomains = @($mxRows | Where-Object { $_.ThirdPartyInbound -eq $true } |
            Select-Object -ExpandProperty Domain -Unique)
        if ($tpDomains.Count) {
            Add-Summary '3rd-party inbound mail filtering' ("Yes: " + ($tpDomains -join ', '))
        } else {
            Add-Summary '3rd-party inbound mail filtering' 'No (delivered directly to M365 / EOP)'
        }
    }
}

# ---------------------------------------------------------------------------
# 4. SharePoint & OneDrive (usage reports)
# ---------------------------------------------------------------------------
if ($Workload -contains 'SharePoint') {
    $spo = Invoke-Collection 'SharePoint site usage' {
        Get-GraphUsageReport -ReportName 'getSharePointSiteUsageDetail' -Period $UsagePeriod | ForEach-Object {
            [pscustomobject]@{
                SiteUrl          = $_.'Site URL'
                OwnerDisplayName = $_.'Owner Display Name'
                IsDeleted        = $_.'Is Deleted'
                FileCount        = [int64]($_.'File Count' -as [int64])
                ActiveFileCount  = [int64]($_.'Active File Count' -as [int64])
                StorageUsedGB    = Format-Gb $_.'Storage Used (Byte)'
                StorageAllocGB   = Format-Gb $_.'Storage Allocated (Byte)'
                LastActivityDate = $_.'Last Activity Date'
                Template         = $_.'Root Web Template'
            }
        }
    }
    $collected['SharePoint'] = $spo
    if ($spo) {
        Add-Summary 'SharePoint sites'        $spo.Count
        Add-Summary 'SharePoint storage (GB)' ([math]::Round((($spo | Measure-Object StorageUsedGB -Sum).Sum), 2))
        Add-Summary '  Sites with no owner'   (@($spo | Where-Object { [string]::IsNullOrWhiteSpace($_.OwnerDisplayName) }).Count)
    }

    # Tenant-wide external sharing posture. One call, and worth its own sheet:
    # this is the single setting that decides whether "Anyone" links are even
    # possible org-wide, and it rarely shows up in day-to-day admin reviews.
    # Requires SharePointTenantSettings.Read.All; wrapped in Invoke-Collection
    # so an older tenant or a missing consent just yields an empty sheet.
    $spSharing = Invoke-Collection 'SharePoint/OneDrive tenant sharing settings' {
        $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'
        [pscustomobject]@{
            SharingCapability           = $resp.sharingCapability
            SharingDomainRestrictionMode = $resp.sharingDomainRestrictionMode
            SharingAllowedDomainList    = ((@($resp.sharingAllowedDomainList)) -join '; ')
            SharingBlockedDomainList    = ((@($resp.sharingBlockedDomainList)) -join '; ')
        }
    }
    $collected['SharePoint-Sharing'] = $spSharing
    if ($spSharing -and $spSharing[0].SharingCapability) {
        Add-Summary 'SharePoint/OneDrive external sharing' $spSharing[0].SharingCapability
    }

    $od = Invoke-Collection 'OneDrive usage' {
        Get-GraphUsageReport -ReportName 'getOneDriveUsageAccountDetail' -Period $UsagePeriod | ForEach-Object {
            [pscustomobject]@{
                OwnerUPN         = $_.'Owner Principal Name'
                OwnerDisplayName = $_.'Owner Display Name'
                IsDeleted        = $_.'Is Deleted'
                FileCount        = [int64]($_.'File Count' -as [int64])
                ActiveFileCount  = [int64]($_.'Active File Count' -as [int64])
                StorageUsedGB    = Format-Gb $_.'Storage Used (Byte)'
                StorageAllocGB   = Format-Gb $_.'Storage Allocated (Byte)'
                LastActivityDate = $_.'Last Activity Date'
            }
        }
    }
    $collected['OneDrive'] = $od
    if ($od) {
        Add-Summary 'OneDrive accounts'     $od.Count
        Add-Summary 'OneDrive storage (GB)' ([math]::Round((($od | Measure-Object StorageUsedGB -Sum).Sum), 2))
    }
}

# ---------------------------------------------------------------------------
# 5. Teams & Groups
# ---------------------------------------------------------------------------
if ($Workload -contains 'Teams') {
    $groups = Invoke-Collection 'M365 groups & Teams' {
        $select = 'id,displayName,mail,visibility,createdDateTime,groupTypes,resourceProvisioningOptions'
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=$select&`$top=999"
        $all = [System.Collections.Generic.List[object]]::new()
        do {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($g in $page.value) {
                $isTeam = $g.resourceProvisioningOptions -contains 'Team'
                # owner / member counts
                $oc = 0; $mc = 0
                try {
                    $oc = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/owners/`$count" -Headers @{ ConsistencyLevel = 'eventual' })
                    $mc = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/members/`$count" -Headers @{ ConsistencyLevel = 'eventual' })
                } catch { }
                $all.Add([pscustomobject]@{
                    DisplayName     = $g.displayName
                    Mail            = $g.mail
                    IsTeam          = $isTeam
                    Visibility      = $g.visibility
                    OwnerCount      = $oc
                    MemberCount     = $mc
                    CreatedDateTime = $g.createdDateTime
                })
            }
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        $all
    }
    $collected['Teams-Groups'] = $groups
    if ($groups) {
        $teams = @($groups | Where-Object IsTeam)
        Add-Summary 'Microsoft 365 groups' $groups.Count
        Add-Summary '  of which Teams'      $teams.Count
        Add-Summary '  Public groups'       (@($groups | Where-Object Visibility -eq 'Public').Count)
        Add-Summary '  Groups w/o owner'    (@($groups | Where-Object { $_.OwnerCount -eq 0 }).Count)
    }
}

# ---------------------------------------------------------------------------
# 5b. Cross-workload hygiene checks (only run when both sides are available)
# ---------------------------------------------------------------------------

# Licensed shared/room/equipment mailboxes: these mailbox types get a free 50GB
# (100GB for shared, post-2023) without a license, so a license sitting on one
# is almost always leftover from a user-to-shared conversion, not a real need.
if ($collected.Contains('Users') -and $collected.Contains('Mailboxes')) {
    $licensedUpns = @{}
    foreach ($u in @($collected['Users'])) {
        if ($u.LicenseCount -gt 0 -and $u.UserPrincipalName) { $licensedUpns[("$($u.UserPrincipalName)").ToLower()] = $true }
    }
    $licensedNonUserMbx = @($collected['Mailboxes'] | Where-Object {
        $_.MailboxType -in @('SharedMailbox', 'RoomMailbox', 'EquipmentMailbox') -and
        $_.UserPrincipalName -and $licensedUpns.ContainsKey(("$($_.UserPrincipalName)").ToLower())
    })
    if ($licensedNonUserMbx.Count) {
        Add-Summary 'Licensed shared/room/equipment mailboxes (cost review)' $licensedNonUserMbx.Count
    }
}

# ---------------------------------------------------------------------------
# 6. Write the workbook
# ---------------------------------------------------------------------------
Write-Step "Writing workbook: $OutputPath" 'INFO'
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$elapsed = (New-TimeSpan -Start $script:StartTime -End (Get-Date))
Add-Summary 'Run duration' ("{0:mm\:ss}" -f $elapsed)

# Summary sheet first
$summary | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow `
    -Title 'M365 Tenant Report Summary' -TitleBold -TitleSize 14 -FreezeTopRow

foreach ($sheet in $collected.Keys) {
    $data = $collected[$sheet]
    if (-not $data -or @($data).Count -eq 0) {
        @([pscustomobject]@{ Note = 'No data collected (workload skipped, empty, or permission/concealment issue).' }) |
            Export-Excel -Path $OutputPath -WorksheetName $sheet -AutoSize
        continue
    }
    $data | Export-Excel -Path $OutputPath -WorksheetName $sheet -AutoSize -AutoFilter `
        -BoldTopRow -FreezeTopRow -TableStyle 'Medium2'
}

Write-Step "Done. Report saved to: $OutputPath" 'OK'
Write-Step "Worksheets: Summary, $($collected.Keys -join ', ')" 'INFO'

try { Disconnect-MgGraph | Out-Null } catch { }
if ($script:ExoConnected) { try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch { } }
