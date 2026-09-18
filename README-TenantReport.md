# M365 Tenant Report

`New-M365TenantReport.ps1` inventories a Microsoft 365 tenant across identity, security,
licensing, mail, and collaboration, and produces a multi-tab Excel workbook. It's useful on
its own for tenant documentation, security/hygiene reviews, and license audits, and equally
as the source-tenant inventory for a tenant-to-tenant migration.

> Formerly published as `New-M365PreMigrationReport.ps1` in
> [Zanrose/M365-PreMigration-Report](https://github.com/Zanrose/M365-PreMigration-Report).
> Same underlying collection, rebranded as a general tenant report and extended with the
> security/hygiene items below (v2.0.0).

## What it collects

| Worksheet            | Workload            | Source |
|----------------------|----------------------|--------|
| `Summary`            | headline roll-up     | all of the below |
| `Tenant`             | always                | `/organization` |
| `Licenses`           | Identity              | `/subscribedSkus` (incl. unused/available seats) |
| `Users`              | Identity              | `/users` (incl. per-user licenses, dir-sync, guest/member) |
| `MFA-Registration`   | Identity              | `/reports/authenticationMethods/userRegistrationDetails` |
| `Domains`            | Identity              | `/domains` (+ `federationConfiguration`) — Managed vs Federated |
| `Privileged-Roles`   | Identity              | `/directoryRoles` + `/directoryRoles/{id}/members` — admin role holders |
| `Mailboxes`          | Exchange              | `getMailboxUsageDetail` report + Exchange Online (`Get-EXOMailbox`) for mailbox type |
| `Delegation`         | Exchange              | Exchange Online — FullAccess / SendAs / SendOnBehalf per mailbox |
| `Mail-Forwarding`    | Exchange              | Exchange Online — mailbox-level forwarding + inbox-rule forward/redirect, flagged external |
| `MailGroups`         | Exchange              | `/groups` (DLs / mail-enabled / M365) |
| `MX-Records`         | Exchange              | DNS MX lookup per verified domain (inbound mail routing / 3rd-party filter) |
| `PublicFolders`      | Exchange              | Exchange Online — folder tree, item count/size, mail-enabled SMTP/proxy addresses |
| `PublicFolder-Permissions` | Exchange        | Exchange Online — per-folder client permissions (Owner/Editor/Reviewer/etc.) |
| `SharePoint`         | SharePoint            | `getSharePointSiteUsageDetail` report (incl. sites with no owner) |
| `SharePoint-Sharing` | SharePoint            | `/admin/sharepoint/settings` — tenant-wide external sharing capability/domain list |
| `OneDrive`           | SharePoint            | `getOneDriveUsageAccountDetail` report |
| `Teams-Groups`       | Teams                 | `/groups` (Unified) + owner/member counts |

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Modules (auto-installed to CurrentUser if missing): `Microsoft.Graph.Authentication`,
  `ImportExcel`, and `ExchangeOnlineManagement` (only when the Exchange workload runs
  without `-SkipMailboxType`, or without `-SkipPublicFolders`)
- Graph permissions: `Organization.Read.All`, `User.Read.All`, `Group.Read.All`,
  `GroupMember.Read.All`, `Directory.Read.All`, `Sites.Read.All`, `Reports.Read.All`,
  `AuditLog.Read.All`, `Domain.Read.All`, `RoleManagement.Read.Directory`,
  `SharePointTenantSettings.Read.All`

## Usage

By default the report is written to **`Documents\M365 Tenant Reports\`** and named
`<TenantName>-TenantReport-<date>.xlsx`. The Documents location is resolved per-user at
runtime (and follows OneDrive's redirected Documents folder if Known Folder Move is enabled),
so the script is portable across machines and users — no paths to edit. Override with
`-OutputPath <folder>` or `-OutputPath <full-path.xlsx>`.

```powershell
# Interactive, all workloads -> Documents\M365 Tenant Reports\<Tenant>-TenantReport-<date>.xlsx
.\New-M365TenantReport.ps1

# Pick workloads + 90-day usage window, custom output folder
.\New-M365TenantReport.ps1 -OutputPath C:\Reports -Workload Identity,Exchange -UsagePeriod D90

# Unattended (app registration with a certificate)
.\New-M365TenantReport.ps1 -UseAppOnly -TenantId contoso.onmicrosoft.com `
    -ClientId <app-guid> -CertificateThumbprint <thumbprint> -OutputPath C:\Reports
```

Parameters: `-OutputPath`, `-TenantId`, `-Workload` (Identity/Exchange/SharePoint/Teams),
`-UsagePeriod` (D7/D30/D90/D180), `-UseAppOnly` / `-ClientId` / `-CertificateThumbprint`,
`-SkipModuleInstall`, `-SkipMailboxType`, `-SkipDelegation`, `-SkipMailForwarding`,
`-SkipPublicFolders`, `-SkipPublicFolderPermissions`, `-OfflineSkuNames`, `-SkuCatalogPath`,
`-SkuCatalogUrl`, `-SkipUpdateCheck`, `-AutoUpdate`, `-UpdateToken`.

## Self-update

On start the script compares its `$ScriptVersion` against the latest published
[GitHub Release](https://github.com/Zanrose/office365-scripts/releases) (falling back
to the `main` branch if no releases exist). If a newer version exists it prints the
release notes (or `CHANGELOG-TenantReport.md` if there's no release yet) and asks before
downloading and replacing itself. The current file is backed up to `<script>.v<old>.bak`
first, and the download is syntax-checked before it's written, so a bad download can't
clobber a working script. After an update it exits and asks you to re-run.

- `-SkipUpdateCheck` — never check (e.g. offline / locked-down environments).
- `-AutoUpdate` — apply a newer version without prompting (for scheduled/unattended runs).
- `-UpdateToken <PAT>` — a GitHub token, only needed against a private fork. If the GitHub
  CLI (`gh`) is installed and signed in, the script uses that automatically.

## License friendly names

Graph only returns a license's `skuId` (GUID) and `skuPartNumber` (e.g. `SPE_E5`) —
never the marketing name ("Microsoft 365 E5"). To get full names the script downloads
Microsoft's published **"Product names and service plan identifiers for licensing"** CSV
(~620 SKUs) at runtime and resolves each license: catalog by GUID → catalog by part
number → built-in static map → raw part number. No tenant permissions are involved; it's
a public download.

- `-OfflineSkuNames` — skip the download and use only the small built-in map (air-gapped /
  no internet). Unmapped SKUs show their `skuPartNumber`.
- `-SkuCatalogPath <file.csv>` — use a previously saved copy of the CSV instead of downloading.
- `-SkuCatalogUrl <url>` — override the download URL if Microsoft moves the file.

If the download fails for any reason the script logs a warning and automatically falls back
to the built-in map — it never aborts the run over license names.

## Important notes / known limitations

- **Concealed report names.** If *Admin center > Settings > Org settings > Reports >
  "Display concealed user, group and site names"* is **ON**, the usage reports return
  blank/GUID names instead of UPNs and site URLs. Turn it off before running for full detail.
- **Mailbox types.** The Graph usage report can't tell a Shared/Room/Equipment mailbox from a
  user mailbox, so the Exchange workload also connects to **Exchange Online** (`Get-EXOMailbox`)
  and fills in `MailboxType` + `IsShared` on the Mailboxes sheet (and a Shared/Room/Equipment
  count on the Summary). This adds the `ExchangeOnlineManagement` module and a second sign-in.
  Pass `-SkipMailboxType` to skip it and stay pure-Graph — `MailboxType` then reads
  "Unknown (skipped)". If the ExO connection fails, the run continues and the type reads
  "Unknown".
- **Delegation.** FullAccess / SendAs / SendOnBehalf grants are inventoried per mailbox
  (explicit grants only — inherited rights and `NT AUTHORITY\SELF` are excluded). Skipped
  with `-SkipDelegation`; needs the same ExO connection as mailbox type.
  - **AutoMapping** column: Exchange Online has no supported way to read the AutoMapping
    flag for an *existing* FullAccess grant (it's stored in the mailbox's
    `msExchDelegateListLink` attribute, which `Get-Mailbox`/`Get-EXOMailbox` don't return).
    The column therefore reflects Microsoft's documented default — individual FullAccess
    grants automap (`On (default)`), SendAs/SendOnBehalf are `N/A`. A grant created with
    `-AutoMapping $false`, and FullAccess granted to a group (which never automaps), can't
    be distinguished here.
- **Mail forwarding.** Combines mailbox-level `ForwardingSmtpAddress`/`ForwardingAddress`
  with a per-mailbox `Get-InboxRule` scan for forward/forward-as-attachment/redirect rules,
  and flags any target outside the tenant's verified domains as `External = True`. Inbox
  rules resolve internal recipients to a display name rather than an address; when a target
  can't be resolved to an address at all, it's reported with `External = $null` (shown up on
  the Summary as "needing manual review") rather than guessed at. Skipped with
  `-SkipMailForwarding`.
- **Public folders.** The Exchange workload also connects to Exchange Online to walk the
  public folder tree (`Get-PublicFolder -Recurse`), pull per-folder item count/size
  (`Get-PublicFolderStatistics`), and match mail-enabled folders to their SMTP/proxy
  addresses (`Get-MailPublicFolder`). Client permissions (Owner/Editor/Reviewer/etc.) are
  collected per folder into a separate `PublicFolder-Permissions` sheet. If the tenant has
  public folders disabled, both sheets are skipped and the Summary notes "Public folders
  enabled: No". Use `-SkipPublicFolders` to skip both sheets entirely, or
  `-SkipPublicFolderPermissions` to keep the folder inventory but skip the (slower,
  one-call-per-folder) permissions sheet. Unlike Delegation, public folder discovery runs
  even when `-SkipMailboxType` is used, since it only needs the Exchange Online connection,
  not the mailbox list.
- **MX / inbound filtering.** Each verified (non-`onmicrosoft.com`) domain is MX-resolved.
  If the lowest-preference host isn't `*.mail.protection.outlook.com`, `ThirdPartyInbound`
  is `True` and `Provider` names the likely gateway (Mimecast, Proofpoint, Barracuda, Cisco
  IronPort, Google, etc.). Requires outbound DNS from the machine running the script.
- **Federated domains.** The `Domains` sheet reports each domain's `AuthenticationType`
  (`Managed` vs `Federated`) with the federation issuer / passive sign-in URI for federated
  ones, flagged on the Summary too — sign-in for a federated domain depends on an external
  IdP being reachable.
- **Privileged roles.** `/directoryRoles` only returns roles that have been activated in the
  tenant, so the sheet lists real assignments rather than every possible role. Members can be
  users, role-assignable groups, or service principals (see `PrincipalType`).
- **SharePoint external sharing.** `SharePoint-Sharing` reports the tenant-wide setting only
  (not per-site overrides). If the tenant or the granted permissions don't support the
  `/admin/sharepoint/settings` endpoint, the sheet comes back empty rather than failing the run.
- **Usage data lag.** Microsoft usage reports are typically 1–2 days behind real time.
- One failing workload won't abort the run — it logs an error and writes an empty sheet.

## Suggested next steps

- Conditional Access policy export (`/identity/conditionalAccess/policies`) and Security
  Defaults status — the two biggest levers on sign-in security, and neither is covered yet.
- Legacy/basic auth usage (sign-in logs), since it bypasses Conditional Access entirely.
- Intune-managed device inventory (count, OS, compliance state) — the last major workload
  with no coverage in this report.
- Mail flow (transport) rules — outbound-side counterpart to the inbound MX/filtering check.
- SPF/DKIM/DMARC records alongside the MX pass.
- Per-site SharePoint sharing overrides, beyond the tenant-wide default.

## License

MIT — see [LICENSE](LICENSE). Free to use, modify, and redistribute with attribution.
