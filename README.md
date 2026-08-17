# Nutanix Prism Central Inventory — PowerShell

PowerShell port of the `cluster_inventory_v4` Python tool. Connects to one or
more Prism Centrals over the **v4 REST API** and writes a formatted, timestamped
Excel workbook per PC.

Feature parity with the Python version: same sheets, same columns, same derived
values and the same colour coding.

## Requirements

| | |
|---|---|
| PowerShell | 7.0+ recommended, 5.1 supported |
| Module | [`ImportExcel`](https://github.com/dfinke/ImportExcel) (only dependency) |
| Access | Prism Central reachable on port 9440, an account with read access |

There is **no dependency on the Nutanix PowerShell/Python SDKs** — every call is
a direct REST request.

> On PowerShell 7 the per-VM statistics are collected with a throttled parallel
> pool. Windows PowerShell 5.1 has no equivalent, so it falls back to a serial
> loop and will be noticeably slower on large estates.

### Installing ImportExcel

Online:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

Offline (air-gapped host) — on a machine with internet access:

```powershell
Save-Module ImportExcel -Path .\modules
```

Copy the `modules\ImportExcel` folder across, then either drop it into a folder
listed in `$env:PSModulePath`, or import it explicitly before running:

```powershell
Import-Module .\modules\ImportExcel\<version>\ImportExcel.psd1
```

## Usage

### 1. Interactive

```powershell
.\Invoke-NtnxInventory.ps1
```

Prompts for the Prism Central endpoint, authentication and cluster scope, and
lets you queue several PCs in one run.

### 2. Direct

```powershell
.\Invoke-NtnxInventory.ps1 -PcIp 10.38.99.39 -PcUser admin -PcSecret 'password'
```

Multiple PCs and a cluster subset:

```powershell
.\Invoke-NtnxInventory.ps1 -PcIp 10.38.99.39,10.38.99.40 `
                           -PcUser admin -PcSecret 'password' `
                           -Clusters 'PROD-CL01','PROD-CL02' `
                           -OutputPath D:\Reports
```

### 3. Credentials file

```powershell
.\Invoke-NtnxInventory.ps1 -CredentialsFile .\pcs.csv
```

`pcs.csv` requires `pc_ip` and `pc_user`; `pc_secret`, `auth_mode`, `clusters`
and `api_path_version` are optional:

```csv
pc_ip,pc_user,pc_secret,auth_mode,clusters
10.38.99.39,admin,password,basic,ALL
10.38.99.40,admin,password,basic,"PROD-CL01,PROD-CL02"
```

### Credentials from the environment

If `-PcIp` / `-PcUser` / `-PcSecret` are omitted, the script falls back to
`PRISM_CENTRAL_VIP`, `PRISM_CENTRAL_USERNAME` and `PRISM_CENTRAL_PASSWORD`. This
matches the pattern used by the Ansible projects, so secrets can be injected at
runtime rather than stored on disk.

```powershell
$env:PRISM_CENTRAL_VIP      = '10.38.99.39'
$env:PRISM_CENTRAL_USERNAME = 'admin'
$env:PRISM_CENTRAL_PASSWORD = 'password'
.\Invoke-NtnxInventory.ps1
```

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-PcIp` | — | One or more PCs. Accepts `host` or `host:port`. |
| `-PcUser` | — | Username, or the API key when `-AuthMode api_key`. |
| `-PcSecret` | — | Password. Required for basic auth. |
| `-AuthMode` | `basic` | `basic` or `api_key`. |
| `-Clusters` | `ALL` | Cluster names to include. |
| `-CredentialsFile` | — | CSV or XLSX describing multiple PCs. |
| `-OutputPath` | `.\output` | Workbook destination. |
| `-UnderProvisionPercent` | `75` | Usage above this is shaded red. |
| `-OverProvisionPercent` | `30` | Usage below this is shaded amber. |
| `-ApiPathVersion` | `auto` | Pin a v4 minor (`v4.0`) instead of probing. |
| `-Interactive` | — | Force the interactive menu. |

### API version negotiation

By default each namespace is probed once and pinned to the minor version the PC
actually serves. This replaces the `sdk_compat.py` workaround in the Python
version, which existed because the SDKs assumed a newer minor than older PC
builds expose. Namespaces the PC does not serve at all are reported and their
sheets are written empty rather than failing the run.

## Output

`Nutanix_PC_Inventory_<pc-name>_<YYYY-MM-DD_HH-MM-SS>.xlsx`, containing:

| # | Sheet | Contents |
|---|---|---|
| 1 | Summary | KPI tiles and a per-cluster snapshot |
| 2 | Cluster Overview | Full cluster configuration and health |
| 3 | Hypervisor Hosts | Host hardware, CPU/memory/storage usage |
| 4 | Host Disks | Physical disks per host |
| 5 | Virtual Machines | VM config plus 24h CPU/memory/disk usage |
| 6 | Storage Containers | Capacity, RF, compression, encryption |
| 7 | Networks | Subnets and IP pool usage |
| 8 | Virtual Switches | vSwitch config per cluster |
| 9 | Recovery Points | Snapshots (newest first, capped at 1000) |
| 10 | Protection Policies | DR policies, RPO and retention |
| 11 | Alerts | Open critical and warning alerts |
| 12 | Licenses | Entitlements and expiry |

When more than one cluster is processed, per-cluster `VMs_<cluster>` and
`Hosts_<cluster>` sheets are added.

Usage cells above `-UnderProvisionPercent` are shaded red and those below
`-OverProvisionPercent` amber, so under- and over-provisioned VMs stand out.

Exit code is `0` when every PC succeeded and `1` if any failed.

## Performance notes

- VM statistics dominate the runtime. Powered-off VMs are skipped entirely.
- Concurrency is deliberately capped at 3 workers; Prism Central rate-limits
  these endpoints and raising the limit reduces throughput. HTTP 429/503 is
  retried with exponential backoff.
- Category lookups are memoised per PC.
- Budget roughly 10–25 minutes for a 10,000-VM estate on PowerShell 7.

## Layout

```
Invoke-NtnxInventory.ps1        Entry point: CLI, interactive, multi-PC
NtnxPcInventory/
  Common.ps1                    Logging, unit conversion, safe field access
  NtnxRest.ps1                  Auth, TLS, pagination, retry, version probing
  Collector.ps1                 API calls and row shaping
  ExcelExport.ps1               Workbook builder
```
