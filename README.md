# Halo-ITSM-API-Scripts

PowerShell scripts for working with projects and project tasks in
[Halo ITSM](https://haloitsm.com) over its REST API.

Each script is self-contained: no modules, no parameters, no config files.
Open one in PowerShell ISE, fill in the `CONFIGURATION` block at the top,
press F5.

## Which script do I want?

| Script | Direction | What it does |
| --- | --- | --- |
| [`New-HaloProject.ps1`](New-HaloProject.ps1) | **writes** | Creates one project and its tasks from a table you fill in. |
| [`New-HaloProjectReport.ps1`](New-HaloProjectReport.ps1) | read-only | HTML report across *every* project in the tenant. |
| [`New-HaloProjectStatusReport.ps1`](New-HaloProjectStatusReport.ps1) | read-only | HTML status report for *one* project. |

## Configuration

Every script follows the same layout. One `CONFIGURATION` block sits at the
top, and it ends at a line you cannot miss:

```
# ===========================================================================
#  NOTHING BELOW HERE NEEDS EDITING
# ===========================================================================
```

Everything above that line is yours to change; everything below it is
machinery. The block is divided into labelled sections, in the order you need
them:

| Section | What is in it |
| --- | --- |
| `--- The only switch you need` | `New-HaloProject.ps1` only: `$Apply`, dry run versus create. |
| `--- Connection` | `$Tenant`, `$ClientId`, `$ClientSecret`, `$Scope`, and `$AuthUrl` / `$ApiBase` for self-hosted Halo. |
| `--- Your tenant's ids` | The project and task ticket type ids. |
| *script-specific* | What goes in the report, which project, the project and its tasks. |
| `--- Agents who have left` | `$FormerAgents`, `$FormerSuffix` (reports only). |
| `--- How the RAG verdict is judged` | Thresholds and status-name lists (status report only). |
| `--- Advanced` | `$PAGE`, the API page size. Rarely touched. |

The variables you are most likely to change, per script:

| Script | Settings |
| --- | --- |
| `New-HaloProject.ps1` | `$Apply`, `$ProjectTypeId`, `$TaskTypeId`, `$PriorityMap`, `$AgentMap`, `$ProjectSummary`, `$ProjectStart`, `$ProjectTarget`, `$ProjectBlurb`, `$SourceNote`, `$Records`, `$ExistingProjectId`, `$OnlyIds` |
| `New-HaloProjectReport.ps1` | `$TYPE_PROJECT`, `$TYPE_TASK`, `$OutFile`, `$IncludeClosed`, `$OpenWhenDone`, `$AgentFilter`, `$FormerAgents` |
| `New-HaloProjectStatusReport.ps1` | `$TYPE_PROJECT`, `$TYPE_TASK`, `$ProjectId`, `$OutFile`, `$OpenWhenDone`, `$FormerAgents`, `$NearEndPct`, `$NearEndDays`, `$DueSoonDays`, the three status-name lists, `$LateInProgressIsRed` |

---

### New-HaloProject.ps1

A template for creating a project and the tasks that hang off it. You describe
the tasks as a table of records at the top of the file — id, summary,
workstream, owner, priority, dates, dependencies, description, acceptance test
— and the script creates the project ticket and one task per record beneath it.

It ships with six example tasks demonstrating the shape. Delete them and write
your own; the header documents every field.

This is the only script here that writes to Halo, and there is no undo, so:

- `$Apply` defaults to `$false`. That is a **dry run** — it authenticates,
  reads your ticket types, priorities and agents, validates the record table,
  prints everything it would create, and writes nothing. Run it, read it, then
  set `$Apply = $true`.
- The record table is validated before the token is even requested: duplicate
  ids, dates that are not `yyyy-MM-dd`, a target before its start, tasks
  outside the project window, an `agentkey` or `priority` that is not in the
  maps, and a dependency naming an id that does not exist. A typo costs you
  nothing.
- Re-running is safe. It reads the tasks already on the project and skips any
  whose summary it recognises, so a partial run can simply be run again. Put
  the project's id in `$ExistingProjectId` to add to an existing project.
- After creating the first task it checks that the task actually attached to
  the project, and stops the whole run if it did not — one stray record to
  clean up rather than thirty.
- After creating everything it re-reads every new ticket and reports any whose
  dates came back different from what was sent, which is usually a working
  calendar or an SLA on the ticket type.

### New-HaloProjectReport.ps1

Reads every project and project task in the tenant, open and closed, and writes
one self-contained HTML file:

- a portfolio overview — counts, delivery load by month and by owner, workload
  per agent, everything with a date problem, projects finishing per quarter
- a section per project — its own window, its rolled-up window, its milestones,
  and every task with a timeline bar

The report has filters, a search box and a dark mode that follows the machine.
It has no external dependencies, so you can mail it, print it or drop it on a
share. Defaults to `Documents\Halo-Projects-<timestamp>.html`.

This one is **slow**: the Halo ticket list omits `startdate`, so every record
needs its own GET — roughly a second per six records. Budget accordingly on a
tenant with a lot of history.

### New-HaloProjectStatusReport.ps1

The same idea narrowed to a single project, which makes it quick — give it a
ticket id and it goes straight there instead of sweeping the tenant. Set
`$ProjectId` to a number to go direct, to text to search project names, or
leave it blank to be asked.

- a RAG verdict with the reasoning printed beside it, never a bare colour
- progress bars for work completed (weighted by how long each task runs, not a
  flat headcount) against schedule elapsed
- who is assigned and what each of them is carrying
- milestones with their own completion
- every task, open and closed, with dates and a timeline bar
- what is overdue, what lands in the next fortnight, and every broken date

Defaults to `Documents\Halo-Project-<id>-<name>-<timestamp>.html`.

---

## Getting started

### 1. Create an API application in Halo

**Configuration > Integrations > Halo API > View Applications > New**

- Authentication Method: **Client ID and Secret (Services)**
- Login Type: **Agent** — and make sure an agent is actually selected
  underneath, not just the login type
- Permissions: `edit:tickets` for `New-HaloProject.ps1`. The two reports use
  scope `all`, because `/api/Agent` and `/api/Status` can 403 on anything
  narrower.

### 2. Set your tenant

Every script ships with the placeholder `$Tenant = 'contoso'`, the first line
of the `--- Connection` section. Change it to your own tenant name, which is
the first part of your Halo URL:

```powershell
$Tenant = 'yourtenant'      # https://yourtenant.haloitsm.com
```

`$AuthUrl` and `$ApiBase` are built from `$Tenant` while they are left blank.
Fill them in only if Halo is self-hosted, or if the **Authorisation Server**
shown on the Halo API page is not `https://<tenant>.haloitsm.com/auth/token`:

```powershell
$AuthUrl = 'https://halo.example.com/auth/token'
$ApiBase = 'https://halo.example.com/api'
```

### 3. Supply credentials

All three scripts look in the same three places, in order:

1. `$ClientId` / `$ClientSecret` in the `--- Connection` section
2. the `HALO_CLIENT_ID` and `HALO_CLIENT_SECRET` environment variables
3. a masked prompt

The credential fields are committed empty and should stay that way — use the
environment variables or the prompt so a secret never reaches the repository:

```powershell
$env:HALO_CLIENT_ID     = '...'
$env:HALO_CLIENT_SECRET = '...'
```

## Before you run these against your own tenant

These scripts were written against one tenant, and a few values in them are
specific to it. Check each one:

- **Ticket type ids.** Under `--- Your tenant's ids`, the reports
  default to `$TYPE_PROJECT = 57` and `$TYPE_TASK = 58`. Yours will differ.
  `New-HaloProject.ps1` is better behaved: leave `$ProjectTypeId` and
  `$TaskTypeId` at `0` and its dry run prints every ticket type it can see,
  with ids — then paste them into all three scripts. Match on **id**, never on
  name, because the display names carry stray whitespace.
- **Priority and agent ids.** Same story: the dry run prints them, and you
  paste them into `$PriorityMap` and `$AgentMap`.
- **Status names.** `New-HaloProjectStatusReport.ps1` decides whether a task is
  not started, on hold or in progress by matching status names, which are
  tenant-specific. It prints the mapping it used and names anything it could
  not place. Read that once, adjust the three lists under
  `--- How the RAG verdict is judged`, and it stays right from then on.
- **Departed agents.** Deactivated agents vanish from `/api/Agent`, so their
  tickets show as `id:NN`. Name them in `$FormerAgents` in either report and
  they read properly throughout.

## Requirements

- Windows PowerShell 5.1 (ISE) or PowerShell 7
- Network access to your Halo tenant, and an API application as above

## Notes on the Halo API

Things that cost time to work out, collected here so they are not learned twice.
Each script's header has the longer version.

- `count=N` is capped at **1000** and returns the *newest* 1000 while still
  reporting `record_count=1000`, so a naive completeness check passes while
  most of the history is missing. Page through; never trust the count.
- The ticket **list** omits `startdate`. Anything that needs start dates needs
  an individual GET per record, which is what makes the portfolio report slow.
- Milestones ride on the **project** ticket as a `milestones[]` array, using
  underscored `start_date` / `target_date`. A task points at one through
  `milestone_id`.
- `1900-01-01` is Halo's "unset" date, not a real one.
- `POST /api/Tickets` takes an **array**, and a payload with no `id` means
  create. A task is attached to its project through `parent_id`.
- A 400 carries a JSON body naming the actual problem (`invalid_client`,
  `invalid_scope`, `unsupported_grant_type`). PowerShell hides it behind a
  generic message — dig it out of `ErrorDetails.Message` or the response
  stream. A *bad tenant* answers with a full HTML error page instead, several
  hundred kilobytes of it, so truncate before printing.
- Dates come back as JSON strings but deserialise into `[datetime]`, whose
  default string form follows the machine locale. Normalise to `yyyy-MM-dd`
  before comparing anything, and parse authored dates with `TryParseExact` and
  the invariant culture, or a machine set to `dd/MM/yyyy` will reinterpret them.
