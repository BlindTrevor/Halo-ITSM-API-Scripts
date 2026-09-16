<#
    New-HaloProjectReport.ps1
    =============================
    READ-ONLY. Reads every project and project task out of Halo - open and
    closed - and writes a single self-contained HTML report:

        * a portfolio overview  - counts, delivery load by month and owner,
                                  workload per agent, everything with a date
                                  problem, projects finishing per quarter
        * a section per project - its own window, its rolled-up window, its
                                  milestones, and every task with a timeline bar

    Open in PowerShell ISE, paste credentials, press F5. Writes nothing to
    Halo and nothing to disk except the report.

    The report is one file with no external dependencies - mail it, print it,
    put it on a share. It has filters, a search box, and a dark mode that
    follows the machine.

    Notes about this tenant, learned the hard way:
      * count=N is capped at 1000 and returns the NEWEST 1000, then reports
        record_count=1000 - so a naive completeness check passes while most
        of the history is missing. We page, and cross-check against the
        open-only sweep, and refuse to write a report if records go missing.
      * The ticket list omits startdate, so every record needs an individual
        GET. That is what makes this slow - roughly one second per 6 records.
      * Project = tickettype_id 57 (its display name has a trailing space),
        Project Task = 58. Match on id, never on name.
      * Milestones ride on the PROJECT ticket as a milestones[] array, with
        underscored start_date / target_date. A task points at one through
        milestone_id.
      * 1900-01-01 is Halo's "unset" date.
      * Deactivated agents are absent from /api/Agent, so their tickets show
        as "id:NN" rather than a name.
#>

# ===========================================================================
#  CONFIGURATION
#  -------------------------------------------------------------------------
#  Everything you may need to change is in this one block. Work down it:
#  connection first, then the two ticket type ids, then taste.
# ===========================================================================

# --- Connection -------------------------------------------------------------

$Tenant       = 'contoso'  # first part of your Halo URL: https://<tenant>.haloitsm.com
$ClientId     = ''         # blank = $env:HALO_CLIENT_ID, then a masked prompt
$ClientSecret = ''         # blank = $env:HALO_CLIENT_SECRET, then a masked prompt
$Scope        = 'all'      # /api/Agent and /api/Status may 403 on anything narrower

# Only needed if Halo is self-hosted, or if the "Authorisation Server" shown
# in Configuration > Integrations > Halo API is not the address below.
# Leave both blank to build them from $Tenant.
$AuthUrl      = ''         # e.g. 'https://halo.example.com/auth/token'
$ApiBase      = ''         # e.g. 'https://halo.example.com/api'

# --- Your tenant's ids ------------------------------------------------------
# These differ per tenant and there is no safe default. Halo's display names
# carry stray whitespace, so match on id and never on name. A dry run of
# New-HaloProject.ps1 prints every ticket type it can see, with ids.

$TYPE_PROJECT = 57
$TYPE_TASK    = 58

# --- What goes in the report ------------------------------------------------

$OutFile       = ''        # blank = Documents\Halo-Projects-<timestamp>.html
$IncludeClosed = $true     # closed records matter - they set rolled-up starts
$OpenWhenDone  = $true     # launch the report in the default browser
$AgentFilter   = '*'       # '*' = whole team, or e.g. 'Jane Smith'

# --- Agents who have left ---------------------------------------------------
# Deactivated agents vanish from /api/Agent, so their tickets come back as
# "id:NN". Name them here and they read properly everywhere - chart, workload
# table, owner filter and every task row.
# NOTE: bare numeric keys in a hashtable literal are INTEGERS. Look them up
# with an int, never with "$id", or the lookup silently never matches.

$FormerAgents = @{
    13 = 'Joe Bloggs'
}
$FormerSuffix  = ' (left)'   # set to '' to show the name with no marker

# --- Advanced ---------------------------------------------------------------

$PAGE = 100                # records per API page. 1000 is Halo's hard cap, and
                           # it returns the NEWEST 1000 while still reporting
                           # record_count=1000 - so page, never count.

# ===========================================================================
#  NOTHING BELOW HERE NEEDS EDITING
# ===========================================================================

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($AuthUrl)) { $AuthUrl = "https://$Tenant.haloitsm.com/auth/token" }
if ([string]::IsNullOrWhiteSpace($ApiBase)) { $ApiBase = "https://$Tenant.haloitsm.com/api" }

# ---------------------------------------------------------------------------
#  HELPERS
# ---------------------------------------------------------------------------

function Resolve-Credential {
    param([string]$Inline, [string]$EnvName, [string]$Prompt)
    if (-not [string]::IsNullOrWhiteSpace($Inline)) { return $Inline.Trim() }
    $e = [Environment]::GetEnvironmentVariable($EnvName)
    if (-not [string]::IsNullOrWhiteSpace($e)) { return $e.Trim() }
    $s = Read-Host -Prompt $Prompt -AsSecureString
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
    try { return ([Runtime.InteropServices.Marshal]::PtrToStringAuto($b)).Trim() }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}

function Get-Val {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-Collection {
    # A Halo endpoint returns either a bare array or an object wrapping one.
    param($Resp, [string]$Name)
    if ($null -eq $Resp) { return @() }
    if ($Resp -is [System.Array]) { return @($Resp) }
    $p = $Resp.PSObject.Properties[$Name]
    if ($p -and $null -ne $p.Value) { return @($p.Value) }
    foreach ($n in @('data','records','items')) {
        $q = $Resp.PSObject.Properties[$n]
        if ($q -and $null -ne $q.Value) { return @($q.Value) }
    }
    return @($Resp)
}

function Format-D {
    # Invoke-RestMethod deserialises JSON dates into [datetime]; their string
    # form follows machine locale, so a UK box yields 24/08/2026. Normalise, or
    # every comparison and every sort is wrong.
    param($v)
    if ($null -eq $v) { return '' }
    if ($v -is [datetime]) {
        if ($v.Year -le 1900) { return '' }
        return $v.ToString('yyyy-MM-dd')
    }
    $s = [string]$v
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    if ($s.StartsWith('1900-01-01') -or $s.StartsWith('1899-') -or $s.StartsWith('0001-')) { return '' }
    if ($s.Length -ge 10 -and $s[4] -eq '-') { return $s.Substring(0,10) }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$d)) {
        if ($d.Year -le 1900) { return '' }
        return $d.ToString('yyyy-MM-dd')
    }
    return ''
}

function Show-D {
    param([string]$iso)
    if ([string]::IsNullOrWhiteSpace($iso)) { return '' }
    return ($iso.Substring(8,2) + '/' + $iso.Substring(5,2) + '/' + $iso.Substring(0,4))
}

function Esc {
    param($t)
    if ($null -eq $t) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$t)
}

function Format-ErrorBody {
    # A bad tenant or a server-side fault answers with a whole error PAGE - a
    # quarter of a megabyte of CSS and inline base64 images, which buries the
    # console. The useful case is the small JSON body Halo returns for a bad
    # credential, so keep that intact and reduce everything else to the first
    # readable sentence.
    param([string]$Body)
    if ([string]::IsNullOrWhiteSpace($Body)) { return '' }
    $s = $Body.Trim()

    if (-not ($s.StartsWith('{') -or $s.StartsWith('['))) {
        # Not JSON, so it is a page. Note that Invoke-RestMethod has usually
        # stripped the tags already, leaving bare stylesheet text behind.
        $s = $s -replace '(?s)<(script|style)\b.*?(</\1>|$)', ' '
        $s = $s -replace '(?s)<[^>]+>', ' '
        $s = [System.Net.WebUtility]::HtmlDecode($s)
        $s = $s -replace 'data:[^;,\s]*;base64,[A-Za-z0-9+/=]*', '[embedded data]'
        $s = $s -replace '(?s)\{[^{}]*\}', ' '          # css rules
        $s = $s -replace '[A-Za-z0-9+/=]{80,}', '[blob]'
    }

    $s = ($s -replace '\s+', ' ').Trim()
    if ($s.Length -gt 400) { $s = $s.Substring(0, 400) + '... (truncated)' }
    return $s
}

function Get-WebErrorBody {
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return Format-ErrorBody $ErrorRecord.ErrorDetails.Message
    }
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($null -eq $resp) { return '' }
        $st = $resp.GetResponseStream()
        try { $st.Position = 0 } catch { }
        return Format-ErrorBody (New-Object System.IO.StreamReader($st)).ReadToEnd()
    } catch { return '' }
}

function Say  { param([string]$m,[string]$c='Gray') Write-Host $m -ForegroundColor $c }
function Head {
    param([string]$t)
    Write-Host ''
    Write-Host ("  " + $t) -ForegroundColor Cyan
    Write-Host ("  " + ('-' * 100)) -ForegroundColor DarkGray
}

function Set-Overall {
    param([int]$Step,[int]$Of,[string]$Name)
    Write-Progress -Id 0 -Activity 'Halo project report' -Status "Step $Step of $Of - $Name" `
                   -PercentComplete ([int](($Step - 1) / $Of * 100))
}
function Set-Step {
    param([string]$Activity,[int]$I,[int]$N,[string]$Status='')
    if ($N -le 0) { return }
    $s = $Status
    if ([string]::IsNullOrWhiteSpace($s)) { $s = "$I of $N" }
    Write-Progress -Id 1 -ParentId 0 -Activity $Activity -Status $s `
                   -PercentComplete ([int]([Math]::Min(100, $I / $N * 100)))
}
function Clear-Step { Write-Progress -Id 1 -Activity ' ' -Completed }
function Clear-All  { Write-Progress -Id 1 -Activity ' ' -Completed; Write-Progress -Id 0 -Activity ' ' -Completed }

function Get-Ticket {
    param([int]$Id)
    try { return Invoke-RestMethod -Uri "$ApiBase/Tickets/$Id" -Headers $script:Headers -Method Get -TimeoutSec 120 }
    catch { return $null }
}

function Get-AllTickets {
    # Page the whole history, then merge the open-only sweep back in as a
    # completeness check. The open-only sweep is under the 1000 cap, so any
    # open record missing from the paged result means the paging is broken.
    $byId = @{}; $openIds = @{}
    Say '      open-only sweep...' 'DarkGray'
    try {
        $r = Invoke-RestMethod -Uri "$ApiBase/Tickets?count=25000&open_only=true" `
             -Headers $script:Headers -Method Get -TimeoutSec 600
        foreach ($t in (Get-Collection $r 'tickets')) {
            if ($null -ne $t.id) { $byId[[int]$t.id] = $t; $openIds[[int]$t.id] = $true }
        }
    } catch { Say '      open-only sweep failed' 'Yellow' }
    Say "      open sweep gave $($openIds.Count) ticket(s) - paging the full history..." 'DarkGray'

    $pg = 1; $total = $null; $fetched = 0
    while ($true) {
        try {
            $r = Invoke-RestMethod -Uri "$ApiBase/Tickets?pageinate=true&page_size=$PAGE&page_no=$pg" `
                 -Headers $script:Headers -Method Get -TimeoutSec 600
        }
        catch { break }
        $b = Get-Collection $r 'tickets'
        if ($null -eq $total) {
            $rc = Get-Val $r 'record_count'
            if ($null -ne $rc) { $total = [int]$rc }
        }
        if ($b.Count -eq 0) { break }
        $fetched += $b.Count
        foreach ($t in $b) {
            if ($null -ne $t.id -and -not $byId.ContainsKey([int]$t.id)) { $byId[[int]$t.id] = $t }
        }
        $pct = 1
        if ($total -and $total -gt 0) { $pct = [int]([Math]::Min(100, $fetched / $total * 100)) }
        Write-Progress -Id 1 -ParentId 0 -Activity 'Reading the ticket list' `
                       -Status "page $pg - $fetched of $total" -PercentComplete $pct
        if ($pg % 10 -eq 0) { Say "      page $pg - $fetched of $total" 'DarkGray' }
        if ($b.Count -lt $PAGE) { break }
        if ($null -ne $total -and $fetched -ge $total) { break }
        $pg++
        if ($pg -gt 800) { break }
    }
    Clear-Step
    $lost = @($openIds.Keys | Where-Object { -not $byId.ContainsKey($_) })
    return @{ tickets = @($byId.Values); total = $total; fetched = $fetched; lost = $lost.Count }
}

# ---------------------------------------------------------------------------
#  RUN
# ---------------------------------------------------------------------------

Write-Host ''
Write-Host '  HALO PROJECT REPORT - read-only' -ForegroundColor Cyan
Write-Host "  tenant $Tenant   closed records: $(if($IncludeClosed){'included'}else{'excluded'})   agent: $AgentFilter"

$cid = Resolve-Credential -Inline $ClientId     -EnvName 'HALO_CLIENT_ID'     -Prompt 'Halo Client ID'
$sec = Resolve-Credential -Inline $ClientSecret -EnvName 'HALO_CLIENT_SECRET' -Prompt 'Halo Client Secret'

Set-Overall 1 5 'authenticating'
Say '  Requesting token...' 'DarkGray'
try {
    $token = (Invoke-RestMethod -Uri $AuthUrl -Method Post -Body @{
        grant_type='client_credentials'; client_id=$cid; client_secret=$sec
        scope=$Scope; tenant=$Tenant } -ContentType 'application/x-www-form-urlencoded').access_token
}
catch {
    Clear-All
    Say "  Auth failed: $($_.Exception.Message)" 'Red'
    $b = Get-WebErrorBody $_; if ($b) { Say "  Server said: $b" 'Red' }
    Say '  unauthorized_client means the Halo application Authentication Method is not' 'Yellow'
    Say '  "Client ID and Secret (Services)". invalid_client means a bad credential.' 'Yellow'
    return
}
$script:Headers = @{ Authorization = "Bearer $token" }

# --- lookups ---------------------------------------------------------------
$agentById = @{}; $statusById = @{}
try {
    foreach ($a in (Get-Collection (Invoke-RestMethod -Uri "$ApiBase/Agent" -Headers $script:Headers -Method Get) 'agents')) {
        if ($null -ne $a.id) { $agentById[[int]$a.id] = [string]$a.name }
    }
} catch { Say '  /api/Agent unavailable - agents will show as ids.' 'Yellow' }
try {
    foreach ($s in (Get-Collection (Invoke-RestMethod -Uri "$ApiBase/Status" -Headers $script:Headers -Method Get) 'statuses')) {
        if ($null -ne $s.id) { $statusById[[int]$s.id] = [string]$s.name }
    }
} catch { Say '  /api/Status unavailable - statuses will show as ids.' 'Yellow' }
Say "  resolved $($agentById.Count) agents, $($statusById.Count) statuses" 'DarkGray'

function Agent-Name {
    param($id)
    if ($null -eq $id) { return 'Unassigned' }
    $i = [int]$id
    if ($i -le 0) { return 'Unassigned' }
    if ($agentById.ContainsKey($i))    { return $agentById[$i] }
    if ($FormerAgents.ContainsKey($i)) { return ($FormerAgents[$i] + $FormerSuffix) }
    return "id:$i"
}
function Status-Name { param($id) if ($null -eq $id) { return '' }
    $i = [int]$id; if ($statusById.ContainsKey($i)) { return $statusById[$i] } ; return "id:$i" }

# --- the sweep -------------------------------------------------------------
Head 'READING'
Set-Overall 2 5 'reading the ticket list'
$all = Get-AllTickets
Say "  $($all.tickets.Count) ticket(s) in total; open records lost in the merge: $($all.lost)" `
    $(if ($all.lost -gt 0) { 'Red' } else { 'DarkGray' })
if ($all.lost -gt 0) {
    Clear-All
    Say '  STOPPING - the sweep is incomplete, so the report would understate the position.' 'Red'
    return
}

$shortlist = @($all.tickets | Where-Object {
    $tt = Get-Val $_ 'tickettype_id'
    $null -ne $tt -and ([int]$tt -eq $TYPE_PROJECT -or [int]$tt -eq $TYPE_TASK)
})
Say "  $($shortlist.Count) of those are projects or project tasks" 'DarkGray'

# The list payload has no startdate, so each one needs an individual read.
Set-Overall 3 5 'reading each record for its start date'
Say "  reading $($shortlist.Count) record(s) individually - this is the slow part" 'DarkGray'
$full = New-Object System.Collections.ArrayList
$i = 0; $n = $shortlist.Count; $sw = [Diagnostics.Stopwatch]::StartNew()
foreach ($t in $shortlist) {
    $i++
    if ($i % 5 -eq 0 -or $i -eq $n) {
        $rate = 0.0
        if ($sw.Elapsed.TotalSeconds -gt 0) { $rate = $i / $sw.Elapsed.TotalSeconds }
        $left = 0
        if ($rate -gt 0) { $left = [int](($n - $i) / $rate) }
        Set-Step 'Reading records' $i $n "$i of $n - about $left second(s) left"
    }
    $d = Get-Ticket -Id ([int]$t.id)
    if ($null -eq $d) { $d = $t }
    [void]$full.Add($d)
}
Clear-Step
Say "  read $($full.Count) record(s) in $([int]$sw.Elapsed.TotalSeconds)s" 'DarkGray'

# ---------------------------------------------------------------------------
#  BUILD THE MODEL
# ---------------------------------------------------------------------------
Head 'BUILDING'
Set-Overall 4 5 'building the model'

$OPEN_STATES = @{}          # ticket id -> $true when open
$projects = @{}             # project id -> hashtable
$tasksByParent = @{}        # parent id -> ArrayList of task hashtables
$orphans = New-Object System.Collections.ArrayList

function New-Row {
    param($t, [bool]$isProject)
    $sid  = Get-Val $t 'status_id'
    $name = Status-Name $sid
    $closedOn = Format-D (Get-Val $t 'dateclosed')
    $isClosed = ($name -match '^(Closed|Resolved|Cancelled|Completed)') -or ($closedOn -ne '')
    return @{
        id        = [int](Get-Val $t 'id')
        summary   = [string](Get-Val $t 'summary')
        agent     = Agent-Name (Get-Val $t 'agent_id')
        status    = $name
        statusId  = $(if ($null -eq $sid) { 0 } else { [int]$sid })
        start     = Format-D (Get-Val $t 'startdate')
        target    = Format-D (Get-Val $t 'targetdate')
        deadline  = Format-D (Get-Val $t 'deadlinedate')
        closedOn  = $closedOn
        closed    = $isClosed
        parentId  = $(if ($null -eq (Get-Val $t 'parent_id')) { 0 } else { [int](Get-Val $t 'parent_id') })
        milestone = $(if ($null -eq (Get-Val $t 'milestone_id')) { 0 } else { [int](Get-Val $t 'milestone_id') })
        client    = [string](Get-Val $t 'client_name')
        site      = [string](Get-Val $t 'site_name')
        details   = [string](Get-Val $t 'details')
        isProject = $isProject
        raw       = $t
    }
}

foreach ($t in $full) {
    $tt = [int](Get-Val $t 'tickettype_id')
    if ($tt -eq $TYPE_PROJECT) {
        $r = New-Row $t $true
        # milestones ride on the project ticket, with UNDERSCORED date fields
        $ms = New-Object System.Collections.ArrayList
        foreach ($m in @(Get-Val $t 'milestones')) {
            if ($null -eq $m) { continue }
            [void]$ms.Add(@{
                id  = [int](Get-Val $m 'id')
                name = [string](Get-Val $m 'name')
                seq  = $(if ($null -eq (Get-Val $m 'sequence')) { 0 } else { [int](Get-Val $m 'sequence') })
                start  = Format-D (Get-Val $m 'start_date')
                target = Format-D (Get-Val $m 'target_date')
            })
        }
        $r.milestones = $ms
        $r.tasks = New-Object System.Collections.ArrayList
        $projects[$r.id] = $r
    }
}
foreach ($t in $full) {
    $tt = [int](Get-Val $t 'tickettype_id')
    if ($tt -ne $TYPE_TASK) { continue }
    $r = New-Row $t $false
    if ($r.parentId -gt 0 -and $projects.ContainsKey($r.parentId)) {
        [void]$projects[$r.parentId].tasks.Add($r)
    } else {
        $r.parentName = [string](Get-Val $t 'parent_summary')
        [void]$orphans.Add($r)
    }
}
Say "  $($projects.Count) project(s), $(($projects.Values | ForEach-Object { $_.tasks.Count } | Measure-Object -Sum).Sum) task(s) under a project, $($orphans.Count) task(s) with no project" 'DarkGray'

# --- roll-ups, flags -------------------------------------------------------
$today = (Get-Date).ToString('yyyy-MM-dd')

function Get-Flags {
    param($r, [bool]$isClosed)
    $f = @()
    if ($isClosed) { return $f }
    if (-not $r.start)  { $f += 'no start' }
    if (-not $r.target) { $f += 'no end' }
    if ($r.start -and $r.target -and $r.target -lt $r.start) { $f += 'end before start' }
    if ($r.target -and $r.target -lt $today) { $f += 'overdue' }
    return $f
}

foreach ($p in $projects.Values) {
    $ss = @(); $tt = @()
    foreach ($k in $p.tasks) {
        if ($k.start)  { $ss += $k.start }
        if ($k.target) { $tt += $k.target }
    }
    # @(...) matters: indexing a SINGLE-element array without it returns the
    # first CHARACTER of the string, which silently produces garbage windows.
    $p.rollStart = ''; $p.rollTarget = ''
    if ($ss.Count -gt 0) { $p.rollStart  = @($ss | Sort-Object)[0] }
    if ($tt.Count -gt 0) { $p.rollTarget = @($tt | Sort-Object)[-1] }
    $p.openCount   = @($p.tasks | Where-Object { -not $_.closed }).Count
    $p.closedCount = $p.tasks.Count - $p.openCount
    $p.flags = Get-Flags $p $p.closed
    foreach ($k in $p.tasks) { $k.flags = Get-Flags $k $k.closed }
}
foreach ($o in $orphans) { $o.flags = Get-Flags $o $o.closed }

# --- filters ---------------------------------------------------------------
$projList = @($projects.Values)
if (-not $IncludeClosed) { $projList = @($projList | Where-Object { -not $_.closed }) }
if ($AgentFilter -ne '*') {
    $projList = @($projList | Where-Object {
        $_.agent -eq $AgentFilter -or @($_.tasks | Where-Object { $_.agent -eq $AgentFilter }).Count -gt 0
    })
}
Say "  $($projList.Count) project(s) after filtering" 'DarkGray'
if ($projList.Count -eq 0) {
    Clear-All
    Say '  Nothing to report with these filters.' 'Yellow'
    return
}

# effective window per project: roll-up if it has children, else its own dates
foreach ($p in $projList) {
    $p.effStart  = $(if ($p.rollStart)  { $p.rollStart }  else { $p.start })
    $p.effTarget = $(if ($p.rollTarget) { $p.rollTarget } else { $p.target })
}

# --- progress per project --------------------------------------------------
# Completion is closed tasks over all tasks. Judged against elapsed schedule
# time, so "60% done" reads differently in month one than in the final week.
# The colour never carries the meaning alone - every meter ships with a label.
$nowD = (Get-Date).Date
foreach ($p in $projList) {
    $tot  = $p.tasks.Count
    $done = $p.closedCount
    $p.pct = 0
    if ($tot -gt 0) { $p.pct = [int][Math]::Round($done / $tot * 100) }

    $elapsed = -1
    if ($p.effStart -and $p.effTarget) {
        $a = [datetime]::ParseExact($p.effStart,  'yyyy-MM-dd', $null)
        $b = [datetime]::ParseExact($p.effTarget, 'yyyy-MM-dd', $null)
        $span = ($b - $a).TotalDays
        if ($span -gt 0) {
            $elapsed = [int][Math]::Round((($nowD - $a).TotalDays / $span) * 100)
            if ($elapsed -lt 0)   { $elapsed = 0 }
            if ($elapsed -gt 100) { $elapsed = 100 }
        }
        elseif ($nowD -ge $b) { $elapsed = 100 }
        else { $elapsed = 0 }
    }
    $p.elapsed = $elapsed

    $overdue = ($p.effTarget -and $p.effTarget -lt $today -and $p.pct -lt 100)
    if ($tot -eq 0)            { $p.pstate = 'none'; $p.plabel = 'no tasks' }
    elseif ($p.pct -ge 100)    { $p.pstate = 'good'; $p.plabel = 'all tasks done' }
    elseif ($overdue)          { $p.pstate = 'bad';  $p.plabel = 'past its end date' }
    elseif ($elapsed -lt 0)    { $p.pstate = 'ok';   $p.plabel = 'no schedule to judge against' }
    elseif ($elapsed -le $p.pct + 10) { $p.pstate = 'ok';   $p.plabel = 'on track' }
    elseif ($elapsed -le $p.pct + 30) { $p.pstate = 'warn'; $p.plabel = 'slipping' }
    else                              { $p.pstate = 'bad';  $p.plabel = 'behind' }
}

$projList = @($projList | Sort-Object `
    @{ Expression = { [int]$_.closed } }, `
    @{ Expression = { if ($_.effTarget) { $_.effTarget } else { '9999-99-99' } } }, `
    @{ Expression = { $_.summary } })

# --- chart data ------------------------------------------------------------
$openTasks = @()
foreach ($p in $projList) { foreach ($k in $p.tasks) { if (-not $k.closed) { $openTasks += $k } } }

$byAgent = @{}
foreach ($p in $projList) {
    if (-not $byAgent.ContainsKey($p.agent)) { $byAgent[$p.agent] = @{ open = 0; closed = 0 } }
    foreach ($k in $p.tasks) {
        if (-not $byAgent.ContainsKey($k.agent)) { $byAgent[$k.agent] = @{ open = 0; closed = 0 } }
        if ($k.closed) { $byAgent[$k.agent].closed++ } else { $byAgent[$k.agent].open++ }
    }
}
# fixed slot order: busiest first, capped at 8, remainder folds into Other
$agentOrder = @($byAgent.GetEnumerator() | Sort-Object @{Expression={$_.Value.open}; Descending=$true}, Name |
                ForEach-Object { $_.Name })
$slotOf = @{}
$slot = 1
foreach ($a in $agentOrder) {
    if ($slot -le 8) { $slotOf[$a] = $slot; $slot++ } else { $slotOf[$a] = 9 }
}

$months = @()
$mstart = (Get-Date).AddMonths(-1)
$cursor = Get-Date -Year $mstart.Year -Month $mstart.Month -Day 1
for ($z = 0; $z -lt 18; $z++) { $months += $cursor.ToString('yyyy-MM'); $cursor = $cursor.AddMonths(1) }
$loadByMonth = @{}
foreach ($m in $months) { $loadByMonth[$m] = @{} }
$spill = 0
foreach ($k in $openTasks) {
    if (-not $k.target) { continue }
    $m = $k.target.Substring(0,7)
    if ($m -lt $months[0]) { $m = $months[0] }
    elseif ($m -gt $months[-1]) { $spill++; continue }
    if (-not $loadByMonth[$m].ContainsKey($k.agent)) { $loadByMonth[$m][$k.agent] = 0 }
    $loadByMonth[$m][$k.agent]++
}
$monthTotals = @{}
$peak = 1
foreach ($m in $months) {
    $s = 0
    foreach ($v in $loadByMonth[$m].Values) { $s += $v }
    $monthTotals[$m] = $s
    if ($s -gt $peak) { $peak = $s }
}

# --- timeline scales -------------------------------------------------------
# The overview chart is global; every PROJECT section gets its own scale, or a
# two-year portfolio would squash each project into an unreadable sliver.

function Get-Scale {
    param($Dates)
    $d = @($Dates | Where-Object { $_ })
    if ($d.Count -eq 0) { return $null }
    $srt = @($d | Sort-Object)
    $mn = [datetime]::ParseExact($srt[0],  'yyyy-MM-dd', $null)
    $mx = [datetime]::ParseExact($srt[-1], 'yyyy-MM-dd', $null)
    $a = New-Object datetime $mn.Year, $mn.Month, 1
    $b = (New-Object datetime $mx.Year, $mx.Month, 1).AddMonths(1).AddDays(-1)
    # never narrower than three months, or a one-week task fills the row
    while (($b - $a).TotalDays -lt 88) { $b = $b.AddMonths(1) }
    return @{ a = $a; b = $b; span = ($b - $a).TotalDays }
}

function Bar-In {
    # "left|width" as percentages of the given scale, or '' when undatable
    param($Scale, [string]$s, [string]$t)
    if ($null -eq $Scale) { return '' }
    if (-not $s -and -not $t) { return '' }
    $x = $(if ($s) { $s } else { $t })
    $y = $(if ($t) { $t } else { $s })
    $da = [datetime]::ParseExact($x, 'yyyy-MM-dd', $null)
    $db = [datetime]::ParseExact($y, 'yyyy-MM-dd', $null)
    if ($db -lt $da) { $tmp = $da; $da = $db; $db = $tmp }
    if ($da -lt $Scale.a) { $da = $Scale.a }
    if ($db -gt $Scale.b) { $db = $Scale.b }
    if ($db -lt $da) { return '' }
    $l = ($da - $Scale.a).TotalDays / $Scale.span * 100
    $w = ($db - $da).TotalDays / $Scale.span * 100
    if ($w -lt 0.8) { $w = 0.8 }
    if ($l + $w -gt 100) { $l = 100 - $w }
    return ("{0:F2}|{1:F2}" -f $l, $w)
}

function Today-In {
    # today's position on a scale, or -1 when today is off the end
    param($Scale)
    if ($null -eq $Scale) { return -1 }
    $t = (Get-Date).Date
    if ($t -lt $Scale.a -or $t -gt $Scale.b) { return -1 }
    return ($t - $Scale.a).TotalDays / $Scale.span * 100
}

function Axis-Row {
    param($Scale, [int]$Before, [int]$After)
    if ($null -eq $Scale) { return '' }
    # tick density has to suit the column, not just the span - 12 monthly ticks
    # in a quarter-width column collide into mush
    $step = 1
    if ($Scale.span -gt 200)  { $step = 2 }
    if ($Scale.span -gt 400)  { $step = 3 }
    if ($Scale.span -gt 800)  { $step = 6 }
    if ($Scale.span -gt 1800) { $step = 12 }
    $out = '<tr><td colspan="' + $Before + '" style="border-bottom:1px solid var(--grid);padding:0"></td><td class="ax">'
    $c = $Scale.a
    $first = $true
    while ($c -le $Scale.b) {
        $pct = ($c - $Scale.a).TotalDays / $Scale.span * 100
        $lab = $c.ToString('MMM')
        if ($c.Month -eq 1 -or $first) { $lab += " " + $c.ToString('yy') }
        $out += '<span style="left:' + ('{0:F2}' -f $pct) + '%">' + $lab + '</span>'
        $first = $false
        $c = $c.AddMonths($step)
    }
    $out += '</td>'
    if ($After -gt 0) { $out += '<td colspan="' + $After + '" style="border-bottom:1px solid var(--grid);padding:0"></td>' }
    return ($out + '</tr>')
}

# ---------------------------------------------------------------------------
#  EMIT THE HTML
# ---------------------------------------------------------------------------
Head 'WRITING THE REPORT'
Set-Overall 5 5 'writing the report'

$sb = New-Object System.Text.StringBuilder
function W { param([string]$s) [void]$sb.AppendLine($s) }

$css = @'
<style>
.viz-root,body{
 color-scheme:light;
 --surface-1:#fcfcfb; --plane:#f9f9f7;
 --text-primary:#0b0b0b; --text-secondary:#52514e; --muted:#898781;
 --grid:#e1e0d9; --axis:#c3c2b7; --ring:rgba(11,11,11,0.10);
 --s1:#2a78d6; --s2:#eb6834; --s3:#1baf7a; --s4:#eda100;
 --s5:#e87ba4; --s6:#008300; --s7:#4a3aa7; --s8:#e34948; --s9:#898781;
 --critical:#d03b3b; --good:#0ca30c;
 --m-ok:#2a78d6;   --m-ok-bg:#cde2fb;
 --m-good:#0ca30c; --m-good-bg:#cceccc;
 --m-warn:#fab219; --m-warn-bg:#fbeacb;
 --m-bad:#d03b3b;  --m-bad-bg:#f5d8d8;
 --gp:#52514e;
}
@media (prefers-color-scheme:dark){
 :root:where(:not([data-theme="light"])) .viz-root,
 :root:where(:not([data-theme="light"])) body{
  color-scheme:dark;
  --surface-1:#1a1a19; --plane:#0d0d0d;
  --text-primary:#ffffff; --text-secondary:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,0.10);
  --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500;
  --s5:#d55181; --s6:#008300; --s7:#9085e9; --s8:#e66767; --s9:#898781;
  --m-ok:#3987e5;   --m-ok-bg:#1b3c62;
  --m-good:#0ca30c; --m-good-bg:#15381a;
  --m-warn:#fab219; --m-warn-bg:#4a3a12;
  --m-bad:#d03b3b;  --m-bad-bg:#4a1f1f;
  --gp:#c3c2b7;
 }
}
:root[data-theme="dark"] .viz-root,:root[data-theme="dark"] body{
 color-scheme:dark;
 --surface-1:#1a1a19; --plane:#0d0d0d;
 --text-primary:#ffffff; --text-secondary:#c3c2b7; --muted:#898781;
 --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,0.10);
 --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500;
 --s5:#d55181; --s6:#008300; --s7:#9085e9; --s8:#e66767; --s9:#898781;
  --m-ok:#3987e5;   --m-ok-bg:#1b3c62;
  --m-good:#0ca30c; --m-good-bg:#15381a;
  --m-warn:#fab219; --m-warn-bg:#4a3a12;
  --m-bad:#d03b3b;  --m-bad-bg:#4a1f1f;
 --gp:#c3c2b7;
}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--text-primary);
 font:15px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.wrap{max-width:1240px;margin:0 auto;padding:30px 20px 90px}
h1{font-size:26px;margin:0 0 5px;letter-spacing:-.01em}
h2{font-size:19px;margin:36px 0 12px;letter-spacing:-.005em}
h3{font-size:13px;margin:0 0 10px;color:var(--muted);font-weight:600;
 text-transform:uppercase;letter-spacing:.5px}
p{margin:0 0 12px;max-width:76ch}
.lede{color:var(--text-secondary);font-size:14px;margin-bottom:22px}
.card{background:var(--surface-1);border:1px solid var(--ring);border-radius:12px;
 padding:18px 20px;margin-bottom:14px;position:relative}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(138px,1fr));gap:11px;margin-bottom:16px}
.tile{background:var(--surface-1);border:1px solid var(--ring);border-radius:11px;padding:14px 16px}
.tile b{display:block;font-size:26px;line-height:1.05;letter-spacing:-.02em;font-weight:600}
.tile span{display:block;color:var(--text-secondary);font-size:12px;margin-top:4px}
.legend{display:flex;flex-wrap:wrap;gap:14px;margin:0 0 14px;font-size:12.5px;
 color:var(--text-secondary);align-items:center}
.legend i{display:inline-block;width:11px;height:11px;border-radius:3px;margin-right:6px;vertical-align:-1px}
.plotbox{position:relative;padding-left:30px}
.plot{display:flex;align-items:flex-end;gap:6px;height:210px;border-bottom:1px solid var(--axis)}
.col{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;
 height:100%;position:relative}
.stack{width:100%;max-width:24px;display:flex;flex-direction:column-reverse}
.seg{width:100%;border-bottom:2px solid var(--surface-1)}
.stack .seg:last-child{border-radius:4px 4px 0 0}
.stack .seg:first-child{border-bottom:none}
.cap{font-size:11px;color:var(--text-secondary);margin-bottom:4px;line-height:1;
 font-variant-numeric:tabular-nums}
.xaxis{display:flex;gap:6px;margin-top:6px}
.xaxis div{flex:1;text-align:center;font-size:10px;color:var(--muted);
 font-variant-numeric:tabular-nums}
.gl{position:absolute;left:0;right:0;border-top:1px solid var(--grid)}
.gl span{position:absolute;left:-28px;top:-8px;font-size:10px;color:var(--muted);
 font-variant-numeric:tabular-nums}
.tip{position:absolute;z-index:40;background:var(--surface-1);border:1px solid var(--ring);
 border-radius:9px;padding:9px 11px;font-size:12.5px;pointer-events:none;opacity:0;
 box-shadow:0 6px 22px rgba(0,0,0,.16);min-width:150px;transition:opacity .08s}
.tip b{display:block;margin-bottom:5px}
.tip div{display:flex;justify-content:space-between;gap:14px;line-height:1.5}
.tip i{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;
 color:var(--muted);font-weight:600;padding:7px 8px;border-bottom:1px solid var(--grid)}
td{padding:6px 8px;border-bottom:1px solid var(--grid);vertical-align:top}
tr:last-child td{border-bottom:none}
td.n{text-align:right;font-variant-numeric:tabular-nums}
td.d{white-space:nowrap;font-variant-numeric:tabular-nums;font-size:12.5px;color:var(--text-secondary)}
td.mono{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:var(--muted)}
tr.closed td{color:var(--muted)}
.track{position:relative;height:10px;background:var(--grid);border-radius:5px;min-width:120px}
.track b{position:absolute;top:0;height:10px;border-radius:4px;display:block}
.track u{position:absolute;top:-3px;width:2px;height:16px;background:var(--critical);
 border-radius:1px}
.mtrack{height:14px}.mtrack b{height:14px;border-radius:5px}
td.ax{position:relative;height:16px;padding:0 8px 4px;border-bottom:1px solid var(--grid)}
td.ax span{position:absolute;bottom:3px;transform:translateX(-50%);white-space:nowrap;
 font-size:10px;color:var(--muted)}
.dot{display:inline-block;width:9px;height:9px;border-radius:3px;margin-right:7px}
.meter{display:inline-block;width:76px;height:8px;border-radius:4px;vertical-align:1px;
 overflow:hidden}
.meter b{display:block;height:8px;border-radius:4px}
.m-ok{background:var(--m-ok-bg)}   .m-ok b{background:var(--m-ok)}
.m-good{background:var(--m-good-bg)} .m-good b{background:var(--m-good)}
.m-warn{background:var(--m-warn-bg)} .m-warn b{background:var(--m-warn)}
.m-bad{background:var(--m-bad-bg)}  .m-bad b{background:var(--m-bad)}
.pct{font-size:12px;font-variant-numeric:tabular-nums;color:var(--text-secondary);
 margin-left:6px}
.pstate{font-size:11px;color:var(--muted);margin-left:5px}
.pwrap{display:inline-flex;align-items:center;white-space:nowrap}
.idx .pwrap{margin-left:9px}
.idx .meter{width:52px;height:7px}
.idx .meter b{height:7px}
.idx .pct{font-size:11.5px;margin-left:5px}
.flag{font-size:10px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;
 padding:2px 6px;border-radius:99px;border:1px solid;white-space:nowrap;margin-left:5px;
 display:inline-block}
.f-bad{color:var(--critical);border-color:var(--critical)}
.f-mut{color:var(--text-secondary);border-color:var(--axis)}
.f-ok{color:var(--good);border-color:var(--good)}
.bar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;position:sticky;top:0;z-index:30;
 background:var(--plane);padding:11px 0;margin-bottom:8px;border-bottom:1px solid var(--grid)}
.bar label{font-size:13px;display:flex;align-items:center;gap:6px;cursor:pointer;white-space:nowrap}
select,input[type=search],button{font:inherit;font-size:13px;padding:5px 9px;
 border:1px solid var(--ring);border-radius:7px;background:var(--surface-1);
 color:var(--text-primary)}
button{cursor:pointer;color:var(--text-secondary)}
button:hover{color:var(--text-primary)}
.proj{background:var(--surface-1);border:1px solid var(--ring);border-radius:12px;
 margin-bottom:13px;overflow:hidden;scroll-margin-top:64px}
.proj.isclosed{opacity:.68}
.phead{display:flex;flex-wrap:wrap;align-items:baseline;gap:9px;padding:12px 16px;cursor:pointer}
.phead:hover{background:var(--plane)}
.pname{font-weight:650;font-size:16px}
.pid{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:var(--muted)}
.pmeta{color:var(--text-secondary);font-size:12.5px}
.pwin{margin-left:auto;font-size:13px;text-align:right;white-space:nowrap;
 font-variant-numeric:tabular-nums}
.pwin small{display:block;color:var(--muted);font-size:11.5px}
.pbody{border-top:1px solid var(--grid);padding:0}
.proj.collapsed .pbody{display:none}
.caret{color:var(--muted);display:inline-block;transition:transform .12s;font-size:11px}
.proj.collapsed .caret{transform:rotate(-90deg)}
.sec{padding:12px 16px;border-bottom:1px solid var(--grid)}
.sec:last-child{border-bottom:none}
.desc{font-size:13px;color:var(--text-secondary);max-width:90ch;white-space:pre-wrap}
.idx{columns:2;column-gap:26px;font-size:13.5px}
.idx a{display:block;color:var(--text-primary);text-decoration:none;padding:3px 0;
 break-inside:avoid;border-bottom:1px solid var(--grid)}
.idx a:hover{color:var(--s1)}
.idx a > span.win{float:right;color:var(--muted);font-size:12px;
 font-variant-numeric:tabular-nums;margin-left:12px}
.hidden{display:none!important}
.gwrap{background:var(--surface-1);border:1px solid var(--ring);border-radius:12px;
 padding:0 0 10px;margin-bottom:14px;overflow:hidden}
.gbar{display:flex;gap:12px;align-items:center;padding:12px 16px 10px;font-size:13px;
 color:var(--text-secondary);flex-wrap:wrap}
.ghead{position:sticky;top:52px;z-index:20;background:var(--surface-1);
 border-bottom:1px solid var(--axis);height:20px}
.ghead .gt{position:relative;height:20px}
.ghead .gt > span{position:absolute;bottom:3px;transform:translateX(-50%);font-size:10px;
 color:var(--muted);white-space:nowrap}
.gbody{position:relative}
.grid{position:absolute;top:0;bottom:0;pointer-events:none}
.grid i{position:absolute;top:0;bottom:0;width:1px;background:var(--grid)}
.grid u{position:absolute;top:0;bottom:0;width:2px;background:var(--critical);
 opacity:.85;border-radius:1px}
.grow{display:flex;align-items:center;height:19px;font-size:12px}
.grow:hover{background:var(--plane)}
.glabel{flex:0 0 300px;padding:0 10px 0 16px;white-space:nowrap;overflow:hidden;
 text-overflow:ellipsis;color:var(--text-secondary)}
.gt{flex:1;position:relative;padding-right:16px}
.grow b{position:absolute;height:8px;border-radius:3px;top:-4px;display:block}
.gproj{height:22px}
.gproj .glabel{font-weight:650;color:var(--text-primary);padding-left:16px}
.gproj b{height:11px;top:-5px;border-radius:4px;background:var(--gp)}
.gtask .glabel{padding-left:32px;font-size:11.5px}
.gclosed b{opacity:.4}
.ghdr{padding:9px 16px 3px;font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;
 color:var(--muted);font-weight:600}
.top{position:fixed;right:18px;bottom:18px;z-index:50;border-radius:99px;padding:8px 14px;
 box-shadow:0 4px 16px rgba(0,0,0,.18)}
@media print{.bar,.top{display:none}.proj{break-inside:avoid}.proj.collapsed .pbody{display:block}}
@media (max-width:760px){.idx{columns:1}.pwin{margin-left:0;text-align:left}}
</style>
'@

$js = @'
<script>
(function(){
 var tip=document.getElementById('tip');
 if(tip){
  document.querySelectorAll('.col').forEach(function(c){
   c.addEventListener('mouseenter',function(){
    tip.innerHTML=c.dataset.tip; tip.style.opacity=1;
    var r=c.getBoundingClientRect(),w=tip.offsetParent.getBoundingClientRect();
    var x=r.left-w.left+r.width/2-tip.offsetWidth/2;
    x=Math.max(4,Math.min(x,w.width-tip.offsetWidth-4));
    var y=r.top-w.top-tip.offsetHeight-10; if(y<6){y=r.bottom-w.top+10;}
    tip.style.left=x+'px'; tip.style.top=y+'px';
   });
   c.addEventListener('mouseleave',function(){tip.style.opacity=0;});
  });
 }
 var cv=document.getElementById('chartview'),tv=document.getElementById('tableview'),
     tb=document.getElementById('tbtn');
 if(tb){tb.addEventListener('click',function(){
   var show=tv.classList.contains('hidden');
   tv.classList.toggle('hidden',!show); cv.classList.toggle('hidden',show);
   tb.textContent=show?'Show as chart':'Show as table';
 });}

 var oc=document.getElementById('fclosed'),ag=document.getElementById('fagent'),
     q=document.getElementById('fq'),pr=document.getElementById('fprob'),
     cnt=document.getElementById('fcount');
 var projs=[].slice.call(document.querySelectorAll('.proj'));
 function apply(){
  var showClosed=oc.checked, a=ag.value, term=q.value.trim().toLowerCase(), onlyProb=pr.checked;
  var shown=0;
  projs.forEach(function(p){
   var rows=[].slice.call(p.querySelectorAll('tr.task')), any=false;
   rows.forEach(function(r){
    var ok=true;
    if(!showClosed && r.dataset.closed==='1') ok=false;
    if(a && r.dataset.agent!==a) ok=false;
    if(onlyProb && r.dataset.prob!=='1') ok=false;
    if(term && r.dataset.q.indexOf(term)<0 && p.dataset.q.indexOf(term)<0) ok=false;
    r.classList.toggle('hidden',!ok); if(ok) any=true;
   });
   var ok=true;
   if(!showClosed && p.dataset.closed==='1') ok=false;
   if(a && p.dataset.agents.split('|').indexOf(a)<0) ok=false;
   if(onlyProb && p.dataset.prob!=='1') ok=false;
   if(term && p.dataset.q.indexOf(term)<0) ok=false;
   if(ok && (a||term||onlyProb) && rows.length && !any) ok=false;
   p.classList.toggle('hidden',!ok); if(ok) shown++;
  });
  cnt.textContent=shown+' of '+projs.length+' projects shown';
  document.querySelectorAll('.idx a').forEach(function(l){
   var t=document.getElementById(l.getAttribute('href').substring(1));
   l.classList.toggle('hidden', !t || t.classList.contains('hidden'));
  });
  gantt(a,term,onlyProb,showClosed);
 }

 // The gantt mirrors what the sections above decided rather than re-running the
 // filter rules, so the two views can never drift apart.
 var gt=document.getElementById('gtasks');
 function gantt(a,term,onlyProb,showClosed){
  var gc=document.getElementById('gcount'); if(!gc) return;
  var showTasks=!gt||gt.checked, np=0, nt=0;
  document.querySelectorAll('.grow[data-for]').forEach(function(r){
   var t=document.getElementById(r.dataset.for);
   var ok=!!t && !t.classList.contains('hidden');
   r.classList.toggle('hidden',!ok); if(ok) np++;
  });
  document.querySelectorAll('.grow[data-task]').forEach(function(r){
   var t=document.querySelector('tr.task[data-id="'+r.dataset.task+'"]');
   var ok=showTasks && !!t && !t.classList.contains('hidden') &&
          !t.closest('.proj').classList.contains('hidden');
   r.classList.toggle('hidden',!ok); if(ok) nt++;
  });
  document.querySelectorAll('.grow[data-orph]').forEach(function(r){
   var ok=showTasks;
   if(a && r.dataset.agent!==a) ok=false;
   if(term && r.dataset.q.indexOf(term)<0) ok=false;
   r.classList.toggle('hidden',!ok); if(ok) nt++;
  });
  var hdr=document.querySelector('.ghdr');
  if(hdr){
   var anyOrph=[].slice.call(document.querySelectorAll('.grow[data-orph]'))
                 .some(function(r){return !r.classList.contains('hidden')});
   hdr.classList.toggle('hidden',!anyOrph);
  }
  gc.textContent=np+' projects / '+nt+' tasks on the chart';
 }
 if(gt) gt.addEventListener('change',apply);
 [oc,pr].forEach(function(e){e.addEventListener('change',apply)});
 ag.addEventListener('change',apply); q.addEventListener('input',apply);
 document.querySelectorAll('.phead').forEach(function(h){
  h.addEventListener('click',function(e){
   if(e.target.tagName==='A') return;
   h.parentNode.classList.toggle('collapsed');
  });
 });
 var cb=document.getElementById('collapse');
 cb.addEventListener('click',function(){
  var doC=cb.textContent.indexOf('Collapse')===0;
  projs.forEach(function(p){p.classList.toggle('collapsed',doC)});
  cb.textContent=doC?'Expand all':'Collapse all';
 });
 document.getElementById('totop').addEventListener('click',function(){
  window.scrollTo({top:0,behavior:'smooth'});
 });
 apply();
 projs.forEach(function(p){p.classList.add('collapsed')});
 cb.textContent='Expand all';
})();
</script>
'@

$stamp = Get-Date -Format 'dd/MM/yyyy HH:mm'
$totProj   = $projList.Count
$totOpenP  = @($projList | Where-Object { -not $_.closed }).Count
$totTasks  = 0; $totOpenT = 0
foreach ($p in $projList) { $totTasks += $p.tasks.Count; $totOpenT += $p.openCount }
$probTasks = 0
foreach ($p in $projList) { foreach ($k in $p.tasks) { if ($k.flags.Count -gt 0) { $probTasks++ } } }

W '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
W '<meta name="viewport" content="width=device-width, initial-scale=1">'
W ("<title>Halo projects report - " + (Esc $stamp) + "</title>")
W $css
W '</head><body class="viz-root"><div class="wrap">'

W '<h1>Halo &mdash; Projects Report</h1>'
W ('<p class="lede">' + (Esc $Tenant) + ' &middot; generated ' + (Esc $stamp) +
   ' &middot; ' + $totProj + ' projects and ' + $totTasks + ' tasks, open and closed.' +
   $(if ($AgentFilter -ne '*') { ' Filtered to ' + (Esc $AgentFilter) + '.' } else { '' }) +
   ' Read-only snapshot &mdash; nothing was changed in Halo.</p>')

# --- tiles -----------------------------------------------------------------
W '<div class="tiles">'
foreach ($pair in @(
    @{v=$totProj;  l='Projects'},
    @{v=$totOpenP; l='Open projects'},
    @{v=($totProj-$totOpenP); l='Closed projects'},
    @{v=$totTasks; l='Tasks'},
    @{v=$totOpenT; l='Open tasks'},
    @{v=($totTasks-$totOpenT); l='Closed tasks'},
    @{v=$probTasks; l='Tasks with a date problem'},
    @{v=@($orphans | Where-Object { -not $_.closed }).Count; l='Open tasks with no project'} )) {
    W ('<div class="tile"><b>' + $pair.v + '</b><span>' + $pair.l + '</span></div>')
}
W '</div>'

# --- delivery load chart ---------------------------------------------------
W '<h2>When the open work is due</h2>'
W '<div class="card">'
W '<div style="display:flex;align-items:baseline;gap:12px"><h3 style="margin:0">Open tasks by deadline month and owner</h3>'
W '<button id="tbtn" style="margin-left:auto">Show as table</button></div>'
W '<div class="legend">'
foreach ($a in $agentOrder) {
    if ($slotOf[$a] -ge 9) { continue }
    W ('<span><i style="background:var(--s' + $slotOf[$a] + ')"></i>' + (Esc $a) + '</span>')
}
if (@($agentOrder | Where-Object { $slotOf[$_] -ge 9 }).Count -gt 0) {
    W '<span><i style="background:var(--s9)"></i>Other</span>'
}
W '</div>'

W '<div id="chartview"><div class="plotbox"><div style="position:relative;height:210px">'
$stepv = 5; if ($peak -gt 30) { $stepv = 10 }; if ($peak -gt 80) { $stepv = 25 }
$g = $stepv
while ($g -le $peak) {
    W ('<div class="gl" style="bottom:' + ('{0:F2}' -f ($g / $peak * 100)) + '%"><span>' + $g + '</span></div>')
    $g += $stepv
}
W '<div class="plot">'
foreach ($m in $months) {
    $tot = $monthTotals[$m]
    $tipHtml = '<b>' + ([datetime]::ParseExact($m + '-01','yyyy-MM-dd',$null).ToString('MMMM yyyy')) + '</b>'
    W ('<div class="col" data-tip="' + '{{TIP' + $m + '}}' + '">')
    if ($tot -gt 0) {
        W ('<div class="cap">' + $tot + '</div>')
        W ('<div class="stack" style="height:calc(' + ('{0:F2}' -f ($tot / $peak * 100)) + '% - 17px)">')
        foreach ($a in $agentOrder) {
            $v = 0
            if ($loadByMonth[$m].ContainsKey($a)) { $v = $loadByMonth[$m][$a] }
            if ($v -le 0) { continue }
            W ('<div class="seg" style="flex:' + $v + ';background:var(--s' + $slotOf[$a] + ')"></div>')
            $tipHtml += '<div><span><i style="background:var(--s' + $slotOf[$a] + ')"></i>' +
                        (Esc $a) + '</span><span>' + $v + '</span></div>'
        }
        W '</div>'
    }
    $tipHtml += '<div style="margin-top:5px;border-top:1px solid var(--grid);padding-top:5px">' +
                '<span>Total</span><span><b>' + $tot + '</b></span></div>'
    $script:sb.Replace('{{TIP' + $m + '}}', [System.Net.WebUtility]::HtmlEncode($tipHtml)) | Out-Null
    W '</div>'
}
W '</div>'
W '<div class="xaxis">'
foreach ($m in $months) {
    $lab = [datetime]::ParseExact($m + '-01','yyyy-MM-dd',$null).ToString('MMM')
    if ($m.Substring(5,2) -eq '01' -or $m -eq $months[0]) { $lab += ' ' + $m.Substring(2,2) }
    W ('<div>' + $lab + '</div>')
}
W '</div></div></div>'

# table view - the relief for light-mode contrast, and it prints
W '<div id="tableview" class="hidden" style="margin-top:8px"><table><thead><tr><th>Month</th>'
foreach ($a in $agentOrder) { W ('<th class="n">' + (Esc ($a -split ' ')[0]) + '</th>') }
W '<th class="n">Total</th></tr></thead><tbody>'
foreach ($m in $months) {
    W ('<tr><td class="d">' + ([datetime]::ParseExact($m + '-01','yyyy-MM-dd',$null).ToString('MMMM yyyy')) + '</td>')
    foreach ($a in $agentOrder) {
        $v = ''
        if ($loadByMonth[$m].ContainsKey($a)) { $v = $loadByMonth[$m][$a] }
        W ('<td class="n">' + $v + '</td>')
    }
    W ('<td class="n"><b>' + $(if ($monthTotals[$m] -gt 0) { $monthTotals[$m] } else { '' }) + '</b></td></tr>')
}
W '</tbody></table></div>'
W '<div class="tip" id="tip"></div>'
if ($spill -gt 0) {
    W ('<p class="lede" style="margin:10px 0 0">' + $spill +
       ' open task(s) fall beyond this 18-month window and are not on the chart. They are still in the project sections below.</p>')
}
W '</div>'

# --- workload table --------------------------------------------------------
W '<h2>Workload by owner</h2><div class="card"><table><thead><tr><th>Owner</th>'
W '<th class="n" style="width:90px">Open</th><th class="n" style="width:90px">Closed</th>'
W '<th class="n" style="width:90px">Total</th></tr></thead><tbody>'
foreach ($a in $agentOrder) {
    if (($byAgent[$a].open + $byAgent[$a].closed) -eq 0) { continue }
    W ('<tr><td><span class="dot" style="background:var(--s' + $slotOf[$a] + ')"></span>' +
       (Esc $a) + '</td><td class="n">' + $byAgent[$a].open + '</td><td class="n">' +
       $byAgent[$a].closed + '</td><td class="n">' + ($byAgent[$a].open + $byAgent[$a].closed) + '</td></tr>')
}
W '</tbody></table></div>'

# --- everything with a date problem ---------------------------------------
W '<h2>Open records with a date problem</h2><div class="card">'
$anyProb = $false
W '<table><thead><tr><th style="width:180px">Problem</th><th>Record</th>'
W '<th style="width:130px">Owner</th><th style="width:24%">Project</th></tr></thead><tbody>'
foreach ($p in $projList) {
    if ($p.flags.Count -gt 0) {
        $anyProb = $true
        W ('<tr><td><span class="flag f-bad">' + (Esc ($p.flags -join ', ')) + '</span></td>' +
           '<td><b>' + (Esc $p.summary) + '</b> <span class="pid">#' + $p.id + '</span></td>' +
           '<td class="d">' + (Esc $p.agent) + '</td><td class="d">the project itself</td></tr>')
    }
    foreach ($k in $p.tasks) {
        if ($k.flags.Count -eq 0) { continue }
        $anyProb = $true
        W ('<tr><td><span class="flag f-bad">' + (Esc ($k.flags -join ', ')) + '</span></td>' +
           '<td>' + (Esc $k.summary) + ' <span class="pid">#' + $k.id + '</span></td>' +
           '<td class="d">' + (Esc $k.agent) + '</td><td class="d">' + (Esc $p.summary) + '</td></tr>')
    }
}
if (-not $anyProb) {
    W '<tr><td colspan="4"><span class="flag f-ok">all clear</span> Every open record has a sensible start and end date.</td></tr>'
}
W '</tbody></table></div>'

# --- orphans ---------------------------------------------------------------
# Open only: a closed task with no project is history, not something anyone
# needs to act on. The count of what is being left out is still stated, so
# nothing disappears silently.
$openOrph = @($orphans | Where-Object { -not $_.closed })
$closedOrph = $orphans.Count - $openOrph.Count
if ($openOrph.Count -gt 0) {
    W ('<h2>Tasks that belong to no project</h2><div class="card">')
    W ('<p class="lede" style="margin:0 0 12px">' + $openOrph.Count + ' open project task(s) ' +
       'are not attached to any project in this report, so they appear on nobody&rsquo;s project plan.' +
       $(if ($closedOrph -gt 0) { ' A further ' + $closedOrph + ' closed one(s) are not listed.' } else { '' }) +
       '</p>')
    W '<table><thead><tr><th style="width:62px">ID</th><th>Task</th><th style="width:130px">Owner</th>'
    W '<th style="width:96px">Starts</th><th style="width:96px">Ends</th>'
    W '<th style="width:110px">Status</th></tr></thead><tbody>'
    foreach ($o in @($openOrph | Sort-Object @{Expression={if($_.target){$_.target}else{'9999-99-99'}}}, @{Expression={$_.summary}})) {
        W ('<tr><td class="mono">' + $o.id + '</td><td>' + (Esc $o.summary) +
           $(if ($o.flags.Count -gt 0) { '<span class="flag f-bad">' + (Esc ($o.flags -join ', ')) + '</span>' } else { '' }) +
           '</td><td class="d">' + (Esc $o.agent) +
           '</td><td class="d">' + $(if ($o.start)  { Show-D $o.start }  else { '&mdash;' }) +
           '</td><td class="d">' + $(if ($o.target) { Show-D $o.target } else { '&mdash;' }) +
           '</td><td class="d">' + (Esc $o.status) + '</td></tr>')
    }
    W '</tbody></table></div>'
}
elseif ($closedOrph -gt 0) {
    W ('<h2>Tasks that belong to no project</h2><div class="card">')
    W ('<p class="lede" style="margin:0"><span class="flag f-ok">all clear</span> ' +
       'Every open task sits under a project. ' + $closedOrph +
       ' closed task(s) have no project, which does not matter.</p></div>')
}

function Meter-Html {
    param($p, [switch]$Compact)
    if ($p.closed) { return '' }
    if ($p.pstate -eq 'none') {
        if ($Compact) { return '' }
        return '<span class="pstate">no tasks</span>'
    }
    $t = 'meter m-' + $p.pstate
    $h = '<span class="pwrap" title="' + (Esc ($p.pct.ToString() + '% of tasks closed' +
         $(if ($p.elapsed -ge 0) { ' - ' + $p.elapsed + '% of the schedule elapsed' } else { '' }) +
         ' - ' + $p.plabel)) + '">'
    $h += '<span class="' + $t + '"><b style="width:' + $p.pct + '%"></b></span>'
    $h += '<span class="pct">' + $p.pct + '%</span>'
    if (-not $Compact) { $h += '<span class="pstate">' + (Esc $p.plabel) + '</span>' }
    return ($h + '</span>')
}

# --- filter bar + index ----------------------------------------------------
W '<h2>Projects</h2>'
W '<div class="bar">'
W '<label><input type="checkbox" id="fclosed"> Closed records</label>'
W '<label><input type="checkbox" id="fprob"> Only date problems</label>'
W '<label>Owner <select id="fagent"><option value="">everyone</option>'
foreach ($a in $agentOrder) { W ('<option>' + (Esc $a) + '</option>') }
W '</select></label>'
W '<input type="search" id="fq" placeholder="Search project or task&hellip;" style="min-width:210px">'
W '<button id="collapse">Collapse all</button>'
W '<span id="fcount" class="pid" style="margin-left:auto"></span>'
W '</div>'

W '<div class="card"><h3>Contents</h3><div class="idx">'
foreach ($p in $projList) {
    $w = ''
    if ($p.effStart -or $p.effTarget) { $w = (Show-D $p.effStart) + ' - ' + (Show-D $p.effTarget) }
    W ('<a href="#p' + $p.id + '"><span class="win">' + $w + '</span>' +
       (Esc $p.summary) + (Meter-Html $p -Compact) + '</a>')
}
W '</div></div>'

# --- one section per project ----------------------------------------------
$pi = 0
foreach ($p in $projList) {
    $pi++
    Set-Step 'Writing project sections' $pi $projList.Count "$pi of $($projList.Count) - $($p.summary)"

    $agents = @{}
    $agents[$p.agent] = $true
    foreach ($k in $p.tasks) { $agents[$k.agent] = $true }
    $agentsAttr = (@($agents.Keys | Sort-Object) -join '|')

    $hay = ($p.summary + ' ' + $p.agent + ' ' + $p.client + ' ' + $p.site)
    foreach ($k in $p.tasks) { $hay += ' ' + $k.summary + ' ' + $k.agent }
    $hasProb = $p.flags.Count -gt 0
    foreach ($k in $p.tasks) { if ($k.flags.Count -gt 0) { $hasProb = $true } }

    # this project's own timeline scale
    $sd = @($p.start, $p.target, $p.effStart, $p.effTarget)
    foreach ($k in $p.tasks)      { $sd += $k.start;  $sd += $k.target }
    foreach ($m in $p.milestones) { $sd += $m.start;  $sd += $m.target }
    $scale = Get-Scale $sd
    $tp = Today-In $scale
    $todayMark = ''
    if ($tp -ge 0) { $todayMark = '<u style="left:' + ('{0:F2}' -f $tp) + '%"></u>' }

    $cls = 'proj'
    if ($p.closed) { $cls += ' isclosed' }
    W ('<div class="' + $cls + '" id="p' + $p.id + '"' +
       ' data-closed="' + $(if ($p.closed) { 1 } else { 0 }) + '"' +
       ' data-prob="' + $(if ($hasProb) { 1 } else { 0 }) + '"' +
       ' data-agents="' + (Esc $agentsAttr) + '"' +
       ' data-q="' + (Esc $hay.ToLower()) + '">')

    # header
    W '<div class="phead"><span class="caret">&#9662;</span>'
    W ('<span class="pid">#' + $p.id + '</span><span class="pname">' + (Esc $p.summary) + '</span>')
    W ('<span class="flag ' + $(if ($p.closed) { 'f-mut">Closed' } else { 'f-ok">Open' }) + '</span>')
    if ($p.status) { W ('<span class="pmeta">' + (Esc $p.status) + '</span>') }
    $meta = @()
    if ($p.agent)  { $meta += (Esc $p.agent) }
    if ($p.site)   { $meta += (Esc $p.site) }
    $meta += ($p.tasks.Count.ToString() + ' task' + $(if ($p.tasks.Count -eq 1) { '' } else { 's' }) +
              ' (' + $p.openCount + ' open / ' + $p.closedCount + ' closed)')
    if ($p.milestones -and $p.milestones.Count -gt 0) { $meta += ($p.milestones.Count.ToString() + ' milestones') }
    W ('<span class="pmeta">' + ($meta -join ' &middot; ') + '</span>')
    W (Meter-Html $p)
    if ($p.flags.Count -gt 0) { W ('<span class="flag f-bad">' + (Esc ($p.flags -join ', ')) + '</span>') }

    $eff = '&mdash;'
    if ($p.effStart -or $p.effTarget) { $eff = (Show-D $p.effStart) + ' &rarr; ' + (Show-D $p.effTarget) }
    W ('<span class="pwin">' + $eff)
    $own = ''
    if ($p.start -or $p.target) { $own = (Show-D $p.start) + ' &rarr; ' + (Show-D $p.target) }
    if ($own -ne ((Show-D $p.effStart) + ' &rarr; ' + (Show-D $p.effTarget))) {
        W ('<small>project record says ' + $(if ($own) { $own } else { 'no dates' }) + '</small>')
    }
    W '</span></div>'

    W '<div class="pbody">'

    if ($p.details) {
        $dtl = $p.details
        if ($dtl.Length -gt 900) { $dtl = $dtl.Substring(0,900) + '...' }
        W ('<div class="sec"><h3>Description</h3><div class="desc">' + (Esc $dtl) + '</div></div>')
    }

    # milestones
    if ($p.milestones -and $p.milestones.Count -gt 0) {
        W '<div class="sec"><h3>Milestones</h3><table>'
        W (Axis-Row $scale 2 0)
        foreach ($m in @($p.milestones | Sort-Object @{Expression={$_.seq}}, @{Expression={$_.start}})) {
            $bp = Bar-In $scale $m.start $m.target
            $kids = @($p.tasks | Where-Object { $_.milestone -eq $m.id }).Count
            W ('<tr><td style="width:170px"><b>' + (Esc $m.name) + '</b>' +
               '<span class="pmeta" style="margin-left:7px">' + $kids + ' task' + $(if ($kids -eq 1) { '' } else { 's' }) + '</span></td>')
            W ('<td class="d" style="width:160px">' + (Show-D $m.start) + ' &rarr; ' + (Show-D $m.target) + '</td>')
            if ($bp) {
                $parts = $bp -split '\|'
                W ('<td><div class="track mtrack"><b style="left:' + $parts[0] + '%;width:' + $parts[1] +
                   '%;background:var(--s1)"></b>' + $todayMark + '</div></td>')
            } else { W '<td><span class="flag f-mut">no dates</span></td>' }
            W '</tr>'
        }
        W '</table></div>'
    }

    # tasks
    W '<div class="sec"><h3>Tasks</h3>'
    if ($p.tasks.Count -eq 0) {
        W '<p class="lede" style="margin:0">This project has no tasks.</p>'
    } else {
        W '<table><thead><tr><th style="width:58px">ID</th><th>Task</th>'
        W '<th style="width:118px">Owner</th><th style="width:112px">Status</th>'
        W '<th style="width:92px">Start</th><th style="width:92px">End</th>'
        W '<th style="width:26%">Timeline</th></tr></thead><tbody>'
        W (Axis-Row $scale 6 0)
        $msName = @{}
        foreach ($m in $p.milestones) { $msName[$m.id] = $m.name }
        $ordered = @($p.tasks | Sort-Object `
            @{Expression={[int]$_.closed}}, `
            @{Expression={if ($_.milestone -gt 0) { $_.milestone } else { 999999 }}}, `
            @{Expression={if ($_.start) { $_.start } elseif ($_.target) { $_.target } else { '9999-99-99' }}})
        foreach ($k in $ordered) {
            $bp = Bar-In $scale $k.start $k.target
            $rc = 'task'
            if ($k.closed) { $rc += ' closed' }
            W ('<tr class="' + $rc + '" data-id="' + $k.id +
               '" data-closed="' + $(if ($k.closed) { 1 } else { 0 }) +
               '" data-prob="' + $(if ($k.flags.Count -gt 0) { 1 } else { 0 }) +
               '" data-agent="' + (Esc $k.agent) +
               '" data-q="' + (Esc ($k.summary + ' ' + $k.agent).ToLower()) + '">')
            W ('<td class="mono">' + $k.id + '</td>')
            W ('<td>' + (Esc $k.summary))
            if ($k.milestone -gt 0 -and $msName.ContainsKey($k.milestone)) {
                W ('<span class="flag f-mut">' + (Esc $msName[$k.milestone]) + '</span>')
            }
            if ($k.flags.Count -gt 0) { W ('<span class="flag f-bad">' + (Esc ($k.flags -join ', ')) + '</span>') }
            W '</td>'
            W ('<td class="d">' + (Esc $k.agent) + '</td>')
            W ('<td class="d">' + (Esc $k.status) + '</td>')
            W ('<td class="d">' + $(if ($k.start) { Show-D $k.start } else { '&mdash;' }) + '</td>')
            $endTxt = '&mdash;'
            if ($k.target) { $endTxt = Show-D $k.target }
            elseif ($k.closedOn) { $endTxt = 'closed ' + (Show-D $k.closedOn) }
            W ('<td class="d">' + $endTxt + '</td>')
            if ($bp) {
                $parts = $bp -split '\|'
                $col = 's' + $slotOf[$k.agent]
                W ('<td><div class="track"><b style="left:' + $parts[0] + '%;width:' + $parts[1] +
                   '%;background:var(--' + $col + ')' + $(if ($k.closed) { ';opacity:.45' } else { '' }) +
                   '"></b>' + $todayMark + '</div></td>')
            } else { W '<td></td>' }
            W '</tr>'
        }
        W '</tbody></table>'
    }
    W '</div>'   # sec
    W '</div>'   # pbody
    W '</div>'   # proj
}
Clear-Step

# ---------------------------------------------------------------------------
#  THE WHOLE PORTFOLIO ON ONE TIMELINE
# ---------------------------------------------------------------------------
# One shared scale, fixed over every record open and closed, so the chart never
# rescales when the filters change - a bar means the same thing all the way
# down. Row visibility mirrors the sections above rather than re-implementing
# the filter rules, so the two can never disagree.

Set-Step 'Writing the gantt' 1 1

$gDates = @()
foreach ($p in $projList) {
    foreach ($d in @($p.effStart, $p.effTarget)) { if ($d) { $gDates += $d } }
    foreach ($k in $p.tasks) { foreach ($d in @($k.start, $k.target)) { if ($d) { $gDates += $d } } }
}
foreach ($o in $openOrph) { foreach ($d in @($o.start, $o.target)) { if ($d) { $gDates += $d } } }
$G = Get-Scale $gDates

if ($null -ne $G) {
    $gToday = Today-In $G

    W '<h2>Everything on one timeline</h2>'
    W ('<p class="lede" style="margin:-4px 0 12px">Every project and task on a single scale, ' +
       'in the same order as the sections above. The filters at the top of the Projects section ' +
       'drive this chart too.</p>')
    W '<div class="gwrap">'
    W '<div class="gbar">'
    W '<label><input type="checkbox" id="gtasks" checked> Show tasks</label>'
    W '<span style="color:var(--muted)">'
    W '<span style="display:inline-block;width:22px;height:8px;border-radius:3px;background:var(--gp);vertical-align:0"></span> project'
    W ' &nbsp; task bars carry their owner&rsquo;s colour &nbsp; '
    W '<span style="display:inline-block;width:2px;height:12px;background:var(--critical);vertical-align:-2px"></span> today'
    W '</span>'
    W '<span id="gcount" class="pid" style="margin-left:auto"></span>'
    W '</div>'

    # month scale, aligned to the track column
    $gstep = 1
    if ($G.span -gt 400)  { $gstep = 2 }
    if ($G.span -gt 900)  { $gstep = 3 }
    if ($G.span -gt 2000) { $gstep = 6 }
    W '<div class="ghead"><div class="grow" style="height:20px"><span class="glabel"></span><span class="gt">'
    $c3 = $G.a; $firstTick = $true
    while ($c3 -le $G.b) {
        $pct = ($c3 - $G.a).TotalDays / $G.span * 100
        $lab = $c3.ToString('MMM')
        if ($c3.Month -eq 1 -or $firstTick) { $lab += ' ' + $c3.ToString('yy') }
        W ('<span style="left:' + ('{0:F2}' -f $pct) + '%">' + $lab + '</span>')
        $firstTick = $false
        $c3 = $c3.AddMonths($gstep)
    }
    W '</span></div></div>'

    # gridlines and the today line, drawn once behind every row
    W '<div class="gbody"><div class="grid" style="left:300px;right:16px">'
    $c4 = $G.a
    while ($c4 -le $G.b) {
        $pct = ($c4 - $G.a).TotalDays / $G.span * 100
        W ('<i style="left:' + ('{0:F2}' -f $pct) + '%"></i>')
        $c4 = $c4.AddMonths($gstep)
    }
    if ($gToday -ge 0) { W ('<u style="left:' + ('{0:F2}' -f $gToday) + '%"></u>') }
    W '</div>'

    foreach ($p in $projList) {
        $bp = Bar-In $G $p.effStart $p.effTarget
        $cl = 'grow gproj'
        if ($p.closed) { $cl += ' gclosed' }
        $ttl = $p.summary + '  ' + (Show-D $p.effStart) + ' - ' + (Show-D $p.effTarget) +
               '  (' + $p.openCount + ' open / ' + $p.closedCount + ' closed)'
        W ('<div class="' + $cl + '" data-for="p' + $p.id + '">' +
           '<span class="glabel" title="' + (Esc $ttl) + '">' + (Esc $p.summary) + '</span><span class="gt">')
        if ($bp) {
            $q = $bp -split '\|'
            W ('<b style="left:' + $q[0] + '%;width:' + $q[1] + '%"></b>')
        }
        W '</span></div>'

        $ordG = @($p.tasks | Sort-Object `
            @{Expression={[int]$_.closed}}, `
            @{Expression={if ($_.milestone -gt 0) { $_.milestone } else { 999999 }}}, `
            @{Expression={if ($_.start) { $_.start } elseif ($_.target) { $_.target } else { '9999-99-99' }}})
        foreach ($k in $ordG) {
            $bk = Bar-In $G $k.start $k.target
            $ck = 'grow gtask'
            if ($k.closed) { $ck += ' gclosed' }
            $tk = $k.summary + '  ' + $k.agent + '  ' +
                  $(if ($k.start) { Show-D $k.start } else { 'no start' }) + ' - ' +
                  $(if ($k.target) { Show-D $k.target } else { 'no end' })
            W ('<div class="' + $ck + '" data-task="' + $k.id + '">' +
               '<span class="glabel" title="' + (Esc $tk) + '">' + (Esc $k.summary) + '</span><span class="gt">')
            if ($bk) {
                $q = $bk -split '\|'
                W ('<b style="left:' + $q[0] + '%;width:' + $q[1] +
                   '%;background:var(--s' + $slotOf[$k.agent] + ')"></b>')
            }
            W '</span></div>'
        }
    }

    if ($openOrph.Count -gt 0) {
        W '<div class="ghdr">No project</div>'
        foreach ($o in @($openOrph | Sort-Object @{Expression={if($_.target){$_.target}else{'9999-99-99'}}}, @{Expression={$_.summary}})) {
            $bo = Bar-In $G $o.start $o.target
            $so = 9
            if ($slotOf.ContainsKey($o.agent)) { $so = $slotOf[$o.agent] }
            $to = $o.summary + '  ' + $o.agent + '  ' +
                  $(if ($o.start) { Show-D $o.start } else { 'no start' }) + ' - ' +
                  $(if ($o.target) { Show-D $o.target } else { 'no end' })
            W ('<div class="grow gtask gorph" data-orph="1" data-agent="' + (Esc $o.agent) +
               '" data-q="' + (Esc ($o.summary + ' ' + $o.agent).ToLower()) + '">' +
               '<span class="glabel" title="' + (Esc $to) + '">' + (Esc $o.summary) + '</span><span class="gt">')
            if ($bo) {
                $q = $bo -split '\|'
                W ('<b style="left:' + $q[0] + '%;width:' + $q[1] + '%;background:var(--s' + $so + ')"></b>')
            }
            W '</span></div>'
        }
    }
    W '</div></div>'
}
Clear-Step

W '<button class="top" id="totop">Back to top</button>'
W $js
W '</div></body></html>'

# --- write -----------------------------------------------------------------
$path = $OutFile
if ([string]::IsNullOrWhiteSpace($path)) {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($docs)) { $docs = $env:TEMP }
    $path = Join-Path $docs ('Halo-Projects-' + (Get-Date -Format 'yyyyMMdd-HHmm') + '.html')
}
try {
    [System.IO.File]::WriteAllText($path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    Clear-All
    Write-Host ''
    Say "  Report written: $path" 'Green'
    Say ("  {0} projects, {1} tasks, {2:N0} KB" -f $totProj, $totTasks, ((Get-Item $path).Length / 1KB)) 'DarkGray'
    if ($OpenWhenDone) { Start-Process $path }
}
catch {
    Clear-All
    Say "  Could not write $path : $($_.Exception.Message)" 'Red'
    Say '  Set $OutFile to somewhere you can write and run again.' 'Yellow'
}
Write-Host ''
