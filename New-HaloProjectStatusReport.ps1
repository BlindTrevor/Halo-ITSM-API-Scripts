<#
    New-HaloProjectStatusReport.ps1
    ===============================
    READ-ONLY. Asks for ONE project - by ticket id, or by part of its name -
    and writes a self-contained HTML status report for that project alone:

        * a RAG verdict with the reasoning spelled out, never a bare colour
        * two progress bars - work completed (weighted by how long each task
          runs, not a flat headcount) against schedule elapsed
        * who is assigned, and what each of them is carrying
        * milestones, with their own completion
        * every task, open and closed, with id, owner, dates, close date
          and a timeline bar
        * what is overdue, what lands in the next fortnight, and every
          record with a broken date

    Open in PowerShell ISE, paste credentials, press F5. Writes nothing to
    Halo and nothing to disk except the report.

    Why this is quick where the portfolio report is slow: giving a ticket id
    goes straight to that record, so there is no sweep of the whole history.
    Only a name search has to page the ticket list.

    Notes about this tenant, learned the hard way:
      * count=N is capped at 1000 and returns the NEWEST 1000 while still
        reporting record_count=1000, so a naive completeness check passes
        while most of the history is missing. Page, never count.
      * The ticket LIST omits startdate, so every task needs an individual
        GET. That is the slow part, and here it is bounded by the size of
        one project rather than the whole tenant.
      * Project = tickettype_id 57 (its display name has a trailing space),
        Project Task = 58. Match on id, never on name.
      * Milestones ride on the PROJECT ticket as a milestones[] array with
        underscored start_date / target_date. A task points at one through
        milestone_id.
      * 1900-01-01 is Halo's "unset" date.
      * Deactivated agents are absent from /api/Agent, so their tickets come
        back as "id:NN" unless they are named in $FormerAgents below.
#>

# ===========================================================================
#  SETTINGS
# ===========================================================================

$ClientId     = ''
$ClientSecret = ''

$ProjectId     = ''        # blank = ask. A number goes straight there; text searches names.
$OutFile       = ''        # blank = Documents\Halo-Project-<id>-<name>-<timestamp>.html
$OpenWhenDone  = $true     # launch the report in the default browser

# Agents who have left are deactivated in Halo and vanish from /api/Agent, so
# their tickets come back as "id:NN". Name them here and they read properly.
# NOTE: bare numeric keys in a hashtable literal are INTEGERS. Look them up
# with an int, never with "$id", or the lookup silently never matches.
$FormerAgents = @{
    13 = 'Joe Bloggs'
}
$FormerSuffix  = ' (left)'   # set to '' to show the name with no marker

# How the RAG verdict is decided. These are the only knobs, and the report
# always prints its reasoning beside the colour, so nobody has to guess.
$AmberBehindBy = 10        # work this many points behind elapsed schedule = amber
$RedBehindBy   = 30        # this far behind = red
$DueSoonDays   = 14        # what counts as "landing shortly"

# ===========================================================================

$ErrorActionPreference = 'Stop'
$Tenant  = 'contoso'
$AuthUrl = "https://$Tenant.haloitsm.com/auth/token"
$ApiBase = "https://$Tenant.haloitsm.com/api"
$Scope   = 'all'           # /api/Agent and /api/Status may 403 on a narrower scope
$TYPE_PROJECT = 57
$TYPE_TASK    = 58
$PAGE         = 100

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

function To-Date {
    param([string]$iso)
    if ([string]::IsNullOrWhiteSpace($iso)) { return $null }
    return [datetime]::ParseExact($iso, 'yyyy-MM-dd', $null)
}

function Esc {
    param($t)
    if ($null -eq $t) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$t)
}

function Strip-Html {
    # Halo details often arrive as HTML. Flatten it, or the description block
    # is a wall of markup instead of something a reader can use.
    param($t)
    if ($null -eq $t) { return '' }
    $s = [string]$t
    if ($s -notmatch '<') { return $s }
    $s = $s -replace '(?is)<(script|style).*?</\1>', ''
    $s = $s -replace '(?i)<br\s*/?>', "`n"
    $s = $s -replace '(?i)</(p|div|tr|li|h[1-6])>', "`n"
    $s = $s -replace '<[^>]+>', ''
    $s = [System.Net.WebUtility]::HtmlDecode($s)
    $s = $s -replace "`r`n", "`n"
    $s = $s -replace "\n{3,}", "`n`n"
    return $s.Trim()
}

function Get-WebErrorBody {
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return $ErrorRecord.ErrorDetails.Message
    }
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($null -eq $resp) { return '' }
        $st = $resp.GetResponseStream()
        try { $st.Position = 0 } catch { }
        return (New-Object System.IO.StreamReader($st)).ReadToEnd()
    } catch { return '' }
}

function Say  { param([string]$m,[string]$c='Gray') Write-Host $m -ForegroundColor $c }
function Head {
    param([string]$t)
    Write-Host ''
    Write-Host ("  " + $t) -ForegroundColor Cyan
    Write-Host ("  " + ('-' * 100)) -ForegroundColor DarkGray
}
function Set-Step {
    param([string]$Activity,[int]$I,[int]$N,[string]$Status='')
    if ($N -le 0) { return }
    $s = $Status
    if ([string]::IsNullOrWhiteSpace($s)) { $s = "$I of $N" }
    Write-Progress -Id 1 -Activity $Activity -Status $s -PercentComplete ([int]([Math]::Min(100, $I / $N * 100)))
}
function Clear-Step { Write-Progress -Id 1 -Activity ' ' -Completed }

function Get-Ticket {
    param([int]$Id)
    try { return Invoke-RestMethod -Uri "$ApiBase/Tickets/$Id" -Headers $script:Headers -Method Get -TimeoutSec 120 }
    catch { return $null }
}

# ---------------------------------------------------------------------------
#  RUN
# ---------------------------------------------------------------------------

Write-Host ''
Write-Host '  HALO PROJECT STATUS REPORT - read-only, one project' -ForegroundColor Cyan
Write-Host "  tenant $Tenant"

$cid = Resolve-Credential -Inline $ClientId     -EnvName 'HALO_CLIENT_ID'     -Prompt 'Halo Client ID'
$sec = Resolve-Credential -Inline $ClientSecret -EnvName 'HALO_CLIENT_SECRET' -Prompt 'Halo Client Secret'

Say '  Requesting token...' 'DarkGray'
try {
    $token = (Invoke-RestMethod -Uri $AuthUrl -Method Post -Body @{
        grant_type='client_credentials'; client_id=$cid; client_secret=$sec
        scope=$Scope; tenant=$Tenant } -ContentType 'application/x-www-form-urlencoded').access_token
}
catch {
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

# ---------------------------------------------------------------------------
#  PICK THE PROJECT
# ---------------------------------------------------------------------------

function Get-ProjectList {
    # Only a name search needs this. Page the whole list - count=N lies on this
    # tenant - and keep just the project records.
    $out = New-Object System.Collections.ArrayList
    $pg = 1; $total = $null; $fetched = 0
    while ($true) {
        try {
            $r = Invoke-RestMethod -Uri "$ApiBase/Tickets?pageinate=true&page_size=$PAGE&page_no=$pg" `
                 -Headers $script:Headers -Method Get -TimeoutSec 600
        } catch { break }
        $b = Get-Collection $r 'tickets'
        if ($null -eq $total) {
            $rc = Get-Val $r 'record_count'
            if ($null -ne $rc) { $total = [int]$rc }
        }
        if ($b.Count -eq 0) { break }
        $fetched += $b.Count
        foreach ($t in $b) {
            $tt = Get-Val $t 'tickettype_id'
            if ($null -ne $tt -and [int]$tt -eq $TYPE_PROJECT) { [void]$out.Add($t) }
        }
        $pct = 1
        if ($total -and $total -gt 0) { $pct = [int]([Math]::Min(100, $fetched / $total * 100)) }
        Write-Progress -Id 1 -Activity 'Searching projects' -Status "page $pg - $fetched of $total" -PercentComplete $pct
        if ($b.Count -lt $PAGE) { break }
        if ($null -ne $total -and $fetched -ge $total) { break }
        $pg++
        if ($pg -gt 800) { break }
    }
    Clear-Step
    return @($out)
}

function Resolve-Project {
    param([string]$Answer)
    $ans = $Answer
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($ans)) {
            Write-Host ''
            $ans = Read-Host '  Project ticket id (or part of the project name)'
            if ([string]::IsNullOrWhiteSpace($ans)) { return $null }
        }
        $ans = $ans.Trim()

        # a bare number goes straight to the record - no sweep, no waiting
        $n = 0
        if ([int]::TryParse($ans, [ref]$n) -and $n -gt 0) {
            Say "  reading ticket $n ..." 'DarkGray'
            $t = Get-Ticket -Id $n
            if ($null -eq $t) {
                Say "  No ticket $n, or it is not visible to this API user." 'Yellow'
                $ans = ''; continue
            }
            $tt = Get-Val $t 'tickettype_id'
            if ($null -eq $tt -or [int]$tt -ne $TYPE_PROJECT) {
                # a task id is the easiest thing in the world to paste by
                # mistake, so follow it up to its project rather than refusing
                if ($null -ne $tt -and [int]$tt -eq $TYPE_TASK) {
                    $par = Get-Val $t 'parent_id'
                    if ($null -ne $par -and [int]$par -gt 0) {
                        Say "  $n is a project TASK, not a project. Its project is $([int]$par) - using that." 'Yellow'
                        $p = Get-Ticket -Id ([int]$par)
                        if ($null -ne $p) { return $p }
                    }
                }
                $what = 'ticket type ' + $(if ($null -eq $tt) { 'unknown' } else { [string][int]$tt })
                Say "  Ticket $n is $what, not a project (57)." 'Yellow'
                $ans = ''; continue
            }
            return $t
        }

        # anything else is a name search
        Say "  searching project names for '$ans' - this pages the ticket list, so give it a moment..." 'DarkGray'
        $all = Get-ProjectList
        $hits = @($all | Where-Object { ([string](Get-Val $_ 'summary')) -like ('*' + $ans + '*') })
        if ($hits.Count -eq 0) {
            Say "  Nothing matched '$ans' across $($all.Count) project(s)." 'Yellow'
            $ans = ''; continue
        }
        $sorted = @($hits | Sort-Object @{Expression={[string](Get-Val $_ 'summary')}})
        Write-Host ''
        Say "  $($sorted.Count) match(es):" 'Cyan'
        $i = 0
        foreach ($h in $sorted) {
            $i++
            Write-Host ("   {0,3}. #{1,-7} {2}" -f $i, [int]$h.id, [string](Get-Val $h 'summary'))
            if ($i -ge 40) { Say '   ...more matches not shown - be more specific' 'DarkGray'; break }
        }
        Write-Host ''
        $pick = Read-Host '  Number from the list, or a ticket id'
        if ([string]::IsNullOrWhiteSpace($pick)) { return $null }
        $k = 0
        if ([int]::TryParse($pick.Trim(), [ref]$k) -and $k -ge 1 -and $k -le [Math]::Min(40, $sorted.Count)) {
            return (Get-Ticket -Id ([int]$sorted[$k - 1].id))
        }
        $ans = $pick
    }
}

Head 'PROJECT'
$proj = Resolve-Project -Answer $ProjectId
if ($null -eq $proj) { Say '  Nothing chosen - stopping.' 'Yellow'; return }
$ProjTicketId = [int](Get-Val $proj 'id')
Say "  #$ProjTicketId  $([string](Get-Val $proj 'summary'))" 'Green'

# ---------------------------------------------------------------------------
#  ITS TASKS
# ---------------------------------------------------------------------------
# Three tiers, cheapest first, and every tier is verified before it is
# trusted: Halo IGNORES a query parameter it does not know rather than
# rejecting it, so an unsupported filter quietly returns the whole tenant and
# looks like a spectacular result.

function Test-IsChild {
    param($t, [int]$Parent)
    $tt = Get-Val $t 'tickettype_id'
    if ($null -eq $tt -or [int]$tt -ne $TYPE_TASK) { return $false }
    $p = Get-Val $t 'parent_id'
    if ($null -eq $p -or [int]$p -ne $Parent) { return $false }
    return $true
}

function Get-ChildIds {
    param($Project)
    $pid_ = [int](Get-Val $Project 'id')
    $ids = @{}

    # tier 1 - the project record may already list its children
    foreach ($f in @('childids','child_ids','children','childtickets')) {
        foreach ($c in @(Get-Val $Project $f)) {
            if ($null -eq $c) { continue }
            if ($c -is [int] -or $c -is [long] -or $c -is [string]) {
                $v = 0
                if ([int]::TryParse([string]$c, [ref]$v) -and $v -gt 0) { $ids[$v] = $true }
            } else {
                $childId = Get-Val $c 'id'
                if ($null -ne $childId) { $ids[[int]$childId] = $true }
            }
        }
    }
    if ($ids.Count -gt 0) {
        Say "      the project record lists $($ids.Count) child ticket(s)" 'DarkGray'
        return @($ids.Keys)
    }

    # tier 2 - ask the server to filter, then prove that it did
    foreach ($qs in @("parent_id=$pid_&open_only=false", "parentid=$pid_&open_only=false")) {
        try {
            $r = Invoke-RestMethod -Uri "$ApiBase/Tickets?$qs" -Headers $script:Headers -Method Get -TimeoutSec 300
            $b = @(Get-Collection $r 'tickets')
            if ($b.Count -eq 0) { continue }
            $good = @($b | Where-Object { Test-IsChild $_ $pid_ })
            # If the filter was honoured, everything coming back belongs to
            # this project. One stray record means the parameter was ignored
            # and we are staring at the whole tenant - discard it.
            if ($good.Count -eq $b.Count) {
                Say "      server-side filter ($qs) returned $($good.Count) task(s)" 'DarkGray'
                $out = @{}
                foreach ($g in $good) { $out[[int]$g.id] = $true }
                return @($out.Keys)
            }
            Say "      $qs was ignored by the server - falling back" 'DarkGray'
        } catch { }
    }

    # tier 3 - page the list and filter here. Slow, but it always works.
    Say '      paging the ticket list to find the tasks...' 'DarkGray'
    $mine = @{}; $unknown = @{}
    $pg = 1; $total = $null; $fetched = 0; $seenTasks = 0
    while ($true) {
        try {
            $r = Invoke-RestMethod -Uri "$ApiBase/Tickets?pageinate=true&page_size=$PAGE&page_no=$pg" `
                 -Headers $script:Headers -Method Get -TimeoutSec 600
        } catch { break }
        $b = Get-Collection $r 'tickets'
        if ($null -eq $total) {
            $rc = Get-Val $r 'record_count'
            if ($null -ne $rc) { $total = [int]$rc }
        }
        if ($b.Count -eq 0) { break }
        $fetched += $b.Count
        foreach ($t in $b) {
            $tt = Get-Val $t 'tickettype_id'
            if ($null -eq $tt -or [int]$tt -ne $TYPE_TASK) { continue }
            $seenTasks++
            $p = Get-Val $t 'parent_id'
            if ($null -eq $p) { $unknown[[int]$t.id] = $true }
            elseif ([int]$p -eq $pid_) { $mine[[int]$t.id] = $true }
        }
        $pct = 1
        if ($total -and $total -gt 0) { $pct = [int]([Math]::Min(100, $fetched / $total * 100)) }
        Write-Progress -Id 1 -Activity 'Reading the ticket list' -Status "page $pg - $fetched of $total" -PercentComplete $pct
        if ($b.Count -lt $PAGE) { break }
        if ($null -ne $total -and $fetched -ge $total) { break }
        $pg++
        if ($pg -gt 800) { break }
    }
    Clear-Step

    if ($unknown.Count -eq 0) {
        Say "      $seenTasks project task(s) in the tenant, $($mine.Count) under this project" 'DarkGray'
        return @($mine.Keys)
    }

    # the list payload carried no parent_id, so those have to be read one by one
    Say "      the list omits parent_id - checking $($unknown.Count) task(s) individually" 'Yellow'
    $i = 0; $n = $unknown.Count
    foreach ($k in @($unknown.Keys)) {
        $i++
        if ($i % 5 -eq 0 -or $i -eq $n) { Set-Step 'Checking tasks' $i $n }
        $d = Get-Ticket -Id ([int]$k)
        if ($null -ne $d -and (Test-IsChild $d $pid_)) { $mine[[int]$k] = $true }
    }
    Clear-Step
    return @($mine.Keys)
}

Head 'TASKS'
$childIds = @(Get-ChildIds $proj)
Say "  $($childIds.Count) task(s) to read" 'DarkGray'

# The list payload has no startdate, so each task needs an individual read.
$taskRecords = New-Object System.Collections.ArrayList
$i = 0; $n = $childIds.Count; $sw = [Diagnostics.Stopwatch]::StartNew()
foreach ($k in ($childIds | Sort-Object)) {
    $i++
    if ($i % 5 -eq 0 -or $i -eq $n) {
        $rate = 0.0
        if ($sw.Elapsed.TotalSeconds -gt 0) { $rate = $i / $sw.Elapsed.TotalSeconds }
        $left = 0
        if ($rate -gt 0) { $left = [int](($n - $i) / $rate) }
        Set-Step 'Reading tasks' $i $n "$i of $n - about $left second(s) left"
    }
    $d = Get-Ticket -Id ([int]$k)
    if ($null -ne $d) { [void]$taskRecords.Add($d) }
}
Clear-Step
if ($n -gt 0) { Say "  read $($taskRecords.Count) task(s) in $([int]$sw.Elapsed.TotalSeconds)s" 'DarkGray' }

# ---------------------------------------------------------------------------
#  BUILD THE MODEL
# ---------------------------------------------------------------------------
Head 'BUILDING'

$today   = (Get-Date).ToString('yyyy-MM-dd')
$nowD    = (Get-Date).Date
$soonIso = $nowD.AddDays($DueSoonDays).ToString('yyyy-MM-dd')

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
        start     = Format-D (Get-Val $t 'startdate')
        target    = Format-D (Get-Val $t 'targetdate')
        deadline  = Format-D (Get-Val $t 'deadlinedate')
        closedOn  = $closedOn
        closed    = $isClosed
        parentId  = $(if ($null -eq (Get-Val $t 'parent_id')) { 0 } else { [int](Get-Val $t 'parent_id') })
        milestone = $(if ($null -eq (Get-Val $t 'milestone_id')) { 0 } else { [int](Get-Val $t 'milestone_id') })
        client    = [string](Get-Val $t 'client_name')
        site      = [string](Get-Val $t 'site_name')
        user      = [string](Get-Val $t 'user_name')
        category  = [string](Get-Val $t 'category_1')
        details   = Strip-Html (Get-Val $t 'details')
        isProject = $isProject
    }
}

function Get-Flags {
    param($r)
    $f = @()
    if ($r.closed) { return $f }
    if (-not $r.start)  { $f += 'no start' }
    if (-not $r.target) { $f += 'no end' }
    if ($r.start -and $r.target -and $r.target -lt $r.start) { $f += 'end before start' }
    if ($r.target -and $r.target -lt $today) { $f += 'overdue' }
    return $f
}

$P = New-Row $proj $true
$P.flags = Get-Flags $P

# milestones ride on the project ticket, with UNDERSCORED date fields
$milestones = New-Object System.Collections.ArrayList
foreach ($m in @(Get-Val $proj 'milestones')) {
    if ($null -eq $m) { continue }
    [void]$milestones.Add(@{
        id     = [int](Get-Val $m 'id')
        name   = [string](Get-Val $m 'name')
        seq    = $(if ($null -eq (Get-Val $m 'sequence')) { 0 } else { [int](Get-Val $m 'sequence') })
        start  = Format-D (Get-Val $m 'start_date')
        target = Format-D (Get-Val $m 'target_date')
    })
}

$tasks = New-Object System.Collections.ArrayList
foreach ($t in $taskRecords) {
    $r = New-Row $t $false
    $r.flags = Get-Flags $r
    [void]$tasks.Add($r)
}

# --- weighting -------------------------------------------------------------
# Progress as a flat count of closed tasks flatters a project that has ticked
# off a dozen half-day jobs and not started the three-week one. Weight each
# task by its own length in days instead, and fall back to a single day when a
# task has no usable window so it still counts for something.
foreach ($k in $tasks) {
    $d = 1.0
    $a = To-Date $k.start
    $b = To-Date $k.target
    if ($null -ne $a -and $null -ne $b -and $b -ge $a) { $d = ($b - $a).TotalDays + 1 }
    $k.days = [Math]::Max(1.0, $d)
}

$totTask  = $tasks.Count
$closedT  = @($tasks | Where-Object { $_.closed })
$openT    = @($tasks | Where-Object { -not $_.closed })
$overdueT = @($openT | Where-Object { $_.target -and $_.target -lt $today })
$dueSoonT = @($openT | Where-Object { $_.target -and $_.target -ge $today -and $_.target -le $soonIso })
$unsched  = @($openT | Where-Object { -not $_.start -or -not $_.target })
$noOwner  = @($openT | Where-Object { $_.agent -eq 'Unassigned' })

$sumDays = 0.0; $doneDays = 0.0
foreach ($k in $tasks) { $sumDays += $k.days; if ($k.closed) { $doneDays += $k.days } }
$pctCount = 0
if ($totTask -gt 0) { $pctCount = [int][Math]::Round($closedT.Count / $totTask * 100) }
$pctDays = 0
if ($sumDays -gt 0) { $pctDays = [int][Math]::Round($doneDays / $sumDays * 100) }

# --- the window ------------------------------------------------------------
# A project record's own dates are frequently stale or blank, so the report
# leads with the window rolled up from the tasks - and says plainly when the
# two disagree rather than quietly picking one.
$ss = @(); $tt2 = @()
foreach ($k in $tasks) {
    if ($k.start)    { $ss  += $k.start }
    if ($k.target)   { $tt2 += $k.target }
    if ($k.closedOn) { $tt2 += $k.closedOn }
}
foreach ($m in $milestones) {
    if ($m.start)  { $ss  += $m.start }
    if ($m.target) { $tt2 += $m.target }
}
# @(...) matters: indexing a SINGLE-element array without it returns the first
# CHARACTER of the string, which silently produces a garbage window.
$rollStart = ''; $rollTarget = ''
if ($ss.Count  -gt 0) { $rollStart  = @($ss  | Sort-Object)[0] }
if ($tt2.Count -gt 0) { $rollTarget = @($tt2 | Sort-Object)[-1] }
$effStart  = $(if ($rollStart)  { $rollStart }  else { $P.start })
$effTarget = $(if ($rollTarget) { $rollTarget } else { $P.target })
if ($P.start  -and (-not $effStart  -or $P.start  -lt $effStart))  { $effStart  = $P.start }
if ($P.target -and (-not $effTarget -or $P.target -gt $effTarget)) { $effTarget = $P.target }

$elapsed = -1
$spanDays = 0; $goneDays = 0; $leftDays = 0
$ad = To-Date $effStart
$bd = To-Date $effTarget
if ($null -ne $ad -and $null -ne $bd -and $bd -ge $ad) {
    $spanDays = [int](($bd - $ad).TotalDays) + 1
    $goneDays = [int](($nowD - $ad).TotalDays)
    if ($goneDays -lt 0) { $goneDays = 0 }
    if ($goneDays -gt $spanDays) { $goneDays = $spanDays }
    $leftDays = [int](($bd - $nowD).TotalDays)
    if ($spanDays -gt 0) {
        $elapsed = [int][Math]::Round($goneDays / [double]$spanDays * 100)
        if ($elapsed -lt 0)   { $elapsed = 0 }
        if ($elapsed -gt 100) { $elapsed = 100 }
    }
}

# --- the RAG verdict -------------------------------------------------------
# One colour, and always the reasons beside it. The colour never carries the
# meaning on its own - the word and the reasons do - so it survives being
# printed in black and white, or read by someone who does not see red and
# green apart.
$rag = 'green'; $ragWord = 'Green'; $reasons = @()

if ($P.closed) {
    $rag = 'done'; $ragWord = 'Complete'
    $reasons += 'The project record is closed'
    if ($openT.Count -gt 0) {
        $rag = 'amber'; $ragWord = 'Amber'
        $reasons += "$($openT.Count) task(s) are still open underneath it"
    }
}
elseif ($totTask -eq 0) {
    $rag = 'amber'; $ragWord = 'Amber'
    $reasons += 'The project has no tasks, so there is nothing to measure progress against'
    if ($effTarget -and $effTarget -lt $today) {
        $rag = 'red'; $ragWord = 'Red'
        $reasons += "The project record is past its end date ($(Show-D $effTarget))"
    }
}
elseif ($pctDays -ge 100) {
    $rag = 'green'; $ragWord = 'Green'
    $reasons += 'Every task is closed - the project record itself has not been closed yet'
}
else {
    $behind = -1
    if ($elapsed -ge 0) { $behind = $elapsed - $pctDays }

    if ($effTarget -and $effTarget -lt $today) {
        $rag = 'red'; $ragWord = 'Red'
        $reasons += "Past its end date ($(Show-D $effTarget)) with $($openT.Count) task(s) still open"
    }
    if ($behind -ge $RedBehindBy) {
        $rag = 'red'; $ragWord = 'Red'
        $reasons += "$behind points behind schedule - $pctDays% of the work done against $elapsed% of the time gone"
    }
    elseif ($behind -ge $AmberBehindBy) {
        if ($rag -ne 'red') { $rag = 'amber'; $ragWord = 'Amber' }
        $reasons += "$behind points behind schedule - $pctDays% of the work done against $elapsed% of the time gone"
    }
    if ($overdueT.Count -gt 0) {
        if ($rag -eq 'green') { $rag = 'amber'; $ragWord = 'Amber' }
        $reasons += "$($overdueT.Count) open task(s) are past their end date"
    }
    if ($elapsed -lt 0) {
        if ($rag -eq 'green') { $rag = 'amber'; $ragWord = 'Amber' }
        $reasons += 'There is no usable schedule to judge the progress against'
    }
    if ($unsched.Count -gt 0) {
        if ($rag -eq 'green') { $rag = 'amber'; $ragWord = 'Amber' }
        $reasons += "$($unsched.Count) open task(s) have no start or no end date"
    }
    if ($noOwner.Count -gt 0) {
        if ($rag -eq 'green') { $rag = 'amber'; $ragWord = 'Amber' }
        $reasons += "$($noOwner.Count) open task(s) have nobody assigned"
    }
    if ($reasons.Count -eq 0) {
        $reasons += "On track - $pctDays% of the work done against $elapsed% of the time gone"
        if ($effTarget -and $leftDays -ge 0) { $reasons += "$leftDays day(s) left to $(Show-D $effTarget)" }
    }
}
Say "  RAG: $ragWord" $(if ($rag -eq 'red') { 'Red' } elseif ($rag -eq 'amber') { 'Yellow' } else { 'Green' })
foreach ($r in $reasons) { Say "    - $r" 'DarkGray' }

# --- per-agent -------------------------------------------------------------
$byAgent = @{}
function Touch-Agent {
    param([string]$a)
    if (-not $byAgent.ContainsKey($a)) {
        $byAgent[$a] = @{ open=0; closed=0; overdue=0; days=0.0; doneDays=0.0; next='' }
    }
}
Touch-Agent $P.agent
foreach ($k in $tasks) {
    Touch-Agent $k.agent
    $e = $byAgent[$k.agent]
    $e.days += $k.days
    if ($k.closed) { $e.closed++; $e.doneDays += $k.days }
    else {
        $e.open++
        if ($k.target -and $k.target -lt $today) { $e.overdue++ }
        if ($k.target -and ((-not $e.next) -or $k.target -lt $e.next)) { $e.next = $k.target }
    }
}
$agentOrder = @($byAgent.GetEnumerator() |
    Sort-Object @{Expression={$_.Value.open}; Descending=$true}, `
                @{Expression={$_.Value.closed}; Descending=$true}, Name |
    ForEach-Object { $_.Name })
$slotOf = @{}
$slot = 1
foreach ($a in $agentOrder) {
    if ($slot -le 8) { $slotOf[$a] = $slot; $slot++ } else { $slotOf[$a] = 9 }
}

# --- milestone roll-up -----------------------------------------------------
$msName = @{}
foreach ($m in $milestones) {
    $msName[$m.id] = $m.name
    $kids = @($tasks | Where-Object { $_.milestone -eq $m.id })
    $m.taskCount = $kids.Count
    $m.doneCount = @($kids | Where-Object { $_.closed }).Count
    $md = 0.0; $mdone = 0.0
    foreach ($k in $kids) { $md += $k.days; if ($k.closed) { $mdone += $k.days } }
    $m.pct = 0
    if ($md -gt 0) { $m.pct = [int][Math]::Round($mdone / $md * 100) }
    $m.overdue = @($kids | Where-Object { -not $_.closed -and $_.target -and $_.target -lt $today }).Count
}
$unMs = @($tasks | Where-Object { $_.milestone -le 0 -or -not $msName.ContainsKey($_.milestone) })

# --- timeline scale --------------------------------------------------------
function Get-Scale {
    param($Dates)
    $d = @($Dates | Where-Object { $_ })
    if ($d.Count -eq 0) { return $null }
    $srt = @($d | Sort-Object)
    $mn = To-Date $srt[0]
    $mx = To-Date $srt[-1]
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
    $da = To-Date $x
    $db = To-Date $y
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
    param($Scale)
    if ($null -eq $Scale) { return -1 }
    $t = (Get-Date).Date
    if ($t -lt $Scale.a -or $t -gt $Scale.b) { return -1 }
    return ($t - $Scale.a).TotalDays / $Scale.span * 100
}
function Axis-Row {
    param($Scale, [int]$Before)
    if ($null -eq $Scale) { return '' }
    # tick density has to suit the column, not just the span - 12 monthly ticks
    # in a quarter-width column collide into mush
    $step = 1
    if ($Scale.span -gt 200)  { $step = 2 }
    if ($Scale.span -gt 400)  { $step = 3 }
    if ($Scale.span -gt 800)  { $step = 6 }
    if ($Scale.span -gt 1800) { $step = 12 }
    $out = '<tr class="axrow"><td colspan="' + $Before + '" style="border-bottom:1px solid var(--grid);padding:0"></td><td class="ax">'
    $c = $Scale.a
    $first = $true
    while ($c -le $Scale.b) {
        $pct = ($c - $Scale.a).TotalDays / $Scale.span * 100
        $lab = $c.ToString('MMM')
        if ($c.Month -eq 1 -or $first) { $lab += ' ' + $c.ToString('yy') }
        $out += '<span style="left:' + ('{0:F2}' -f $pct) + '%">' + $lab + '</span>'
        $first = $false
        $c = $c.AddMonths($step)
    }
    return ($out + '</td></tr>')
}

$sd = @($P.start, $P.target, $effStart, $effTarget, $today)
foreach ($k in $tasks)      { $sd += $k.start; $sd += $k.target; $sd += $k.closedOn }
foreach ($m in $milestones) { $sd += $m.start; $sd += $m.target }
$scale = Get-Scale $sd
$tp = Today-In $scale
$todayMark = ''
if ($tp -ge 0) { $todayMark = '<u style="left:' + ('{0:F2}' -f $tp) + '%"></u>' }

# ---------------------------------------------------------------------------
#  EMIT THE HTML
# ---------------------------------------------------------------------------
Head 'WRITING THE REPORT'

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
 --r-red:#d03b3b;   --r-red-bg:#f7e2e2;
 --r-amber:#b37400; --r-amber-bg:#fbeacb;
 --r-green:#0ca30c; --r-green-bg:#d8f0d8;
 --r-done:#2a78d6;  --r-done-bg:#dbeafd;
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
  --r-red:#e66767;   --r-red-bg:#3d1c1c;
  --r-amber:#fab219; --r-amber-bg:#3f3110;
  --r-green:#3fbf3f; --r-green-bg:#173318;
  --r-done:#3987e5;  --r-done-bg:#18334f;
 }
}
:root[data-theme="dark"] .viz-root,:root[data-theme="dark"] body{
 color-scheme:dark;
 --surface-1:#1a1a19; --plane:#0d0d0d;
 --text-primary:#ffffff; --text-secondary:#c3c2b7; --muted:#898781;
 --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,0.10);
 --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500;
 --s5:#d55181; --s6:#008300; --s7:#9085e9; --s8:#e66767; --s9:#898781;
 --r-red:#e66767;   --r-red-bg:#3d1c1c;
 --r-amber:#fab219; --r-amber-bg:#3f3110;
 --r-green:#3fbf3f; --r-green-bg:#173318;
 --r-done:#3987e5;  --r-done-bg:#18334f;
}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--text-primary);
 font:15px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:30px 20px 90px}
h1{font-size:25px;margin:0 0 4px;letter-spacing:-.01em}
h2{font-size:18px;margin:34px 0 11px;letter-spacing:-.005em}
h3{font-size:13px;margin:0 0 10px;color:var(--muted);font-weight:600;
 text-transform:uppercase;letter-spacing:.5px}
p{margin:0 0 12px;max-width:80ch}
.lede{color:var(--text-secondary);font-size:14px;margin-bottom:20px}
.card{background:var(--surface-1);border:1px solid var(--ring);border-radius:12px;
 padding:18px 20px;margin-bottom:14px}
.sub{color:var(--text-secondary);font-size:13px}
.kv{display:flex;flex-wrap:wrap;gap:8px 26px;font-size:13.5px}
.kv div{white-space:nowrap}
.kv b{color:var(--muted);font-weight:600;font-size:10.5px;text-transform:uppercase;
 letter-spacing:.4px;margin-right:7px}
.rag{border-radius:14px;padding:17px 20px;margin:18px 0 16px;border:1px solid}
.rag.red{background:var(--r-red-bg);border-color:var(--r-red)}
.rag.amber{background:var(--r-amber-bg);border-color:var(--r-amber)}
.rag.green{background:var(--r-green-bg);border-color:var(--r-green)}
.rag.done{background:var(--r-done-bg);border-color:var(--r-done)}
.ragtop{display:flex;align-items:center;gap:13px;flex-wrap:wrap}
.ragbadge{font-size:13px;font-weight:700;letter-spacing:.9px;text-transform:uppercase;
 padding:6px 15px;border-radius:99px;color:#fff;white-space:nowrap}
.rag.red .ragbadge{background:var(--r-red)}
.rag.amber .ragbadge{background:var(--r-amber)}
.rag.green .ragbadge{background:var(--r-green)}
.rag.done .ragbadge{background:var(--r-done)}
.ragline{font-size:15.5px;font-weight:650}
.rag ul{margin:11px 0 0;padding-left:20px;font-size:13.5px;color:var(--text-secondary)}
.rag li{margin:3px 0}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(132px,1fr));gap:11px;margin-bottom:16px}
.tile{background:var(--surface-1);border:1px solid var(--ring);border-radius:11px;padding:13px 15px}
.tile b{display:block;font-size:25px;line-height:1.05;letter-spacing:-.02em;font-weight:600}
.tile span{display:block;color:var(--text-secondary);font-size:12px;margin-top:4px}
.tile.warn b{color:var(--critical)}
.pline{margin:16px 0 4px}
.pline:first-of-type{margin-top:4px}
.plab{display:flex;justify-content:space-between;font-size:12.5px;color:var(--text-secondary);
 margin-bottom:5px;gap:12px}
.plab b{font-variant-numeric:tabular-nums;color:var(--text-primary);font-size:14px}
.big{position:relative;height:20px;background:var(--grid);border-radius:10px;overflow:hidden}
.big b{position:absolute;left:0;top:0;height:20px;border-radius:10px;display:block}
.big.work b{background:var(--r-green)}
.big.work.red b{background:var(--r-red)}
.big.work.amber b{background:var(--r-amber)}
.big.time b{background:var(--s9);opacity:.5}
.marker{position:relative;height:15px;margin-top:2px}
.marker i{position:absolute;top:2px;transform:translateX(-50%);font-size:10.5px;
 color:var(--muted);white-space:nowrap;font-style:normal}
.marker u{position:absolute;top:-23px;width:2px;height:22px;background:var(--critical);
 border-radius:1px;text-decoration:none}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;
 color:var(--muted);font-weight:600;padding:7px 8px;border-bottom:1px solid var(--grid)}
td{padding:6px 8px;border-bottom:1px solid var(--grid);vertical-align:top}
tr:last-child td{border-bottom:none}
td.n{text-align:right;font-variant-numeric:tabular-nums}
td.d{white-space:nowrap;font-variant-numeric:tabular-nums;font-size:12.5px;color:var(--text-secondary)}
td.mono{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:var(--muted)}
tr.done td{color:var(--muted)}
tr.grp td{border-bottom:none}
.track{position:relative;height:10px;background:var(--grid);border-radius:5px;min-width:110px}
.track b{position:absolute;top:0;height:10px;border-radius:4px;display:block}
.track u{position:absolute;top:-3px;width:2px;height:16px;background:var(--critical);
 border-radius:1px;text-decoration:none}
.mtrack{height:14px}.mtrack b{height:14px;border-radius:5px}
td.ax{position:relative;height:16px;padding:0 8px 4px;border-bottom:1px solid var(--grid)}
td.ax span{position:absolute;bottom:3px;transform:translateX(-50%);white-space:nowrap;
 font-size:10px;color:var(--muted)}
.dot{display:inline-block;width:9px;height:9px;border-radius:3px;margin-right:7px}
.meter{display:inline-block;width:70px;height:8px;border-radius:4px;vertical-align:1px;
 overflow:hidden;background:var(--grid)}
.meter b{display:block;height:8px;border-radius:4px;background:var(--s1)}
.pct{font-size:12px;font-variant-numeric:tabular-nums;color:var(--text-secondary);margin-left:7px}
.flag{font-size:10px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;
 padding:2px 6px;border-radius:99px;border:1px solid;white-space:nowrap;margin-left:5px;
 display:inline-block}
.f-bad{color:var(--critical);border-color:var(--critical)}
.f-mut{color:var(--text-secondary);border-color:var(--axis)}
.f-ok{color:var(--good);border-color:var(--good)}
.desc{font-size:13.5px;color:var(--text-secondary);max-width:92ch;white-space:pre-wrap}
.bar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;position:sticky;top:0;z-index:30;
 background:var(--plane);padding:11px 0;margin-bottom:8px;border-bottom:1px solid var(--grid)}
.bar label{font-size:13px;display:flex;align-items:center;gap:6px;cursor:pointer;white-space:nowrap}
select,input[type=search],button{font:inherit;font-size:13px;padding:5px 9px;
 border:1px solid var(--ring);border-radius:7px;background:var(--surface-1);
 color:var(--text-primary)}
button{cursor:pointer;color:var(--text-secondary)}
button:hover{color:var(--text-primary)}
.pid{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:var(--muted)}
.hidden{display:none!important}
.top{position:fixed;right:18px;bottom:18px;z-index:50;border-radius:99px;padding:8px 14px;
 box-shadow:0 4px 16px rgba(0,0,0,.18)}
@media print{
 .bar,.top{display:none}
 body{background:#fff}
 tr{break-inside:avoid}
 .rag{border-width:2px}
}
@media (max-width:760px){.kv{gap:6px 16px}}
</style>
'@

$js = @'
<script>
(function(){
 var q=document.getElementById('fq'), ag=document.getElementById('fagent'),
     sc=document.getElementById('fclosed'), pr=document.getElementById('fprob'),
     cnt=document.getElementById('fcount'), pb=document.getElementById('pbtn');
 var rows=[].slice.call(document.querySelectorAll('tr.task'));
 function apply(){
  var term=q.value.trim().toLowerCase(), a=ag.value,
      showClosed=sc.checked, onlyProb=pr.checked, shown=0;
  rows.forEach(function(r){
   var ok=true;
   if(!showClosed && r.dataset.closed==='1') ok=false;
   if(a && r.dataset.agent!==a) ok=false;
   if(onlyProb && r.dataset.prob!=='1') ok=false;
   if(term && r.dataset.q.indexOf(term)<0) ok=false;
   r.classList.toggle('hidden',!ok); if(ok) shown++;
  });
  cnt.textContent=shown+' of '+rows.length+' tasks shown';
  // a milestone heading with nothing under it is noise - it goes with its rows
  document.querySelectorAll('tr.grp').forEach(function(g){
   var any=false,n=g.nextElementSibling;
   while(n && !n.classList.contains('grp')){
    if(n.classList.contains('task') && !n.classList.contains('hidden')) any=true;
    n=n.nextElementSibling;
   }
   g.classList.toggle('hidden',!any);
  });
 }
 [sc,pr].forEach(function(e){e.addEventListener('change',apply)});
 ag.addEventListener('change',apply);
 q.addEventListener('input',apply);
 // print the whole plan, not whatever happens to be filtered on screen
 if(pb) pb.addEventListener('click',function(){
  sc.checked=true; pr.checked=false; ag.value=''; q.value=''; apply(); window.print();
 });
 document.getElementById('totop').addEventListener('click',function(){
  window.scrollTo({top:0,behavior:'smooth'});
 });
 apply();
})();
</script>
'@

$stamp = Get-Date -Format 'dd/MM/yyyy HH:mm'

W '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
W '<meta name="viewport" content="width=device-width, initial-scale=1">'
W ('<title>' + (Esc ($P.summary + ' - project status ' + $stamp)) + '</title>')
W $css
W '</head><body class="viz-root"><div class="wrap">'

# --- heading ---------------------------------------------------------------
W ('<h1>' + (Esc $P.summary) + '</h1>')
W ('<p class="lede"><span class="pid">Project #' + $P.id + '</span> &middot; ' + (Esc $Tenant) +
   ' &middot; status as at ' + (Esc $stamp) +
   ' &middot; read-only snapshot, nothing was changed in Halo.</p>')

W '<div class="card"><div class="kv">'
$kvs = @()
if ($P.client)   { $kvs += @{k='Client';   v=$P.client} }
if ($P.site)     { $kvs += @{k='Site';     v=$P.site} }
if ($P.user)     { $kvs += @{k='Contact';  v=$P.user} }
$kvs += @{k='Project manager'; v=$P.agent}
if ($P.status)   { $kvs += @{k='Status';   v=$P.status} }
if ($P.category) { $kvs += @{k='Category'; v=$P.category} }
$kvs += @{k='Window'; v=$(if ($effStart -or $effTarget) { (Show-D $effStart) + ' to ' + (Show-D $effTarget) } else { 'no dates' })}
if ($P.start -or $P.target) {
    $own = (Show-D $P.start) + ' to ' + (Show-D $P.target)
    if ($own -ne ((Show-D $effStart) + ' to ' + (Show-D $effTarget))) {
        $kvs += @{k='Project record says'; v=$own}
    }
}
if ($P.deadline) { $kvs += @{k='Deadline'; v=(Show-D $P.deadline)} }
foreach ($x in $kvs) { W ('<div><b>' + (Esc $x.k) + '</b>' + (Esc $x.v) + '</div>') }
W '</div></div>'

# --- RAG -------------------------------------------------------------------
$ragBlurb = 'On track'
if ($rag -eq 'red')   { $ragBlurb = 'Needs intervention' }
if ($rag -eq 'amber') { $ragBlurb = 'Watch this one' }
if ($rag -eq 'done')  { $ragBlurb = 'Delivered' }
W ('<div class="rag ' + $rag + '">')
W ('<div class="ragtop"><span class="ragbadge">' + (Esc $ragWord) + '</span>' +
   '<span class="ragline">' + (Esc $ragBlurb) + '</span></div>')
W '<ul>'
foreach ($r in $reasons) { W ('<li>' + (Esc $r) + '</li>') }
W '</ul></div>'

# --- progress --------------------------------------------------------------
W '<div class="card"><h3>Progress</h3>'
if ($totTask -eq 0) {
    # an empty bar reading 0% would look like a project that has started and
    # achieved nothing, which is a different thing entirely from one nobody
    # has broken down yet
    W ('<p class="sub" style="margin:4px 0 0">There are no tasks under this project, so there is no ' +
       'completed work to measure. Until the plan is broken down in Halo the only thing this report ' +
       'can say about progress is that nothing is being tracked.</p>')
} else {
    $workCls = 'big work'
    if ($rag -eq 'red')   { $workCls += ' red' }
    if ($rag -eq 'amber') { $workCls += ' amber' }
    W '<div class="pline">'
    W ('<div class="plab"><span>Work completed &mdash; closed tasks, weighted by how long each one runs</span><b>' +
       $pctDays + '%</b></div>')
    W ('<div class="' + $workCls + '"><b style="width:' + $pctDays + '%"></b></div>')
    W ('<div class="sub" style="margin-top:6px">' + $closedT.Count + ' of ' + $totTask + ' tasks closed (' +
       $pctCount + '% by simple count) &middot; ' + [int][Math]::Round($doneDays) + ' of ' +
       [int][Math]::Round($sumDays) + ' task-days done</div>')
    W '</div>'
}

if ($elapsed -ge 0) {
    W '<div class="pline">'
    W ('<div class="plab"><span>Schedule elapsed &mdash; ' + (Show-D $effStart) + ' to ' + (Show-D $effTarget) +
       '</span><b>' + $elapsed + '%</b></div>')
    W ('<div class="big time"><b style="width:' + $elapsed + '%"></b></div>')
    W ('<div class="marker"><u style="left:' + $elapsed + '%"></u><i style="left:' + $elapsed + '%">today</i></div>')
    W '</div>'
    $gap = $elapsed - $pctDays
    $verdict = 'level with the schedule'
    if ($gap -ge $AmberBehindBy)     { $verdict = "$gap points behind the schedule" }
    elseif ($gap -gt 0)              { $verdict = "$gap points behind, which is within tolerance" }
    elseif ($gap -lt 0)              { $verdict = "$([Math]::Abs($gap)) points ahead of the schedule" }
    $leftTxt = "$leftDays day(s) left"
    if ($leftDays -lt 0) { $leftTxt = "$([Math]::Abs($leftDays)) day(s) past the end date" }
    W ('<p class="sub" style="margin:14px 0 0">A ' + $spanDays + ' day project: ' + $goneDays + ' gone, ' +
       $leftTxt + '. The work is ' + (Esc $verdict) + '.</p>')
} else {
    W ('<p class="sub" style="margin:14px 0 0">Neither this project nor its tasks carry a usable ' +
       'start-and-end window, so there is no schedule to measure the work against.</p>')
}
W '</div>'

# --- tiles -----------------------------------------------------------------
W '<div class="tiles">'
$tiles = @(
    @{v=$totTask;          l='Tasks';                      w=$false},
    @{v=$openT.Count;      l='Open';                       w=$false},
    @{v=$closedT.Count;    l='Closed';                     w=$false},
    @{v=$overdueT.Count;   l='Overdue';                    w=($overdueT.Count -gt 0)},
    @{v=$dueSoonT.Count;   l=("Due in $DueSoonDays days"); w=$false},
    @{v=$milestones.Count; l='Milestones';                 w=$false},
    @{v=$agentOrder.Count; l='People assigned';            w=$false},
    @{v=$unsched.Count;    l='Open, no dates';             w=($unsched.Count -gt 0)}
)
foreach ($t in $tiles) {
    W ('<div class="tile' + $(if ($t.w) { ' warn' } else { '' }) + '"><b>' + $t.v + '</b><span>' +
       (Esc $t.l) + '</span></div>')
}
W '</div>'

# --- description -----------------------------------------------------------
if ($P.details) {
    $dtl = $P.details
    if ($dtl.Length -gt 2500) { $dtl = $dtl.Substring(0,2500) + "`n..." }
    W '<h2>Description</h2><div class="card"><div class="desc">'
    W (Esc $dtl)
    W '</div></div>'
}

# --- the people ------------------------------------------------------------
W '<h2>Who is assigned</h2><div class="card">'
W '<table><thead><tr><th>Person</th><th class="n" style="width:66px">Open</th>'
W '<th class="n" style="width:74px">Closed</th><th class="n" style="width:80px">Overdue</th>'
W '<th style="width:190px">Their share done</th><th style="width:120px">Next due</th></tr></thead><tbody>'
foreach ($a in $agentOrder) {
    $e = $byAgent[$a]
    if (($e.open + $e.closed) -eq 0 -and $a -ne $P.agent) { continue }
    $ap = 0
    if ($e.days -gt 0) { $ap = [int][Math]::Round($e.doneDays / $e.days * 100) }
    $who = (Esc $a)
    if ($a -eq $P.agent) { $who += '<span class="flag f-mut">project manager</span>' }
    W ('<tr><td><span class="dot" style="background:var(--s' + $slotOf[$a] + ')"></span>' + $who + '</td>')
    W ('<td class="n">' + $e.open + '</td><td class="n">' + $e.closed + '</td>')
    W ('<td class="n">' + $(if ($e.overdue -gt 0) { '<span class="flag f-bad">' + $e.overdue + '</span>' } else { '0' }) + '</td>')
    if (($e.open + $e.closed) -gt 0) {
        W ('<td><span class="meter"><b style="width:' + $ap + '%"></b></span><span class="pct">' + $ap + '%</span></td>')
    } else {
        W '<td class="d">no tasks of their own</td>'
    }
    W ('<td class="d">' + $(if ($e.next) { Show-D $e.next } else { '&mdash;' }) + '</td></tr>')
}
W '</tbody></table></div>'

# --- milestones ------------------------------------------------------------
if ($milestones.Count -gt 0) {
    W '<h2>Milestones</h2><div class="card"><table>'
    W '<thead><tr><th>Milestone</th><th class="n" style="width:88px">Tasks</th>'
    W '<th style="width:148px">Done</th><th class="d" style="width:166px">Window</th>'
    W '<th style="width:28%">Timeline</th></tr></thead><tbody>'
    W (Axis-Row $scale 4)
    foreach ($m in @($milestones | Sort-Object @{Expression={$_.seq}}, @{Expression={$_.start}})) {
        W '<tr>'
        $tag = ''
        if ($m.overdue -gt 0) { $tag = '<span class="flag f-bad">' + $m.overdue + ' overdue</span>' }
        elseif ($m.taskCount -gt 0 -and $m.doneCount -eq $m.taskCount) { $tag = '<span class="flag f-ok">complete</span>' }
        W ('<td><b>' + (Esc $m.name) + '</b>' + $tag + '</td>')
        W ('<td class="n">' + $m.doneCount + ' / ' + $m.taskCount + '</td>')
        if ($m.taskCount -gt 0) {
            W ('<td><span class="meter"><b style="width:' + $m.pct + '%"></b></span><span class="pct">' + $m.pct + '%</span></td>')
        } else {
            W '<td class="d">no tasks</td>'
        }
        W ('<td class="d">' + $(if ($m.start -or $m.target) { (Show-D $m.start) + ' &rarr; ' + (Show-D $m.target) } else { '&mdash;' }) + '</td>')
        $bp = Bar-In $scale $m.start $m.target
        if ($bp) {
            $parts = $bp -split '\|'
            W ('<td><div class="track mtrack"><b style="left:' + $parts[0] + '%;width:' + $parts[1] +
               '%;background:var(--s1)"></b>' + $todayMark + '</div></td>')
        } else {
            W '<td><span class="flag f-mut">no dates</span></td>'
        }
        W '</tr>'
    }
    W '</tbody></table></div>'
}

# --- needs attention -------------------------------------------------------
if ($overdueT.Count -gt 0 -or $dueSoonT.Count -gt 0) {
    W '<h2>What needs attention</h2><div class="card"><table>'
    W '<thead><tr><th style="width:132px">When</th><th style="width:56px">ID</th><th>Task</th>'
    W '<th style="width:140px">Owner</th><th class="d" style="width:100px">Due</th></tr></thead><tbody>'
    foreach ($k in @($overdueT | Sort-Object @{Expression={$_.target}})) {
        $late = [int]((To-Date $today) - (To-Date $k.target)).TotalDays
        W ('<tr><td><span class="flag f-bad" style="margin-left:0">' + $late + ' day(s) late</span></td>' +
           '<td class="mono">' + $k.id + '</td><td>' + (Esc $k.summary) + '</td>' +
           '<td class="d">' + (Esc $k.agent) + '</td><td class="d">' + (Show-D $k.target) + '</td></tr>')
    }
    foreach ($k in @($dueSoonT | Sort-Object @{Expression={$_.target}})) {
        $inD = [int]((To-Date $k.target) - (To-Date $today)).TotalDays
        $when = 'in ' + $inD + ' day(s)'
        if ($inD -eq 0) { $when = 'today' }
        W ('<tr><td><span class="flag f-mut" style="margin-left:0">' + $when + '</span></td>' +
           '<td class="mono">' + $k.id + '</td><td>' + (Esc $k.summary) + '</td>' +
           '<td class="d">' + (Esc $k.agent) + '</td><td class="d">' + (Show-D $k.target) + '</td></tr>')
    }
    W '</tbody></table></div>'
}

# --- date problems ---------------------------------------------------------
$probRows = @()
if ($P.flags.Count -gt 0) {
    $probRows += @{ isProj=$true; id=$P.id; sum=$P.summary; ag=$P.agent; fl=$P.flags }
}
foreach ($k in $tasks) {
    if ($k.flags.Count -gt 0) { $probRows += @{ isProj=$false; id=$k.id; sum=$k.summary; ag=$k.agent; fl=$k.flags } }
}
W '<h2>Records with a date problem</h2><div class="card">'
if ($probRows.Count -eq 0) {
    W '<p style="margin:0"><span class="flag f-ok" style="margin-left:0">all clear</span> Every open record on this project has a sensible start and end date.</p>'
} else {
    W '<table><thead><tr><th style="width:210px">Problem</th><th style="width:56px">ID</th><th>Record</th>'
    W '<th style="width:140px">Owner</th></tr></thead><tbody>'
    foreach ($r in $probRows) {
        W ('<tr><td><span class="flag f-bad" style="margin-left:0">' + (Esc ($r.fl -join ', ')) + '</span></td>' +
           '<td class="mono">' + $r.id + '</td><td>' + (Esc $r.sum) +
           $(if ($r.isProj) { ' <span class="flag f-mut">the project itself</span>' } else { '' }) +
           '</td><td class="d">' + (Esc $r.ag) + '</td></tr>')
    }
    W '</tbody></table>'
}
W '</div>'

# --- every task ------------------------------------------------------------
W '<h2>Tasks</h2>'
W '<div class="bar">'
W '<label><input type="checkbox" id="fclosed" checked> Closed tasks</label>'
W '<label><input type="checkbox" id="fprob"> Only date problems</label>'
W '<label>Owner <select id="fagent"><option value="">everyone</option>'
foreach ($a in $agentOrder) {
    if (($byAgent[$a].open + $byAgent[$a].closed) -eq 0) { continue }
    W ('<option>' + (Esc $a) + '</option>')
}
W '</select></label>'
W '<input type="search" id="fq" placeholder="Search tasks&hellip;" style="min-width:200px">'
W '<button id="pbtn">Print / save as PDF</button>'
W '<span id="fcount" class="pid" style="margin-left:auto"></span>'
W '</div>'

W '<div class="card" style="padding:12px 16px">'
if ($totTask -eq 0) {
    W '<p style="margin:0" class="sub">This project has no tasks attached to it in Halo, so there is nothing to list.</p>'
} else {
    W '<table><thead><tr><th style="width:54px">ID</th><th>Task</th>'
    W '<th style="width:116px">Owner</th><th style="width:116px">Status</th>'
    W '<th style="width:90px">Start</th><th style="width:90px">End</th>'
    W '<th class="n" style="width:54px">Days</th><th style="width:94px">Closed</th>'
    W '<th style="width:22%">Timeline</th></tr></thead><tbody>'
    W (Axis-Row $scale 8)

    # grouped by milestone, in milestone order, with the unassigned ones last -
    # that is how the plan reads in Halo, so it is how it reads here
    $groups = New-Object System.Collections.ArrayList
    foreach ($m in @($milestones | Sort-Object @{Expression={$_.seq}}, @{Expression={$_.start}})) {
        $kids = @($tasks | Where-Object { $_.milestone -eq $m.id })
        if ($kids.Count -gt 0) { [void]$groups.Add(@{ name=$m.name; rows=$kids }) }
    }
    if ($unMs.Count -gt 0) {
        $nm = ''
        if ($milestones.Count -gt 0) { $nm = 'No milestone' }
        [void]$groups.Add(@{ name=$nm; rows=@($unMs) })
    }

    foreach ($g in $groups) {
        if ($g.name) {
            W ('<tr class="grp"><td colspan="9" style="padding-top:15px"><h3 style="margin:0">' +
               (Esc $g.name) + '</h3></td></tr>')
        }
        $ordered = @($g.rows | Sort-Object `
            @{Expression={[int]$_.closed}}, `
            @{Expression={if ($_.start) { $_.start } elseif ($_.target) { $_.target } else { '9999-99-99' }}}, `
            @{Expression={$_.id}})
        foreach ($k in $ordered) {
            $rc = 'task'
            if ($k.closed) { $rc += ' done' }
            W ('<tr class="' + $rc + '" data-closed="' + $(if ($k.closed) { 1 } else { 0 }) +
               '" data-prob="' + $(if ($k.flags.Count -gt 0) { 1 } else { 0 }) +
               '" data-agent="' + (Esc $k.agent) +
               '" data-q="' + (Esc (($k.summary + ' ' + $k.agent + ' ' + $k.status + ' ' + $k.id).ToLower())) + '">')
            W ('<td class="mono">' + $k.id + '</td>')
            W ('<td>' + (Esc $k.summary))
            if ($k.flags.Count -gt 0) { W ('<span class="flag f-bad">' + (Esc ($k.flags -join ', ')) + '</span>') }
            W '</td>'
            W ('<td class="d">' + (Esc $k.agent) + '</td>')
            W ('<td class="d">' + (Esc $k.status) + '</td>')
            W ('<td class="d">' + $(if ($k.start)  { Show-D $k.start }  else { '&mdash;' }) + '</td>')
            W ('<td class="d">' + $(if ($k.target) { Show-D $k.target } else { '&mdash;' }) + '</td>')
            W ('<td class="n">' + [int]$k.days + '</td>')
            $cl = '&mdash;'
            if ($k.closedOn) { $cl = Show-D $k.closedOn } elseif ($k.closed) { $cl = 'closed' }
            W ('<td class="d">' + $cl + '</td>')
            $bp = Bar-In $scale $k.start $k.target
            if ($bp) {
                $parts = $bp -split '\|'
                W ('<td><div class="track"><b style="left:' + $parts[0] + '%;width:' + $parts[1] +
                   '%;background:var(--s' + $slotOf[$k.agent] + ')' +
                   $(if ($k.closed) { ';opacity:.45' } else { '' }) + '"></b>' + $todayMark + '</div></td>')
            } else {
                W '<td></td>'
            }
            W '</tr>'
        }
    }
    W '</tbody></table>'
}
W '</div>'

W ('<p class="lede" style="margin-top:24px">Progress counts closed tasks weighted by each task&rsquo;s length in days, ' +
   'so a three-week task carries more than a half-day one; a task with no usable window counts as a single day. ' +
   'The RAG verdict compares that figure with how much of the schedule has elapsed - ' + $AmberBehindBy +
   ' points behind is amber, ' + $RedBehindBy + ' is red - and every reason behind it is listed in full above. ' +
   'Generated ' + (Esc $stamp) + ' from Halo ticket #' + $P.id + '.</p>')

W '<button class="top" id="totop">Back to top</button>'
W $js
W '</div></body></html>'

# --- write -----------------------------------------------------------------
$path = $OutFile
if ([string]::IsNullOrWhiteSpace($path)) {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($docs)) { $docs = $env:TEMP }
    $safe = ($P.summary -replace '[^A-Za-z0-9 \-_]', '').Trim()
    if ($safe.Length -gt 48) { $safe = $safe.Substring(0,48).Trim() }
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'project' }
    $path = Join-Path $docs ('Halo-Project-' + $P.id + '-' + $safe + '-' + (Get-Date -Format 'yyyyMMdd-HHmm') + '.html')
}
try {
    [System.IO.File]::WriteAllText($path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    Write-Host ''
    Say "  Report written: $path" 'Green'
    Say ("  {0} task(s), {1}% complete, RAG {2}, {3:N0} KB" -f $totTask, $pctDays, $ragWord, ((Get-Item $path).Length / 1KB)) 'DarkGray'
    if ($OpenWhenDone) { Start-Process $path }
}
catch {
    Say "  Could not write $path : $($_.Exception.Message)" 'Red'
    Say '  Set $OutFile to somewhere you can write and run again.' 'Yellow'
}
Write-Host ''
