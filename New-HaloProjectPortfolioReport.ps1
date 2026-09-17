<#
    New-HaloProjectPortfolioReport.ps1
    ==================================
    READ-ONLY. The two existing reports in one file, for running a project
    meeting from a single document.

    It reads every project and project task in the tenant - open and closed -
    and writes ONE self-contained HTML file containing:

        * a portfolio view   - a RAG verdict per project with the reasoning,
                               progress against schedule, who owns it, what is
                               late, and a timeline bar per project on a shared
                               scale, plus everything across the whole estate
                               that needs attention this week
        * a page per project - the full status report: RAG with its reasons,
                               work-completed against schedule-elapsed, who is
                               assigned, milestones, what needs attention, the
                               records with broken dates, and every task with
                               its own timeline bar

    Click a project in the portfolio view to open its page; "Back to the
    portfolio" returns. It is one file with no external dependencies, so it
    mails, prints and survives being opened off a share.

    Open in PowerShell ISE, fill in the CONFIGURATION block, press F5. Writes
    nothing to Halo and nothing to disk except the report.

    This is the SLOW one, because it is both reports at once: the ticket list
    omits startdate, so every project and every task needs an individual GET -
    roughly a second per six records. A tenant with a thousand project records
    is a three minute run. Both other scripts remain useful: the portfolio-only
    report is lighter, and the single-project report is quick because it never
    sweeps the tenant.

    Notes about the Halo API, learned the hard way:
      * count=N is capped at 1000 and returns the NEWEST 1000 while still
        reporting record_count=1000, so a naive completeness check passes while
        most of the history is missing. Page, never count. The open-only sweep
        is merged back in as a cross-check and the run stops if records went
        missing rather than writing a report that understates the position.
      * Project = tickettype_id 57 on the tenant this was written against,
        Project Task = 58. Set them in the CONFIGURATION block. Match on id,
        never on name - the display names carry stray whitespace.
      * Milestones ride on the PROJECT ticket as a milestones[] array with
        underscored start_date / target_date. A task points at one through
        milestone_id.
      * 1900-01-01 is Halo's "unset" date.
      * Deactivated agents are absent from /api/Agent, so their tickets come
        back as "id:NN" unless they are named in $FormerAgents below.
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

$OutFile         = ''      # blank = Documents\Halo-Portfolio-<timestamp>.html
$OpenWhenDone    = $true   # launch the report in the default browser
$IncludeClosed   = $true   # closed projects still set rolled-up windows, and a
                           # meeting usually wants to see what landed
$ProjectFilter   = '*'     # '*' = every project, or e.g. '*migration*'
$AgentFilter     = '*'     # '*' = whole team, or e.g. 'Jane Smith' to keep only
                           # projects that person manages
$AttentionLimit  = 40      # rows in the portfolio-wide "needs attention" table

# --- Agents who have left ----------------------------------------------------
# Deactivated agents vanish from /api/Agent, so their tickets come back as
# "id:NN". Name them here and they read properly everywhere.
# NOTE: bare numeric keys in a hashtable literal are INTEGERS. Look them up
# with an int, never with "$id", or the lookup silently never matches.

$FormerAgents = @{
    13 = 'Joe Bloggs'
}
$FormerSuffix  = ' (left)'   # set to '' to show the name with no marker

# --- How the RAG verdict is judged -------------------------------------------
# A project is not in trouble because a burn-up figure says so - a team can be
# 40% through the calendar with 10% of the work closed and be perfectly fine,
# because the big tasks are under way. It is in trouble when individual tasks
# are past their end date, or when a task is running out of its OWN window and
# nobody has picked it up. So the verdict is built from the tasks, and every
# reason is printed beside the colour.

$NearEndPct  = 75          # a task this far through its own window is closing in
$NearEndDays = 5           # ...or this close to its end date, whichever comes first
$DueSoonDays = 14          # horizon for the "what is coming" table

# Whether anybody is actually working a task is a question about its status,
# and status names are tenant-specific. These are matched with -like, which is
# case-insensitive. Anything not listed here and not closed is taken to be IN
# PROGRESS, so an unfamiliar status never invents a problem. The report prints
# the mapping it used and the console names anything it could not place - read
# it once, adjust these, and it stays right from then on.
$NotStartedStatuses = @('New','Logged','Raised','To Do','Backlog','Scheduled','Unassigned','*not started*')
$OnHoldStatuses     = @('*hold*','*waiting*','*with customer*','*with user*','*with supplier*',
                        '*with third party*','*pending*','*blocked*','*approval*','*deferred*')
$InProgressStatuses = @('*progress*','*started*','*underway*','*being *','*in hand*','*with agent*','*assigned*')

# A task past its end date that somebody is actively working: red, or amber?
# It defaults to red, because the date has gone whatever anyone is doing about
# it, and a status report that quietly swallows that is not worth reading.
# Set it to $false if your team runs to soft end dates and only wants red when
# a late task has nobody on it.
$LateInProgressIsRed = $true

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
        $s = $s -replace '(?s)<(script|style)\b.*?(</\1>|$)', ' '
        $s = $s -replace '(?s)<[^>]+>', ' '
        $s = [System.Net.WebUtility]::HtmlDecode($s)
        $s = $s -replace 'data:[^;,\s]*;base64,[A-Za-z0-9+/=]*', '[embedded data]'
        $s = $s -replace '(?s)\{[^{}]*\}', ' '
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
    Write-Progress -Id 0 -Activity 'Halo portfolio report' -Status "Step $Step of $Of - $Name" `
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
#  THE MODEL - one project at a time
# ---------------------------------------------------------------------------

function Test-StatusListed {
    # did the tenant's status name actually appear in one of the lists, or did
    # it fall through to the in-progress default? The report has to be able to
    # say which, or the mapping is unauditable.
    param([string]$s)
    foreach ($p in $NotStartedStatuses) { if ($s -like $p) { return $true } }
    foreach ($p in $OnHoldStatuses)     { if ($s -like $p) { return $true } }
    foreach ($p in $InProgressStatuses) { if ($s -like $p) { return $true } }
    return $false
}

function Get-TaskState {
    param([string]$StatusName, [bool]$IsClosed)
    if ($IsClosed) { return 'closed' }
    foreach ($p in $NotStartedStatuses) { if ($StatusName -like $p) { return 'new' } }
    foreach ($p in $OnHoldStatuses)     { if ($StatusName -like $p) { return 'hold' } }
    return 'progress'
}

function New-Row {
    param($t, [bool]$isProject)
    $sid  = Get-Val $t 'status_id'
    $name = Status-Name $sid
    $closedOn = Format-D (Get-Val $t 'dateclosed')
    $isClosed = ($name -match '^(Closed|Resolved|Cancelled|Completed)') -or ($closedOn -ne '')
    return @{
        state     = Get-TaskState $name $isClosed
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
    # Broken DATA only. Being late is a delivery problem, not a data problem,
    # and it is handled by the risk machinery below - flagging it here as well
    # only meant the same task was reported twice in two different voices.
    param($r)
    $f = @()
    if ($r.closed) { return $f }
    if (-not $r.start)  { $f += 'no start' }
    if (-not $r.target) { $f += 'no end' }
    if ($r.start -and $r.target -and $r.target -lt $r.start) { $f += 'end before start' }
    return $f
}

function Get-TaskRisk {
    # A task is judged against its OWN window, not the project's: three days
    # left on a three-day task is business as usual, three days left on a
    # forty-day task nobody has started is the thing you want to hear about.
    # The absolute-days test is the floor underneath that, so a short task
    # still gets a warning before it lands rather than after.
    param($k, [datetime]$NowD)
    $r = @{ level='ok'; why=''; daysLeft=$null; ownPct=-1 }
    if ($k.closed) { $r.level = 'done'; return $r }
    if (-not $k.target) {
        $r.level = 'unknown'
        $r.why   = 'no end date, so there is nothing to judge it against'
        return $r
    }
    $t = To-Date $k.target
    $r.daysLeft = [int](($t - $NowD).TotalDays)
    if ($k.hasWindow) {
        $s = To-Date $k.start
        $span = ($t - $s).TotalDays
        if ($span -gt 0) {
            $r.ownPct = [int][Math]::Round((($NowD - $s).TotalDays / $span) * 100)
            if ($r.ownPct -lt 0) { $r.ownPct = 0 }
        }
    }
    $word = 'in progress'
    if ($k.state -eq 'new')  { $word = 'not started' }
    if ($k.state -eq 'hold') { $word = 'on hold' }

    if ($r.daysLeft -lt 0) {
        $r.level = 'late'
        $r.why   = "$([Math]::Abs($r.daysLeft)) day(s) past its end date, $word"
        return $r
    }
    if (-not (($r.daysLeft -le $NearEndDays) -or ($r.ownPct -ge $NearEndPct))) { return $r }

    $left = "$($r.daysLeft) day(s) left"
    if ($k.hasWindow)      { $left = "$($r.daysLeft) of its $([int]$k.days) days left" }
    if ($r.daysLeft -eq 0) { $left = 'due today' }
    if ($k.state -eq 'progress') {
        # somebody is on it - worth a mention, not worth a colour
        $r.level = 'watch'
        $r.why   = "$left, in progress"
    } else {
        $r.level = 'atrisk'
        $r.why   = "$left, $word"
    }
    return $r
}

function Get-RiskBullets {
    param($Rows, [int]$Show = 3, [string]$More = 'more')
    $out = @(); $i = 0
    foreach ($r in @($Rows | Sort-Object @{Expression={$_.target}})) {
        $i++
        if ($i -gt $Show) { break }
        $out += ($r.summary + ' (#' + $r.id + ') - ' + $r.riskWhy)
    }
    if ($Rows.Count -gt $Show) { $out += ('and ' + ($Rows.Count - $Show) + ' ' + $More) }
    return $out
}

function Get-ProjectModel {
    # Everything the renderer needs for one project, worked out in one place
    # so the portfolio view and the project page can never disagree.
    param($Project, $TaskRecords)

    $today = (Get-Date).ToString('yyyy-MM-dd')
    $nowD  = (Get-Date).Date

    $P = New-Row $Project $true
    $P.flags = Get-Flags $P

    # milestones ride on the project ticket, with UNDERSCORED date fields
    $milestones = New-Object System.Collections.ArrayList
    foreach ($m in @(Get-Val $Project 'milestones')) {
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
    foreach ($t in $TaskRecords) {
        $r = New-Row $t $false
        $r.flags = Get-Flags $r
        [void]$tasks.Add($r)
    }

    # --- weighting ---------------------------------------------------------
    # Progress as a flat count of closed tasks flatters a project that has
    # ticked off a dozen half-day jobs and not started the three-week one.
    # Weight each task by its own length in days instead, and fall back to a
    # single day when a task has no usable window so it still counts.
    foreach ($k in $tasks) {
        $d = 1.0
        $a = To-Date $k.start
        $b = To-Date $k.target
        $k.hasWindow = ($null -ne $a -and $null -ne $b -and $b -ge $a)
        if ($k.hasWindow) { $d = ($b - $a).TotalDays + 1 }
        $k.days = [Math]::Max(1.0, $d)
    }

    foreach ($k in $tasks) {
        $r = Get-TaskRisk $k $nowD
        $k.risk     = $r.level
        $k.riskWhy  = $r.why
        $k.daysLeft = $r.daysLeft
        $k.ownPct   = $r.ownPct
    }

    $totTask  = $tasks.Count
    $closedT  = @($tasks | Where-Object { $_.closed })
    $openT    = @($tasks | Where-Object { -not $_.closed })
    $lateT    = @($tasks | Where-Object { $_.risk -eq 'late' })
    $atRiskT  = @($tasks | Where-Object { $_.risk -eq 'atrisk' })
    $watchT   = @($tasks | Where-Object { $_.risk -eq 'watch' })
    $unknownT = @($tasks | Where-Object { $_.risk -eq 'unknown' })
    $inProgT  = @($openT | Where-Object { $_.state -eq 'progress' })
    $onHoldT  = @($openT | Where-Object { $_.state -eq 'hold' })
    $newT     = @($openT | Where-Object { $_.state -eq 'new' })
    $dueSoonT = @($openT | Where-Object { $_.risk -eq 'ok' -and $null -ne $_.daysLeft -and $_.daysLeft -le $DueSoonDays })
    $unsched  = @($openT | Where-Object { -not $_.start -or -not $_.target })
    $noOwner  = @($openT | Where-Object { $_.agent -eq 'Unassigned' })

    # how much of the remaining work is simply invisible to the risk test
    $unknownDays = 0.0; $openDays = 0.0
    foreach ($k in $unknownT) { $unknownDays += $k.days }
    foreach ($k in $openT)    { $openDays    += $k.days }
    $unknownShare = 0
    if ($openDays -gt 0) { $unknownShare = [int][Math]::Round($unknownDays / $openDays * 100) }

    $sumDays = 0.0; $doneDays = 0.0
    foreach ($k in $tasks) { $sumDays += $k.days; if ($k.closed) { $doneDays += $k.days } }
    $pctCount = 0
    if ($totTask -gt 0) { $pctCount = [int][Math]::Round($closedT.Count / $totTask * 100) }
    $pctDays = 0
    if ($sumDays -gt 0) { $pctDays = [int][Math]::Round($doneDays / $sumDays * 100) }

    # --- the window --------------------------------------------------------
    # A project record's own dates are frequently stale or blank, so the report
    # leads with the window rolled up from the tasks - and says plainly when
    # the two disagree rather than quietly picking one.
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
    # @(...) matters: indexing a SINGLE-element array without it returns the
    # first CHARACTER of the string, which silently produces a garbage window.
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

    # --- the RAG verdict ---------------------------------------------------
    # One colour, and always the reasons beside it. The colour never carries
    # the meaning on its own - the word and the reasons do - so it survives
    # being printed in black and white, or read by someone who does not see red
    # and green apart.
    #
    #   red    something is already past its end date
    #   amber  something is running out of its window and nobody has picked it up
    #   green  everything open is either in hand, or not yet near its end date
    #
    # Data-quality gripes - no dates, nobody assigned - are listed separately
    # and deliberately do not move the colour.
    $rag = 'green'; $ragWord = 'Green'; $reasons = @(); $notes = @()

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
        $reasons += 'The project has no tasks, so there is nothing to judge'
        if ($effTarget -and $effTarget -lt $today) {
            $rag = 'red'; $ragWord = 'Red'
            $reasons += "The project record is past its end date ($(Show-D $effTarget))"
        }
    }
    elseif ($openT.Count -eq 0) {
        $rag = 'green'; $ragWord = 'Green'
        $reasons += 'Every task is closed - the project record itself has not been closed yet'
    }
    else {
        if ($lateT.Count -gt 0) {
            $lateUnworked = @($lateT | Where-Object { $_.state -ne 'progress' })
            if ($LateInProgressIsRed -or $lateUnworked.Count -gt 0) { $rag = 'red';   $ragWord = 'Red' }
            else                                                    { $rag = 'amber'; $ragWord = 'Amber' }
            $reasons += (Get-RiskBullets $lateT 3 'more past their end date')
        }
        elseif ($effTarget -and $effTarget -lt $today) {
            # only reachable when the open work carries no dates of its own
            $rag = 'red'; $ragWord = 'Red'
            $reasons += "The project is past its end date ($(Show-D $effTarget)) with $($openT.Count) task(s) still open"
        }
        if ($atRiskT.Count -gt 0) {
            if ($rag -ne 'red') { $rag = 'amber'; $ragWord = 'Amber' }
            $reasons += (Get-RiskBullets $atRiskT 3 'more running out of time')
        }
        if ($unknownT.Count -gt 0 -and $unknownShare -ge 50) {
            if ($rag -eq 'green') { $rag = 'amber'; $ragWord = 'Amber' }
            $reasons += ("$($unknownT.Count) open task(s) have no end date and account for $unknownShare% of " +
                         'the work left, so there is not enough in Halo to call this one')
        }
        if ($rag -eq 'green') {
            $bits = @()
            if ($inProgT.Count -gt 0) { $bits += "$($inProgT.Count) in progress" }
            if ($newT.Count -gt 0)    { $bits += "$($newT.Count) not started but not yet near its end date" }
            $reasons += ("Nothing is late and nothing is running out of time" +
                         $(if ($bits.Count -gt 0) { ' - ' + ($bits -join ', ') } else { '' }))
            if ($effTarget -and $leftDays -ge 0) { $reasons += "$leftDays day(s) left to $(Show-D $effTarget)" }
        }
        if ($watchT.Count -gt 0) {
            $reasons += ("$($watchT.Count) task(s) are close to their end date but in progress" +
                         ' - worth a word at the next catch-up rather than an escalation')
        }
    }

    # housekeeping - said out loud, kept out of the verdict
    if ($unsched.Count -gt 0) { $notes += "$($unsched.Count) open task(s) have no start or no end date" }
    if ($noOwner.Count -gt 0) { $notes += "$($noOwner.Count) open task(s) have nobody assigned" }
    if ($onHoldT.Count -gt 0) { $notes += "$($onHoldT.Count) open task(s) are on hold or waiting on somebody" }

    # --- how the statuses were read ----------------------------------------
    $stateOfStatus = @{}
    foreach ($k in $tasks) { if (-not $stateOfStatus.ContainsKey($k.status)) { $stateOfStatus[$k.status] = $k.state } }
    $unlisted = @()
    foreach ($s in @($stateOfStatus.Keys | Sort-Object)) {
        if ($stateOfStatus[$s] -eq 'progress' -and -not (Test-StatusListed $s)) { $unlisted += $s }
    }

    # --- per-agent ---------------------------------------------------------
    $byAgent = @{}
    foreach ($a in @($P.agent) + @($tasks | ForEach-Object { $_.agent })) {
        if (-not $byAgent.ContainsKey($a)) {
            $byAgent[$a] = @{ open=0; closed=0; late=0; atrisk=0; days=0.0; doneDays=0.0; next='' }
        }
    }
    foreach ($k in $tasks) {
        $e = $byAgent[$k.agent]
        $e.days += $k.days
        if ($k.closed) { $e.closed++; $e.doneDays += $k.days }
        else {
            $e.open++
            if ($k.risk -eq 'late')   { $e.late++ }
            if ($k.risk -eq 'atrisk') { $e.atrisk++ }
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

    # --- milestone roll-up -------------------------------------------------
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
        $m.late   = @($kids | Where-Object { $_.risk -eq 'late' }).Count
        $m.atRisk = @($kids | Where-Object { $_.risk -eq 'atrisk' }).Count
    }
    $unMs = @($tasks | Where-Object { $_.milestone -le 0 -or -not $msName.ContainsKey($_.milestone) })

    # --- this project's own timeline scale ---------------------------------
    $sd = @($P.start, $P.target, $effStart, $effTarget, $today)
    foreach ($k in $tasks)      { $sd += $k.start; $sd += $k.target; $sd += $k.closedOn }
    foreach ($m in $milestones) { $sd += $m.start; $sd += $m.target }
    $scale = Get-Scale $sd
    $tp = Today-In $scale
    $todayMark = ''
    if ($tp -ge 0) { $todayMark = '<u style="left:' + ('{0:F2}' -f $tp) + '%"></u>' }

    # the single most useful date for a portfolio row: the next open end date
    $nextDue = ''
    foreach ($k in @($openT | Where-Object { $_.target } | Sort-Object @{Expression={$_.target}})) {
        $nextDue = $k.target; break
    }

    return @{
        P = $P; tasks = $tasks; milestones = $milestones
        rag = $rag; ragWord = $ragWord; reasons = $reasons; notes = $notes
        effStart = $effStart; effTarget = $effTarget
        elapsed = $elapsed; spanDays = $spanDays; goneDays = $goneDays; leftDays = $leftDays
        totTask = $totTask; closedT = $closedT; openT = $openT
        lateT = $lateT; atRiskT = $atRiskT; watchT = $watchT; unknownT = $unknownT
        inProgT = $inProgT; onHoldT = $onHoldT; newT = $newT; dueSoonT = $dueSoonT
        unsched = $unsched; noOwner = $noOwner
        pctCount = $pctCount; pctDays = $pctDays; sumDays = $sumDays; doneDays = $doneDays
        unknownShare = $unknownShare
        byAgent = $byAgent; agentOrder = $agentOrder; slotOf = $slotOf
        msName = $msName; unMs = $unMs
        stateOfStatus = $stateOfStatus; unlisted = $unlisted
        scale = $scale; todayMark = $todayMark; today = $today
        nextDue = $nextDue
    }
}

# ---------------------------------------------------------------------------
#  TIMELINE SCALES
# ---------------------------------------------------------------------------

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
    # today's position on a scale, or -1 when today is off the end
    param($Scale)
    if ($null -eq $Scale) { return -1 }
    $t = (Get-Date).Date
    if ($t -lt $Scale.a -or $t -gt $Scale.b) { return -1 }
    return ($t - $Scale.a).TotalDays / $Scale.span * 100
}

function Axis-Row {
    param($Scale, [int]$Before, [int]$After = 0)
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
    $out += '</td>'
    if ($After -gt 0) { $out += '<td colspan="' + $After + '" style="border-bottom:1px solid var(--grid);padding:0"></td>' }
    return ($out + '</tr>')
}

function Get-RiskBadge {
    param([string]$level)
    switch ($level) {
        'late'   { return '<span class="flag f-bad"  style="margin-left:0">late</span>' }
        'atrisk' { return '<span class="flag f-warn" style="margin-left:0">at risk</span>' }
        'watch'  { return '<span class="flag f-mut"  style="margin-left:0">closing in</span>' }
        'ok'     { return '<span class="flag f-mut"  style="margin-left:0">due soon</span>' }
        default  { return '' }
    }
}

function Get-RagBlurb {
    param([string]$rag)
    if ($rag -eq 'red')   { return 'Needs intervention' }
    if ($rag -eq 'amber') { return 'Watch this one' }
    if ($rag -eq 'done')  { return 'Delivered' }
    return 'On track'
}

# ---------------------------------------------------------------------------
#  THE HTML
# ---------------------------------------------------------------------------

$script:CSS = @'
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
a{color:var(--s1)}
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
.ragnote{margin:13px 0 0;padding-top:10px;border-top:1px solid var(--ring);
 font-size:12.5px;color:var(--text-secondary)}
.ragnote b{font-weight:600}
/* the portfolio row's colour chip - same palette, small enough for a table */
.chip{display:inline-block;font-size:10px;font-weight:700;letter-spacing:.6px;
 text-transform:uppercase;padding:3px 9px;border-radius:99px;color:#fff;white-space:nowrap}
.chip.red{background:var(--r-red)}
.chip.amber{background:var(--r-amber)}
.chip.green{background:var(--r-green)}
.chip.done{background:var(--r-done)}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(132px,1fr));gap:11px;margin-bottom:16px}
.tile{background:var(--surface-1);border:1px solid var(--ring);border-radius:11px;padding:13px 15px}
.tile b{display:block;font-size:25px;line-height:1.05;letter-spacing:-.02em;font-weight:600}
.tile span{display:block;color:var(--text-secondary);font-size:12px;margin-top:4px}
.tile.warn b{color:var(--critical)}
.tile.red b{color:var(--r-red)}
.tile.amber b{color:var(--r-amber)}
.tile.green b{color:var(--r-green)}
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
/* a portfolio row is a link to its own page - make that obvious */
tr.prow{cursor:pointer}
tr.prow:hover td{background:var(--plane)}
tr.prow td:first-child{white-space:nowrap}
.pname{font-weight:600;text-decoration:none;color:var(--text-primary)}
.pname:hover{text-decoration:underline}
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
.meter.red b{background:var(--r-red)}
.meter.amber b{background:var(--r-amber)}
.meter.green b{background:var(--r-green)}
.pct{font-size:12px;font-variant-numeric:tabular-nums;color:var(--text-secondary);margin-left:7px}
.flag{font-size:10px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;
 padding:2px 6px;border-radius:99px;border:1px solid;white-space:nowrap;margin-left:5px;
 display:inline-block}
.f-bad{color:var(--critical);border-color:var(--critical)}
.f-warn{color:var(--r-amber);border-color:var(--r-amber)}
.f-mut{color:var(--text-secondary);border-color:var(--axis)}
.f-ok{color:var(--good);border-color:var(--good)}
.why{color:var(--text-secondary);font-size:12.5px}
.desc{font-size:13.5px;color:var(--text-secondary);max-width:92ch;white-space:pre-wrap}
/* the task table is genuinely wide - let it scroll inside its own card rather
   than dragging the whole page sideways on a laptop or a phone */
.scrollx{overflow-x:auto}
.scrollx table{min-width:840px}
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
/* one file, several pages: only the current view is on screen */
.view{display:none}
.view.active{display:block}
.back{display:inline-block;margin:0 0 14px;font-size:13px;text-decoration:none;color:var(--s1)}
.back:hover{text-decoration:underline}
.crumb{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:6px}
@media print{
 .bar,.top,.back,.noprint{display:none}
 body{background:#fff}
 tr{break-inside:avoid}
 .rag{border-width:2px}
 .scrollx{overflow:visible}
 .scrollx table{min-width:0}
 /* printing the pack prints every project, each starting on its own page */
 body[data-print="all"] .view{display:block!important}
 body[data-print="all"] .view.pv{break-before:page;page-break-before:always}
}
@media (max-width:760px){.kv{gap:6px 16px}}
</style>
'@

$script:JS = @'
<script>
(function(){
 // ---- one file, several pages ---------------------------------------------
 var views=[].slice.call(document.querySelectorAll('.view'));
 function show(id){
  var found=false;
  views.forEach(function(v){
   var on=(v.id===id);
   v.classList.toggle('active',on);
   if(on) found=true;
  });
  if(!found){ document.getElementById('overview').classList.add('active'); }
  window.scrollTo({top:0});
 }
 function fromHash(){
  var h=(location.hash||'').replace(/^#/,'');
  show(h||'overview');
 }
 window.addEventListener('hashchange',fromHash);

 // a whole row is clickable, but the name is a real link so the keyboard and
 // "open in new tab" both still work
 document.querySelectorAll('tr.prow').forEach(function(r){
  r.addEventListener('click',function(e){
   if(e.target.closest('a')) return;
   location.hash='#'+r.dataset.go;
  });
 });

 // ---- the portfolio filters -----------------------------------------------
 var oq=document.getElementById('oq'), orag=document.getElementById('orag'),
     oag=document.getElementById('oagent'), ocl=document.getElementById('oclosed'),
     ocount=document.getElementById('ocount');
 var prows=[].slice.call(document.querySelectorAll('tr.prow'));
 function oapply(){
  if(!oq) return;
  var term=oq.value.trim().toLowerCase(), want=orag.value, who=oag.value,
      showDone=ocl.checked, shown=0;
  prows.forEach(function(r){
   var ok=true, rag=r.dataset.rag;
   if(!showDone && rag==='done') ok=false;
   if(want==='attn'){ if(!(rag==='red'||rag==='amber')) ok=false; }
   else if(want && rag!==want) ok=false;
   if(who && r.dataset.agent!==who) ok=false;
   if(term && r.dataset.q.indexOf(term)<0) ok=false;
   r.classList.toggle('hidden',!ok); if(ok) shown++;
  });
  ocount.textContent=shown+' of '+prows.length+' projects shown';
 }
 if(oq){
  [orag,oag,ocl].forEach(function(e){e.addEventListener('change',oapply)});
  oq.addEventListener('input',oapply);
 }

 // ---- the task filters, one set per project page --------------------------
 var ATTN={late:1,atrisk:1,watch:1,unknown:1};
 document.querySelectorAll('section.pv').forEach(function(sec){
  var q=sec.querySelector('.f-q'), ag=sec.querySelector('.f-agent'),
      sc=sec.querySelector('.f-closed'), rk=sec.querySelector('.f-risk'),
      cnt=sec.querySelector('.f-count');
  if(!q) return;
  var rows=[].slice.call(sec.querySelectorAll('tr.task'));
  function apply(){
   var term=q.value.trim().toLowerCase(), a=ag.value,
       showClosed=sc.checked, want=rk.value, shown=0;
   rows.forEach(function(r){
    var ok=true, risk=r.dataset.risk;
    if(!showClosed && r.dataset.closed==='1') ok=false;
    if(a && r.dataset.agent!==a) ok=false;
    if(want==='prob' && r.dataset.prob!=='1') ok=false;
    if(want==='late' && risk!=='late') ok=false;
    if(want==='risk' && !(risk==='late'||risk==='atrisk')) ok=false;
    if(want==='attn' && !ATTN[risk]) ok=false;
    if(term && r.dataset.q.indexOf(term)<0) ok=false;
    r.classList.toggle('hidden',!ok); if(ok) shown++;
   });
   cnt.textContent=shown+' of '+rows.length+' tasks shown';
   // a milestone heading with nothing under it is noise - it goes with its rows
   sec.querySelectorAll('tr.grp').forEach(function(g){
    var any=false,n=g.nextElementSibling;
    while(n && !n.classList.contains('grp')){
     if(n.classList.contains('task') && !n.classList.contains('hidden')) any=true;
     n=n.nextElementSibling;
    }
    g.classList.toggle('hidden',!any);
   });
  }
  [sc,rk,ag].forEach(function(e){e.addEventListener('change',apply)});
  q.addEventListener('input',apply);
  sec._resetFilters=function(){ sc.checked=true; rk.value=''; ag.value=''; q.value=''; apply(); };
  apply();
 });

 // ---- printing ------------------------------------------------------------
 // Printing a filtered table prints a half-truth, so every print resets the
 // filters first. "This project" prints the page you are on; "the pack"
 // prints the portfolio and every project behind it.
 function resetAll(){
  document.querySelectorAll('section.pv').forEach(function(s){ if(s._resetFilters) s._resetFilters(); });
  if(oq){ oq.value=''; orag.value=''; oag.value=''; ocl.checked=true; oapply(); }
 }
 document.querySelectorAll('.printone').forEach(function(b){
  b.addEventListener('click',function(){
   document.body.dataset.print='one'; resetAll(); window.print();
  });
 });
 document.querySelectorAll('.printall').forEach(function(b){
  b.addEventListener('click',function(){
   document.body.dataset.print='all'; resetAll(); window.print();
  });
 });
 window.addEventListener('afterprint',function(){ document.body.dataset.print='one'; });

 var tt=document.getElementById('totop');
 if(tt) tt.addEventListener('click',function(){ window.scrollTo({top:0,behavior:'smooth'}); });

 fromHash();
})();
</script>
'@

function Write-ProjectView {
    # One project's page - the single-project status report, in a <section>
    # that the portfolio view links to.
    param($M, $SB)

    function W { param([string]$s) [void]$SB.AppendLine($s) }

    $P          = $M.P
    $tasks      = $M.tasks
    $milestones = $M.milestones
    $rag        = $M.rag
    $scale      = $M.scale
    $todayMark  = $M.todayMark
    $byAgent    = $M.byAgent
    $agentOrder = $M.agentOrder
    $slotOf     = $M.slotOf
    $effStart   = $M.effStart
    $effTarget  = $M.effTarget

    W ('<section class="view pv" id="p' + $P.id + '">')
    W '<div class="crumb"><a class="back" href="#overview">&larr; Back to the portfolio</a>'
    W '<button class="printone noprint">Print this project</button></div>'

    # --- heading -----------------------------------------------------------
    W ('<h1>' + (Esc $P.summary) + '</h1>')
    W ('<p class="lede"><span class="pid">Project #' + $P.id + '</span> &middot; ' +
       (Esc $M.ragWord) + ' &middot; ' + $M.totTask + ' task(s)</p>')

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

    # --- RAG ---------------------------------------------------------------
    W ('<div class="rag ' + $rag + '">')
    W ('<div class="ragtop"><span class="ragbadge">' + (Esc $M.ragWord) + '</span>' +
       '<span class="ragline">' + (Esc (Get-RagBlurb $rag)) + '</span></div>')
    W '<ul>'
    foreach ($r in $M.reasons) { W ('<li>' + (Esc $r) + '</li>') }
    W '</ul>'
    if ($M.notes.Count -gt 0) {
        W ('<p class="ragnote"><b>Housekeeping, and deliberately not part of the verdict:</b> ' +
           (Esc ($M.notes -join '; ')) + '.</p>')
    }
    W '</div>'

    # --- progress ----------------------------------------------------------
    W '<div class="card"><h3>Progress</h3>'
    if ($M.totTask -eq 0) {
        # an empty bar reading 0% would look like a project that has started
        # and achieved nothing, which is a different thing entirely from one
        # nobody has broken down yet
        W ('<p class="sub" style="margin:4px 0 0">There are no tasks under this project, so there is no ' +
           'completed work to measure. Until the plan is broken down in Halo the only thing this report ' +
           'can say about progress is that nothing is being tracked.</p>')
    } else {
        $workCls = 'big work'
        if ($rag -eq 'red')   { $workCls += ' red' }
        if ($rag -eq 'amber') { $workCls += ' amber' }
        W '<div class="pline">'
        W ('<div class="plab"><span>Work completed &mdash; closed tasks, weighted by how long each one runs</span><b>' +
           $M.pctDays + '%</b></div>')
        W ('<div class="' + $workCls + '"><b style="width:' + $M.pctDays + '%"></b></div>')
        W ('<div class="sub" style="margin-top:6px">' + $M.closedT.Count + ' of ' + $M.totTask + ' tasks closed (' +
           $M.pctCount + '% by simple count) &middot; ' + [int][Math]::Round($M.doneDays) + ' of ' +
           [int][Math]::Round($M.sumDays) + ' task-days done</div>')
        W '</div>'
    }

    if ($M.elapsed -ge 0) {
        W '<div class="pline">'
        W ('<div class="plab"><span>Schedule elapsed &mdash; ' + (Show-D $effStart) + ' to ' + (Show-D $effTarget) +
           '</span><b>' + $M.elapsed + '%</b></div>')
        W ('<div class="big time"><b style="width:' + $M.elapsed + '%"></b></div>')
        W ('<div class="marker"><u style="left:' + $M.elapsed + '%"></u><i style="left:' + $M.elapsed + '%">today</i></div>')
        W '</div>'
        $gap = $M.elapsed - $M.pctDays
        $verdict = 'level with the schedule'
        if ($gap -gt 0)     { $verdict = "$gap points behind the schedule" }
        elseif ($gap -lt 0) { $verdict = "$([Math]::Abs($gap)) points ahead of the schedule" }
        $leftTxt = "$($M.leftDays) day(s) left"
        if ($M.leftDays -lt 0) { $leftTxt = "$([Math]::Abs($M.leftDays)) day(s) past the end date" }
        W ('<p class="sub" style="margin:14px 0 0">A ' + $M.spanDays + ' day project: ' + $M.goneDays + ' gone, ' +
           $leftTxt + '. The work is ' + (Esc $verdict) + ', which is context rather than a verdict &mdash; ' +
           'a project loaded with long tasks that are all under way will read behind this line and still land.</p>')
    } else {
        W ('<p class="sub" style="margin:14px 0 0">Neither this project nor its tasks carry a usable ' +
           'start-and-end window, so there is no schedule to measure the work against.</p>')
    }
    W '</div>'

    # --- tiles -------------------------------------------------------------
    W '<div class="tiles">'
    $tiles = @(
        @{v=$M.totTask;        l='Tasks';          w=$false},
        @{v=$M.closedT.Count;  l='Closed';         w=$false},
        @{v=$M.inProgT.Count;  l='In progress';    w=$false},
        @{v=$M.newT.Count;     l='Not started';    w=$false},
        @{v=$M.onHoldT.Count;  l='On hold';        w=$false},
        @{v=$M.lateT.Count;    l='Late';           w=($M.lateT.Count -gt 0)},
        @{v=$M.atRiskT.Count;  l='At risk';        w=($M.atRiskT.Count -gt 0)},
        @{v=$M.watchT.Count;   l='Closing in';     w=$false},
        @{v=$M.unknownT.Count; l='No end date';    w=$false}
    )
    foreach ($t in $tiles) {
        W ('<div class="tile' + $(if ($t.w) { ' warn' } else { '' }) + '"><b>' + $t.v + '</b><span>' +
           (Esc $t.l) + '</span></div>')
    }
    W '</div>'

    # --- description -------------------------------------------------------
    if ($P.details) {
        $dtl = $P.details
        if ($dtl.Length -gt 2500) { $dtl = $dtl.Substring(0,2500) + "`n..." }
        W '<h2>Description</h2><div class="card"><div class="desc">'
        W (Esc $dtl)
        W '</div></div>'
    }

    # --- the people --------------------------------------------------------
    W '<h2>Who is assigned</h2><div class="card">'
    W '<table><thead><tr><th>Person</th><th class="n" style="width:62px">Open</th>'
    W '<th class="n" style="width:70px">Closed</th><th class="n" style="width:66px">Late</th>'
    W '<th class="n" style="width:76px">At risk</th>'
    W '<th style="width:174px">Their share done</th><th style="width:112px">Next due</th></tr></thead><tbody>'
    foreach ($a in $agentOrder) {
        $e = $byAgent[$a]
        if (($e.open + $e.closed) -eq 0 -and $a -ne $P.agent) { continue }
        $ap = 0
        if ($e.days -gt 0) { $ap = [int][Math]::Round($e.doneDays / $e.days * 100) }
        $who = (Esc $a)
        if ($a -eq $P.agent) { $who += '<span class="flag f-mut">project manager</span>' }
        W ('<tr><td><span class="dot" style="background:var(--s' + $slotOf[$a] + ')"></span>' + $who + '</td>')
        W ('<td class="n">' + $e.open + '</td><td class="n">' + $e.closed + '</td>')
        W ('<td class="n">' + $(if ($e.late   -gt 0) { '<span class="flag f-bad" style="margin-left:0">'  + $e.late   + '</span>' } else { '0' }) + '</td>')
        W ('<td class="n">' + $(if ($e.atrisk -gt 0) { '<span class="flag f-warn" style="margin-left:0">' + $e.atrisk + '</span>' } else { '0' }) + '</td>')
        if (($e.open + $e.closed) -gt 0) {
            W ('<td><span class="meter"><b style="width:' + $ap + '%"></b></span><span class="pct">' + $ap + '%</span></td>')
        } else {
            W '<td class="d">no tasks of their own</td>'
        }
        W ('<td class="d">' + $(if ($e.next) { Show-D $e.next } else { '&mdash;' }) + '</td></tr>')
    }
    W '</tbody></table></div>'

    # --- milestones --------------------------------------------------------
    if ($milestones.Count -gt 0) {
        W '<h2>Milestones</h2><div class="card"><table>'
        W '<thead><tr><th>Milestone</th><th class="n" style="width:88px">Tasks</th>'
        W '<th style="width:148px">Done</th><th class="d" style="width:166px">Window</th>'
        W '<th style="width:28%">Timeline</th></tr></thead><tbody>'
        W (Axis-Row $scale 4)
        foreach ($m in @($milestones | Sort-Object @{Expression={$_.seq}}, @{Expression={$_.start}})) {
            W '<tr>'
            $tag = ''
            if ($m.late   -gt 0) { $tag += '<span class="flag f-bad">'  + $m.late   + ' late</span>' }
            if ($m.atRisk -gt 0) { $tag += '<span class="flag f-warn">' + $m.atRisk + ' at risk</span>' }
            if (-not $tag -and $m.taskCount -gt 0 -and $m.doneCount -eq $m.taskCount) {
                $tag = '<span class="flag f-ok">complete</span>'
            }
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

    # --- needs attention ---------------------------------------------------
    # Ordered by how much it matters, not by date: what is already late, then
    # what is running out of time unattended, then what is close but in hand,
    # then what is merely coming up.
    $attn = @()
    $attn += @($M.lateT    | Sort-Object @{Expression={$_.target}})
    $attn += @($M.atRiskT  | Sort-Object @{Expression={$_.target}})
    $attn += @($M.watchT   | Sort-Object @{Expression={$_.target}})
    $attn += @($M.dueSoonT | Sort-Object @{Expression={$_.target}})
    if ($attn.Count -gt 0) {
        W ('<h2>What needs attention</h2><div class="card">')
        W ('<p class="sub" style="margin:0 0 12px">Each task judged against its own window, so a long ' +
           'task with a little time left ranks above a short one with the same days on the clock.</p>')
        W '<table><thead><tr><th style="width:86px">Risk</th><th style="width:54px">ID</th><th>Task</th>'
        W '<th style="width:32%">Why</th><th style="width:124px">Owner</th>'
        W '<th class="d" style="width:94px">End</th></tr></thead><tbody>'
        foreach ($k in $attn) {
            $why = $k.riskWhy
            if (-not $why) {
                $why = "$($k.daysLeft) day(s) to its end date"
                if ($k.daysLeft -eq 0) { $why = 'due today' }
                $sw = 'in progress'
                if ($k.state -eq 'new')  { $sw = 'not started' }
                if ($k.state -eq 'hold') { $sw = 'on hold' }
                $why += ", $sw"
            }
            W ('<tr><td>' + (Get-RiskBadge $k.risk) + '</td><td class="mono">' + $k.id + '</td>' +
               '<td>' + (Esc $k.summary) + '</td><td class="why">' + (Esc $why) + '</td>' +
               '<td class="d">' + (Esc $k.agent) + '</td><td class="d">' + (Show-D $k.target) + '</td></tr>')
        }
        W '</tbody></table></div>'
    }

    # --- date problems -----------------------------------------------------
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

    # --- every task --------------------------------------------------------
    W '<h2>Tasks</h2>'
    W '<div class="bar">'
    W '<label><input type="checkbox" class="f-closed" checked> Closed tasks</label>'
    W '<label>Show <select class="f-risk">'
    W '<option value="">every task</option>'
    W '<option value="late">late only</option>'
    W '<option value="risk">late or at risk</option>'
    W '<option value="attn">anything needing attention</option>'
    W '<option value="prob">date problems</option>'
    W '</select></label>'
    W '<label>Owner <select class="f-agent"><option value="">everyone</option>'
    foreach ($a in $agentOrder) {
        if (($byAgent[$a].open + $byAgent[$a].closed) -eq 0) { continue }
        W ('<option>' + (Esc $a) + '</option>')
    }
    W '</select></label>'
    W '<input type="search" class="f-q" placeholder="Search tasks&hellip;" style="min-width:200px">'
    W '<span class="f-count pid" style="margin-left:auto"></span>'
    W '</div>'

    W '<div class="card scrollx" style="padding:12px 16px">'
    if ($M.totTask -eq 0) {
        W '<p style="margin:0" class="sub">This project has no tasks attached to it in Halo, so there is nothing to list.</p>'
    } else {
        W '<table><thead><tr><th style="width:54px">ID</th><th>Task</th>'
        W '<th style="width:116px">Owner</th><th style="width:116px">Status</th>'
        W '<th style="width:90px">Start</th><th style="width:90px">End</th>'
        W '<th class="n" style="width:54px">Days</th><th style="width:94px">Closed</th>'
        W '<th style="width:22%">Timeline</th></tr></thead><tbody>'
        W (Axis-Row $scale 8)

        # grouped by milestone, in milestone order, with the unassigned ones
        # last - that is how the plan reads in Halo, so it is how it reads here
        $groups = New-Object System.Collections.ArrayList
        foreach ($m in @($milestones | Sort-Object @{Expression={$_.seq}}, @{Expression={$_.start}})) {
            $kids = @($tasks | Where-Object { $_.milestone -eq $m.id })
            if ($kids.Count -gt 0) { [void]$groups.Add(@{ name=$m.name; rows=$kids }) }
        }
        if ($M.unMs.Count -gt 0) {
            $nm = ''
            if ($milestones.Count -gt 0) { $nm = 'No milestone' }
            [void]$groups.Add(@{ name=$nm; rows=@($M.unMs) })
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
                   '" data-risk="' + $k.risk +
                   '" data-agent="' + (Esc $k.agent) +
                   '" data-q="' + (Esc (($k.summary + ' ' + $k.agent + ' ' + $k.status + ' ' + $k.id).ToLower())) + '">')
                W ('<td class="mono">' + $k.id + '</td>')
                W ('<td>' + (Esc $k.summary))
                # the badge carries its own reason in the tooltip, so the row
                # stays narrow but the detail is a hover away
                if ($k.risk -eq 'late' -or $k.risk -eq 'atrisk' -or $k.risk -eq 'watch') {
                    $b = Get-RiskBadge $k.risk
                    $b = $b -replace 'style="margin-left:0"', ('title="' + (Esc $k.riskWhy) + '"')
                    W $b
                }
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

    W '<p class="sub" style="margin-top:18px"><a class="back" href="#overview">&larr; Back to the portfolio</a></p>'
    W '</section>'
}

function New-ReportHtml {
    param($Models, [string]$Stamp, [hashtable]$Stats)

    $sb = New-Object System.Text.StringBuilder
    function W { param([string]$s) [void]$sb.AppendLine($s) }

    W '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
    W '<meta name="viewport" content="width=device-width, initial-scale=1">'
    W ('<title>' + (Esc ($Tenant + ' project portfolio ' + $Stamp)) + '</title>')
    W $script:CSS
    W '</head><body class="viz-root" data-print="one"><div class="wrap">'

    # =======================================================================
    #  THE PORTFOLIO VIEW
    # =======================================================================
    W '<section class="view active" id="overview">'
    W '<h1>Project portfolio</h1>'
    W ('<p class="lede">' + (Esc $Tenant) + ' &middot; ' + $Models.Count + ' project(s) &middot; as at ' +
       (Esc $Stamp) + ' &middot; read-only snapshot, nothing was changed in Halo. ' +
       'Click a project for its full status page.</p>')

    W '<div class="tiles">'
    $otiles = @(
        @{v=$Models.Count;      l='Projects';        c=''},
        @{v=$Stats.red;         l='Red';             c=$(if ($Stats.red   -gt 0) { ' red' }   else { '' })},
        @{v=$Stats.amber;       l='Amber';           c=$(if ($Stats.amber -gt 0) { ' amber' } else { '' })},
        @{v=$Stats.green;       l='Green';           c=' green'},
        @{v=$Stats.done;        l='Complete';        c=''},
        @{v=$Stats.tasks;       l='Tasks';           c=''},
        @{v=$Stats.lateTasks;   l='Late tasks';      c=$(if ($Stats.lateTasks   -gt 0) { ' warn' } else { '' })},
        @{v=$Stats.atRiskTasks; l='At risk';         c=$(if ($Stats.atRiskTasks -gt 0) { ' warn' } else { '' })},
        @{v=$Stats.noTasks;     l='No tasks yet';    c=''}
    )
    foreach ($t in $otiles) {
        W ('<div class="tile' + $t.c + '"><b>' + $t.v + '</b><span>' + (Esc $t.l) + '</span></div>')
    }
    W '</div>'

    # --- filter bar --------------------------------------------------------
    W '<div class="bar">'
    W '<label>Show <select id="orag">'
    W '<option value="">every project</option>'
    W '<option value="attn">anything red or amber</option>'
    W '<option value="red">red only</option>'
    W '<option value="amber">amber only</option>'
    W '<option value="green">green only</option>'
    W '</select></label>'
    W '<label>Manager <select id="oagent"><option value="">everyone</option>'
    foreach ($a in @($Models | ForEach-Object { $_.P.agent } | Sort-Object -Unique)) {
        W ('<option>' + (Esc $a) + '</option>')
    }
    W '</select></label>'
    W '<label><input type="checkbox" id="oclosed" checked> Completed projects</label>'
    W '<input type="search" id="oq" placeholder="Search projects&hellip;" style="min-width:200px">'
    W '<button class="printall">Print the pack</button>'
    W '<span id="ocount" class="pid" style="margin-left:auto"></span>'
    W '</div>'

    # --- the portfolio table ------------------------------------------------
    # One row per project: the verdict, how far through the work and the
    # calendar it is, and where it sits against every other project on one
    # shared timeline.
    $allDates = @()
    foreach ($m in $Models) {
        if ($m.effStart)  { $allDates += $m.effStart }
        if ($m.effTarget) { $allDates += $m.effTarget }
    }
    $allDates += (Get-Date).ToString('yyyy-MM-dd')
    $pScale = Get-Scale $allDates
    $pt = Today-In $pScale
    $pMark = ''
    if ($pt -ge 0) { $pMark = '<u style="left:' + ('{0:F2}' -f $pt) + '%"></u>' }

    W '<div class="card scrollx" style="padding:12px 16px">'
    W '<table><thead><tr><th style="width:92px">Verdict</th><th>Project</th>'
    W '<th style="width:128px">Manager</th><th class="d" style="width:168px">Window</th>'
    W '<th style="width:150px">Work done</th><th class="n" style="width:78px">Elapsed</th>'
    W '<th class="n" style="width:62px">Late</th><th class="n" style="width:74px">At risk</th>'
    W '<th class="d" style="width:96px">Next due</th><th style="width:20%">Timeline</th></tr></thead><tbody>'
    W (Axis-Row $pScale 9)

    # worst first: that is the order a project meeting works through them
    $ragOrder = @{ red=0; amber=1; green=2; done=3 }
    $ordered = @($Models | Sort-Object `
        @{Expression={$ragOrder[$_.rag]}}, `
        @{Expression={if ($_.effTarget) { $_.effTarget } else { '9999-99-99' }}}, `
        @{Expression={$_.P.summary}})

    foreach ($m in $ordered) {
        $P = $m.P
        $q = ($P.summary + ' ' + $P.agent + ' ' + $P.status + ' ' + $P.id + ' ' + $P.client + ' ' + $P.site).ToLower()
        W ('<tr class="prow" data-go="p' + $P.id + '" data-rag="' + $m.rag +
           '" data-agent="' + (Esc $P.agent) + '" data-q="' + (Esc $q) + '">')
        W ('<td><span class="chip ' + $m.rag + '">' + (Esc $m.ragWord) + '</span></td>')
        W ('<td><a class="pname" href="#p' + $P.id + '">' + (Esc $P.summary) + '</a>' +
           '<div class="sub"><span class="pid">#' + $P.id + '</span> &middot; ' + $m.totTask + ' task(s), ' +
           $m.closedT.Count + ' closed' +
           $(if ($m.totTask -eq 0) { ' <span class="flag f-warn">no tasks</span>' } else { '' }) + '</div></td>')
        W ('<td class="d">' + (Esc $P.agent) + '</td>')
        W ('<td class="d">' + $(if ($m.effStart -or $m.effTarget) { (Show-D $m.effStart) + ' &rarr; ' + (Show-D $m.effTarget) } else { '&mdash;' }) + '</td>')
        $mcls = 'meter'
        if ($m.rag -eq 'red')   { $mcls += ' red' }
        if ($m.rag -eq 'amber') { $mcls += ' amber' }
        if ($m.rag -eq 'green' -or $m.rag -eq 'done') { $mcls += ' green' }
        W ('<td><span class="' + $mcls + '"><b style="width:' + $m.pctDays + '%"></b></span>' +
           '<span class="pct">' + $m.pctDays + '%</span></td>')
        W ('<td class="n">' + $(if ($m.elapsed -ge 0) { "$($m.elapsed)%" } else { '&mdash;' }) + '</td>')
        W ('<td class="n">' + $(if ($m.lateT.Count   -gt 0) { '<span class="flag f-bad"  style="margin-left:0">' + $m.lateT.Count   + '</span>' } else { '0' }) + '</td>')
        W ('<td class="n">' + $(if ($m.atRiskT.Count -gt 0) { '<span class="flag f-warn" style="margin-left:0">' + $m.atRiskT.Count + '</span>' } else { '0' }) + '</td>')
        W ('<td class="d">' + $(if ($m.nextDue) { Show-D $m.nextDue } else { '&mdash;' }) + '</td>')
        $bp = Bar-In $pScale $m.effStart $m.effTarget
        if ($bp) {
            $parts = $bp -split '\|'
            $col = 'var(--r-green)'
            if ($m.rag -eq 'red')   { $col = 'var(--r-red)' }
            if ($m.rag -eq 'amber') { $col = 'var(--r-amber)' }
            if ($m.rag -eq 'done')  { $col = 'var(--r-done)' }
            W ('<td><div class="track mtrack"><b style="left:' + $parts[0] + '%;width:' + $parts[1] +
               '%;background:' + $col + $(if ($m.rag -eq 'done') { ';opacity:.45' } else { '' }) + '"></b>' + $pMark + '</div></td>')
        } else {
            W '<td><span class="flag f-mut">no dates</span></td>'
        }
        W '</tr>'
    }
    W '</tbody></table></div>'

    # --- what needs attention, across everything ---------------------------
    $all = @()
    foreach ($m in $ordered) {
        foreach ($k in @($m.lateT))   { $all += @{ k=$k; m=$m; ord=0 } }
        foreach ($k in @($m.atRiskT)) { $all += @{ k=$k; m=$m; ord=1 } }
    }
    W '<h2>What needs attention</h2>'
    W '<div class="card">'
    if ($all.Count -eq 0) {
        W ('<p style="margin:0"><span class="flag f-ok" style="margin-left:0">all clear</span> ' +
           'Nothing across the portfolio is past its end date or running out of its window unattended.</p>')
    } else {
        $shown = @($all | Sort-Object @{Expression={$_.ord}}, @{Expression={$_.k.target}})
        $cut = $shown
        if ($AttentionLimit -gt 0 -and $shown.Count -gt $AttentionLimit) {
            $cut = @($shown | Select-Object -First $AttentionLimit)
        }
        W ('<p class="sub" style="margin:0 0 12px">Every task across the portfolio that is late or running ' +
           'out of time unattended, worst first. Each one links to its project.</p>')
        W '<table><thead><tr><th style="width:86px">Risk</th><th>Task</th><th style="width:24%">Project</th>'
        W '<th style="width:26%">Why</th><th style="width:118px">Owner</th>'
        W '<th class="d" style="width:94px">End</th></tr></thead><tbody>'
        foreach ($row in $cut) {
            $k = $row.k; $m = $row.m
            W ('<tr><td>' + (Get-RiskBadge $k.risk) + '</td>' +
               '<td><span class="pid">#' + $k.id + '</span> ' + (Esc $k.summary) + '</td>' +
               '<td><a href="#p' + $m.P.id + '">' + (Esc $m.P.summary) + '</a></td>' +
               '<td class="why">' + (Esc $k.riskWhy) + '</td>' +
               '<td class="d">' + (Esc $k.agent) + '</td>' +
               '<td class="d">' + (Show-D $k.target) + '</td></tr>')
        }
        W '</tbody></table>'
        if ($cut.Count -lt $shown.Count) {
            W ('<p class="sub" style="margin:12px 0 0">Showing the first ' + $cut.Count + ' of ' + $shown.Count +
               ' &mdash; the rest are on their own project pages. Raise $AttentionLimit to show more.</p>')
        }
    }
    W '</div>'

    # --- how this was judged ------------------------------------------------
    W '<h2>How this was judged</h2><div class="card">'
    W ('<p style="margin:0 0 10px"><b>The colour comes from the tasks, one at a time.</b> ' +
       'Each open task is measured against its own window rather than the project&rsquo;s, so a task counts as ' +
       '<i>closing in</i> once ' + $NearEndPct + '% of its own span has gone, or when it is within ' + $NearEndDays +
       ' day(s) of its end date &mdash; whichever comes first. That way a ' + $NearEndDays +
       '-day gap means something different on a six-week task than on a three-day one.</p>')
    W '<ul style="margin:0 0 12px;padding-left:20px;font-size:13.5px;color:var(--text-secondary)">'
    W ('<li><b>Red</b> &mdash; something is already past its end date' +
       $(if ($LateInProgressIsRed) { ', whether or not anybody is working it' }
         else { ' and nobody is working it; a late task in progress is amber on this report' }) + '.</li>')
    W '<li><b>Amber</b> &mdash; something is closing in and nobody has picked it up: not started, or on hold.</li>'
    W '<li><b>Green</b> &mdash; everything open is either in progress, or not yet near its end date. A task that is closing in but in progress is listed, not punished.</li>'
    W '<li><b>Complete</b> &mdash; the project record itself is closed.</li>'
    W '</ul>'
    W ('<p style="margin:0 0 10px">The completion figure &mdash; closed tasks weighted by each one&rsquo;s length in days, ' +
       'so a three-week task carries more than a half-day one &mdash; is shown as context and does <b>not</b> set the colour. ' +
       'Missing dates and unassigned tasks are listed as housekeeping on each project page and do not set the colour either, ' +
       'unless tasks with no end date account for half or more of the work left &mdash; at which point there is not enough in Halo to judge.</p>')
    if ($Stats.unlisted.Count -gt 0) {
        W ('<p style="margin:0 0 10px"><b>Statuses assumed to mean in progress:</b> ' +
           (Esc ($Stats.unlisted -join ', ')) + '. They were not in the script&rsquo;s status lists, and in-progress ' +
           'is the option that never invents a problem &mdash; if one of those really means not-started or on-hold, ' +
           'it belongs in <code>$NotStartedStatuses</code> or <code>$OnHoldStatuses</code> at the top of the script.</p>')
    }
    W ('<p style="margin:0" class="sub">Generated ' + (Esc $Stamp) + ' from ' + (Esc $Tenant) +
       '. Read-only &mdash; nothing was changed in Halo.</p>')
    W '</div>'
    W '</section>'

    # =======================================================================
    #  A PAGE PER PROJECT
    # =======================================================================
    foreach ($m in $ordered) { Write-ProjectView -M $m -SB $sb }

    W '<button class="top" id="totop">Back to top</button>'
    W $script:JS
    W '</div></body></html>'

    return $sb.ToString()
}

# ---------------------------------------------------------------------------
#  RUN
# ---------------------------------------------------------------------------

if ($script:SkipRun) { return }   # set by the test harness; unused in normal runs

Write-Host ''
Write-Host '  HALO PROJECT PORTFOLIO REPORT - read-only, every project with drill-down' -ForegroundColor Cyan
Write-Host "  tenant $Tenant   closed projects: $(if($IncludeClosed){'included'}else{'excluded'})   manager: $AgentFilter"

$cid = Resolve-Credential -Inline $ClientId     -EnvName 'HALO_CLIENT_ID'     -Prompt 'Halo Client ID'
$sec = Resolve-Credential -Inline $ClientSecret -EnvName 'HALO_CLIENT_SECRET' -Prompt 'Halo Client Secret'

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
Set-Overall 1 4 'reading the ticket list'
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
if ($shortlist.Count -eq 0) {
    Clear-All
    Say "  Nothing to report. Check `$TYPE_PROJECT ($TYPE_PROJECT) and `$TYPE_TASK ($TYPE_TASK)" 'Yellow'
    Say '  against the ticket type ids in your tenant.' 'Yellow'
    return
}

# The list payload has no startdate, so each one needs an individual read.
Set-Overall 2 4 'reading each record for its dates'
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

# --- group -----------------------------------------------------------------
Head 'BUILDING'
Set-Overall 3 4 'building the model'

$projectRecords = @()
$tasksByParent  = @{}
$orphanTasks    = 0
foreach ($t in $full) {
    $tt = [int](Get-Val $t 'tickettype_id')
    if ($tt -eq $TYPE_PROJECT) { $projectRecords += $t; continue }
    $p = Get-Val $t 'parent_id'
    if ($null -eq $p -or [int]$p -le 0) { $orphanTasks++; continue }
    $key = [int]$p
    if (-not $tasksByParent.ContainsKey($key)) { $tasksByParent[$key] = New-Object System.Collections.ArrayList }
    [void]$tasksByParent[$key].Add($t)
}
Say "  $($projectRecords.Count) project(s), $($full.Count - $projectRecords.Count) task(s)" 'DarkGray'
if ($orphanTasks -gt 0) {
    Say "  $orphanTasks task(s) have no parent project and are not on this report" 'Yellow'
}

# --- a model per project ---------------------------------------------------
$models = New-Object System.Collections.ArrayList
$i = 0; $n = $projectRecords.Count
foreach ($proj in $projectRecords) {
    $i++
    Set-Step 'Working out each project' $i $n
    $pid_ = [int](Get-Val $proj 'id')
    $kids = @()
    if ($tasksByParent.ContainsKey($pid_)) { $kids = @($tasksByParent[$pid_]) }
    [void]$models.Add((Get-ProjectModel -Project $proj -TaskRecords $kids))
}
Clear-Step

# --- filters ---------------------------------------------------------------
$before = $models.Count
$models = @($models | Where-Object { $_.P.summary -like $ProjectFilter })
if ($AgentFilter -ne '*') { $models = @($models | Where-Object { $_.P.agent -like $AgentFilter }) }
if (-not $IncludeClosed)  { $models = @($models | Where-Object { -not $_.P.closed }) }
if ($models.Count -ne $before) {
    Say "  $($models.Count) of $before project(s) after filtering" 'DarkGray'
}
if ($models.Count -eq 0) {
    Clear-All
    Say '  No projects left after filtering - check $ProjectFilter, $AgentFilter and $IncludeClosed.' 'Yellow'
    return
}

# --- portfolio roll-up -----------------------------------------------------
$stats = @{
    red = 0; amber = 0; green = 0; done = 0
    tasks = 0; lateTasks = 0; atRiskTasks = 0; noTasks = 0
    unlisted = @()
}
foreach ($m in $models) {
    switch ($m.rag) {
        'red'   { $stats.red++ }
        'amber' { $stats.amber++ }
        'green' { $stats.green++ }
        'done'  { $stats.done++ }
    }
    $stats.tasks       += $m.totTask
    $stats.lateTasks   += $m.lateT.Count
    $stats.atRiskTasks += $m.atRiskT.Count
    if ($m.totTask -eq 0) { $stats.noTasks++ }
    foreach ($u in $m.unlisted) { if ($stats.unlisted -notcontains $u) { $stats.unlisted += $u } }
}
Say ("  RAG: {0} red, {1} amber, {2} green, {3} complete" -f $stats.red, $stats.amber, $stats.green, $stats.done) `
    $(if ($stats.red -gt 0) { 'Red' } elseif ($stats.amber -gt 0) { 'Yellow' } else { 'Green' })
foreach ($m in @($models | Where-Object { $_.rag -eq 'red' })) {
    Say ("    RED  " + $m.P.summary + " - " + (@($m.reasons)[0])) 'DarkGray'
}
if ($stats.unlisted.Count -gt 0) {
    Say ("  assumed to mean in progress: " + ($stats.unlisted -join ', ')) 'Yellow'
    Say '  if any of those really mean not-started or on-hold, add them to $NotStartedStatuses' 'Yellow'
    Say '  or $OnHoldStatuses at the top and run again - the verdicts depend on it.' 'Yellow'
}

# --- write -----------------------------------------------------------------
Head 'WRITING THE REPORT'
Set-Overall 4 4 'writing the report'
$stamp = Get-Date -Format 'dd/MM/yyyy HH:mm'
$html  = New-ReportHtml -Models $models -Stamp $stamp -Stats $stats

$path = $OutFile
if ([string]::IsNullOrWhiteSpace($path)) {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($docs)) { $docs = $env:TEMP }
    $path = Join-Path $docs ('Halo-Portfolio-' + (Get-Date -Format 'yyyyMMdd-HHmm') + '.html')
}
try {
    [System.IO.File]::WriteAllText($path, $html, (New-Object System.Text.UTF8Encoding($true)))
    Clear-All
    Write-Host ''
    Say "  Report written: $path" 'Green'
    Say ("  {0} project(s), {1} task(s), {2:N0} KB" -f $models.Count, $stats.tasks, ((Get-Item $path).Length / 1KB)) 'DarkGray'
    if ($OpenWhenDone) { Start-Process $path }
}
catch {
    Clear-All
    Say "  Could not write $path : $($_.Exception.Message)" 'Red'
    Say '  Set $OutFile to somewhere you can write and run again.' 'Yellow'
}
Write-Host ''
