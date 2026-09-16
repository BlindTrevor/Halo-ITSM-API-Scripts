<#
    New-HaloProject.ps1
    -------------------
    TEMPLATE. Creates one project in Halo ITSM and a set of project tasks
    hanging off it, from a table of records held at the top of this file.

    Open it in PowerShell ISE (or any editor + console) and press F5. No
    parameters, no external files - everything you change is in the two
    settings blocks below.

      1. Leave $Apply as $false the first time. That is a DRY RUN - it gets a
         token, reads your ticket types, priorities and agents, validates the
         record table, prints every record it would create, and writes
         nothing.
      2. Fill in $ProjectTypeId, $TaskTypeId and the two maps from what the
         dry run prints.
      3. Replace $ProjectSummary, the dates and the $Records table with your
         own. Run the dry run again and read it properly.
      4. Set $Apply = $true and run to create.

    This script CREATES records and there is no undo, so the dry run defaults
    to on and stays on until you change it.

    Re-running is safe: before creating anything it reads the tasks already
    hanging off the project and skips any whose summary it already sees. If
    the project itself already exists, put its id in $ExistingProjectId.

    EDIT CHECKLIST
      $ProjectSummary / $ProjectStart / $ProjectTarget / $ProjectBlurb
      $ProjectTypeId / $TaskTypeId          (from the dry run)
      $PriorityMap / $AgentMap              (from the dry run)
      $HaloClientId / $HaloSiteId           (optional)
      $Records                              (your tasks)
      $Tenant                               (your Halo tenant)

    RECORD SCHEMA - one [pscustomobject] per task:
      id           short unique key, e.g. A1. Sorts and prefixes the summary,
                   and is what the duplicate check matches on.
      summary      the ticket summary.
      workstream   free text, e.g. 'A - Scoping'. The leading token before
                   the first '-' is used for the project's workstream legend.
      owner        free text, shown in the ticket details.
      agentkey     a key in $AgentMap - who the ticket is assigned to.
      priority     a key in $PriorityMap.
      start/target yyyy-MM-dd. Must sit inside the project window.
      days         optional. Working days; computed from the dates if 0.
      deps         free text, e.g. 'A1, B2', or '-' for none. Documentation
                   only - Halo is not told about the dependency.
      description  the body of the ticket.
      donewhen     the acceptance test, appended to the description.

    API key: Halo > Configuration > Integrations > Halo API > View
    Applications > New. Authentication Method "Client ID and Secret
    (Services)", Login Type "Agent", permission edit:tickets.
#>

# ===========================================================================
#  THE ONLY SWITCH YOU NEED
# ===========================================================================

$Apply = $false        # $false = dry run (safe).  $true = create records.

# Paste your Halo API credentials between the quotes. Leave them empty and
# the script falls back to $env:HALO_CLIENT_ID / $env:HALO_CLIENT_SECRET,
# and then to prompting you.
$ClientId     = ''
$ClientSecret = ''

$Tenant       = 'contoso'      # your Halo tenant name

# ===========================================================================
#  SETTINGS - THE PROJECT
# ===========================================================================

$ProjectSummary = 'Example Project - rename me'
$ProjectStart   = '2026-10-01'
$ProjectTarget  = '2026-12-18'

# Free text for the top of the project ticket. The workstream legend and the
# task count are appended automatically from $Records, so do not list them
# here - they will only drift.
$ProjectBlurb = @(
    'One line on what this project is for.'
    'One line on where it came from - a meeting, a ticket, a decision.'
) -join "`r`n"

# Stamped at the bottom of every task's details. Set it to the meeting,
# document or ticket the task list came from.
$SourceNote = 'Created by New-HaloProject.ps1.'

$ExistingProjectId = 0   # 0 = create the project. After a successful run the
                         # script prints the new id - paste it here to add or
                         # re-run tasks against the same project.

# ===========================================================================
#  SETTINGS - YOUR TENANT
# ===========================================================================

# ---------------------------------------------------------------------------
# These two are REQUIRED and are specific to your tenant. Run the dry run
# once with both at 0: it prints the ticket types it can see, with their ids.
# ---------------------------------------------------------------------------
$ProjectTypeId = 0        # ticket type id used for Projects
$TaskTypeId    = 0        # ticket type id used for Project Tasks

$HaloClientId  = 0        # customer/client record id, 0 to let Halo default
$HaloSiteId    = 0        # site id, 0 to let Halo default

# How a task is attached to its project. parent_id is the usual field; the
# dry run cannot test it, so the first task created is checked and the run
# stops if the link did not stick. Change this only if that check fails.
$LinkField     = 'parent_id'

# Fill these from what the dry run prints. Any left at 0 is simply not sent,
# and Halo applies its own default. The keys are what $Records refers to, so
# rename keys and records together.
$PriorityMap = @{
    'Critical' = 0
    'High'     = 0
    'Medium'   = 0
    'Low'      = 0
}

$AgentMap = @{
    'Alex'   = 0
    'Jordan' = 0
}

$ProjectOwnerKey = 'Alex'      # which $AgentMap key owns the project ticket

$PrefixSummaryWithId = $true   # "A1 - Confirm scope and success criteria"
                               # Keeps them in order in Halo and is what the
                               # duplicate check matches on. Leave it on.

$OnlyIds = @()                 # Optional: limit to certain task ids,
                               # e.g. @('B1','B2'). Empty = all of them.

# ===========================================================================
#  THE TASKS
# ===========================================================================
# Replace everything below with your own. The six here are a worked example
# of the shape, not a suggestion - delete them.

$Records = @(
    [pscustomobject]@{ id='A1'; summary='Confirm scope and success criteria'; workstream='A - Scoping and approvals'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-10-01'; target='2026-10-07'; days=5; deps='-'; description='Write down what is in and out of scope, who the stakeholders are, and what the project has to achieve to count as finished. Circulate for comment rather than assuming agreement.'; donewhen='Scope agreed in writing by the stakeholders named above.' }
    [pscustomobject]@{ id='A2'; summary='Obtain budget approval'; workstream='A - Scoping and approvals'; owner='Alex'; agentkey='Alex'; priority='Critical'; start='2026-10-08'; target='2026-10-14'; days=5; deps='A1'; description='Cost the work from the agreed scope, including anything recurring. Take it to the budget holder with the options and a recommendation rather than a single number.'; donewhen='Approval recorded in writing and the PO raised.' }
    [pscustomobject]@{ id='B1'; summary='Order hardware and prepare the environment'; workstream='B - Build'; owner='Jordan'; agentkey='Jordan'; priority='High'; start='2026-10-15'; target='2026-10-28'; days=10; deps='A2'; description='Place the order and track the lead time. In parallel prepare whatever the build depends on - accounts, licences, network, firewall changes through the normal change process.'; donewhen='Kit delivered and the environment ready to build into.' }
    [pscustomobject]@{ id='B2'; summary='Configure and test'; workstream='B - Build'; owner='Jordan'; agentkey='Jordan'; priority='High'; start='2026-10-29'; target='2026-11-20'; days=17; deps='B1'; description='Build to the agreed scope. Test against the success criteria from A1, not against what was convenient to build. Record what was tested and what failed first time.'; donewhen='Test results recorded and every failure either fixed or accepted in writing.' }
    [pscustomobject]@{ id='C1'; summary='Document and publish a knowledgebase article'; workstream='C - Handover'; owner='Jordan'; agentkey='Jordan'; priority='Medium'; start='2026-11-23'; target='2026-11-27'; days=5; deps='B2'; description='Write up how it works, how to support it and who to call, and publish it in the Halo Knowledgebase so it is not knowledge held by one person.'; donewhen='Article published and linked from this project.' }
    [pscustomobject]@{ id='C2'; summary='Handover to support and close the project'; workstream='C - Handover'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-11-30'; target='2026-12-18'; days=15; deps='C1'; description='Walk the support team through it, agree who owns it from here, and confirm monitoring and backups are in place. Close any tasks still open or move them to business as usual with an owner.'; donewhen='Support have accepted ownership and every task on this project is closed or reassigned.' }
)

# ===========================================================================
#  NOTHING BELOW HERE NEEDS EDITING
# ===========================================================================

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$AuthUrl   = "https://$Tenant.haloitsm.com/auth/token"
$ApiBase   = "https://$Tenant.haloitsm.com/api"
$TimeOfDay = 'T12:00:00'
$Scope     = 'edit:tickets'

# --- credentials ----------------------------------------------------------
function Resolve-Credential {
    # In order: the value typed at the top of this script, then the matching
    # environment variable, then a masked prompt.
    param([string]$Inline, [string]$EnvName, [string]$Prompt)

    if (-not [string]::IsNullOrWhiteSpace($Inline)) { return $Inline.Trim() }

    $fromEnv = [Environment]::GetEnvironmentVariable($EnvName)
    if (-not [string]::IsNullOrWhiteSpace($fromEnv)) { return $fromEnv.Trim() }

    $secure = Read-Host -Prompt $Prompt -AsSecureString
    $bstr   = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try   { return ([Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)).Trim() }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# --- record table ---------------------------------------------------------
function ConvertTo-Date {
    # Records are authored as yyyy-MM-dd strings. Parse invariantly so a
    # machine set to dd/MM/yyyy cannot reinterpret them.
    param([string]$Value)
    $parsed = [datetime]::MinValue
    $ok = [datetime]::TryParseExact($Value, 'yyyy-MM-dd',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$parsed)
    if ($ok) { return $parsed }
    return $null
}

function Get-WorkingDays {
    # Inclusive weekday count. Knows nothing about bank holidays - it is only
    # used to fill in the details block when a record omits 'days'.
    param([datetime]$From, [datetime]$To)
    if ($To -lt $From) { return 0 }
    $n = 0
    for ($d = $From; $d -le $To; $d = $d.AddDays(1)) {
        if ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday') { $n++ }
    }
    return $n
}

function Test-Records {
    # Everything that can be checked without calling Halo. Returns the list of
    # problems; an empty list means the table is sound.
    param($All, $Selected)

    $problems = @()
    $ids = @($All | ForEach-Object { "$($_.id)".Trim() })

    $dupes = $ids | Group-Object | Where-Object { $_.Count -gt 1 }
    foreach ($d in $dupes) { $problems += "id '$($d.Name)' is used $($d.Count) times - ids must be unique" }

    $projStart = ConvertTo-Date $ProjectStart
    $projEnd   = ConvertTo-Date $ProjectTarget
    if (-not $projStart) { $problems += "`$ProjectStart '$ProjectStart' is not yyyy-MM-dd" }
    if (-not $projEnd)   { $problems += "`$ProjectTarget '$ProjectTarget' is not yyyy-MM-dd" }
    if ($projStart -and $projEnd -and $projEnd -lt $projStart) {
        $problems += "`$ProjectTarget is before `$ProjectStart"
    }

    foreach ($r in $Selected) {
        $id = "$($r.id)".Trim()
        if ([string]::IsNullOrWhiteSpace($id))         { $problems += 'a record has no id'; continue }
        if ([string]::IsNullOrWhiteSpace($r.summary))  { $problems += "$id has no summary" }

        $s = ConvertTo-Date $r.start
        $t = ConvertTo-Date $r.target
        if (-not $s) { $problems += "$id start '$($r.start)' is not yyyy-MM-dd" }
        if (-not $t) { $problems += "$id target '$($r.target)' is not yyyy-MM-dd" }
        if ($s -and $t) {
            if ($t -lt $s) { $problems += "$id target $($r.target) is before start $($r.start)" }
            if ($projStart -and $s -lt $projStart) { $problems += "$id starts $($r.start), before the project starts $ProjectStart" }
            if ($projEnd   -and $t -gt $projEnd)   { $problems += "$id targets $($r.target), after the project targets $ProjectTarget" }
        }

        if ($r.agentkey -and -not $AgentMap.ContainsKey($r.agentkey)) {
            $problems += "$id agentkey '$($r.agentkey)' is not a key in `$AgentMap"
        }
        if ($r.priority -and -not $PriorityMap.ContainsKey($r.priority)) {
            $problems += "$id priority '$($r.priority)' is not a key in `$PriorityMap"
        }

        # deps are documentation only, but a typo in one is still worth saying
        if ($r.deps -and "$($r.deps)".Trim() -ne '-') {
            foreach ($dep in ("$($r.deps)" -split ',')) {
                $dep = $dep.Trim()
                if ($dep -and $ids -notcontains $dep) { $problems += "$id depends on '$dep', which is not an id in `$Records" }
            }
        }
    }

    return $problems
}

function Get-WorkstreamLegend {
    # The distinct workstreams, in first-seen order, for the project ticket.
    param($Selected)
    $seen = @()
    foreach ($r in $Selected) {
        $w = "$($r.workstream)".Trim()
        if ($w -and $seen -notcontains $w) { $seen += $w }
    }
    return $seen
}

function Get-WorkstreamTag {
    # 'A - Scoping and approvals' -> 'A'. Falls back to the first character,
    # then to a dash, so the summary table never throws on odd input.
    param($Workstream)
    $w = "$Workstream".Trim()
    if (-not $w) { return '-' }
    $tag = ($w -split '\s*-\s*')[0].Trim()
    if ($tag -and $tag.Length -le 3) { return $tag }
    return $w.Substring(0, 1)
}

# --- start ----------------------------------------------------------------
$clientIdValue     = Resolve-Credential -Inline $ClientId     -EnvName 'HALO_CLIENT_ID'     -Prompt 'Halo Client ID'
$clientSecretValue = Resolve-Credential -Inline $ClientSecret -EnvName 'HALO_CLIENT_SECRET' -Prompt 'Halo Client Secret'

if ([string]::IsNullOrWhiteSpace($clientIdValue) -or [string]::IsNullOrWhiteSpace($clientSecretValue)) {
    throw 'No credentials supplied - stopping.'
}

if (-not $Records -or $Records.Count -eq 0) { throw '$Records is empty - nothing to create.' }

$work = if ($OnlyIds.Count -gt 0) { $Records | Where-Object { $OnlyIds -contains $_.id } } else { $Records }
$work = @($work)
if ($work.Count -eq 0) { throw 'Nothing to do - check $OnlyIds against the ids in $Records.' }

$problems = Test-Records -All $Records -Selected $work
if ($problems.Count -gt 0) {
    Write-Host ''
    Write-Host "  $($problems.Count) problem(s) in `$Records - nothing has been created:" -ForegroundColor Red
    $problems | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    Write-Host ''
    return
}

# --- token ----------------------------------------------------------------
function Format-ErrorBody {
    # A bad tenant or a server-side fault answers with a whole error PAGE -
    # a quarter of a megabyte of CSS and inline base64 images, which buries
    # the console. The useful case is the small JSON body Halo returns for a
    # bad credential, so keep that intact and reduce everything else to the
    # first readable sentence.
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
    # A 400 from Halo carries a JSON body naming the real problem
    # (invalid_client, invalid_scope, unsupported_grant_type...). PowerShell
    # hides it behind a generic message, so dig it out. Works on 5.1 and 7.
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return Format-ErrorBody $ErrorRecord.ErrorDetails.Message
    }
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($null -eq $resp) { return '' }
        $stream = $resp.GetResponseStream()
        $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream)
        return Format-ErrorBody $reader.ReadToEnd()
    }
    catch { return '' }
}

function Request-HaloToken {
    param([hashtable]$Form)
    return (Invoke-RestMethod -Uri $AuthUrl -Method Post -Body $Form `
            -ContentType 'application/x-www-form-urlencoded').access_token
}

Write-Host ''
Write-Host '  Requesting token...' -ForegroundColor DarkGray

$baseForm = @{
    grant_type    = 'client_credentials'
    client_id     = $clientIdValue
    client_secret = $clientSecretValue
    scope         = $Scope
    tenant        = $Tenant      # required for cloud-hosted Halo
}

$token = $null
try { $token = Request-HaloToken -Form $baseForm }
catch {
    $body = Get-WebErrorBody $_
    Write-Host ''
    Write-Host "  First attempt failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
    if ($body) { Write-Host "  Server said: $body" -ForegroundColor DarkYellow }
    Write-Host '  Trying other combinations...' -ForegroundColor DarkGray

    $attempts = @(
        @{ Label = "scope='all' with tenant";         Scope = 'all';          Tenant = $Tenant }
        @{ Label = "scope='edit:tickets', no tenant"; Scope = 'edit:tickets'; Tenant = $null }
        @{ Label = "scope='all', no tenant";          Scope = 'all';          Tenant = $null }
    )

    foreach ($a in $attempts) {
        $form = @{
            grant_type    = 'client_credentials'
            client_id     = $clientIdValue
            client_secret = $clientSecretValue
            scope         = $a.Scope
        }
        if ($a.Tenant) { $form.tenant = $a.Tenant }

        try {
            $token = Request-HaloToken -Form $form
            Write-Host ''
            Write-Host "  Worked: $($a.Label)" -ForegroundColor Green
            Write-Host "  Set `$Scope = '$($a.Scope)' at the top to skip this next time." -ForegroundColor Green
            Write-Host ''
            break
        }
        catch {
            $b = Get-WebErrorBody $_
            $detail = if ($b) { " - $b" } else { '' }
            Write-Host "    $($a.Label): failed$detail" -ForegroundColor DarkGray
        }
    }
}

if (-not $token) {
    Write-Host ''
    Write-Host '  Could not get a token with any combination.' -ForegroundColor Red
    Write-Host ''
    Write-Host '  Check these, in Halo > Configuration > Integrations > Halo API:' -ForegroundColor Yellow
    Write-Host '   1. The "Authorisation Server" URL shown on that page. If it is not'
    Write-Host "      $AuthUrl"
    Write-Host "      then set `$AuthUrl at the top of this script to match it."
    Write-Host '   2. Open your application > the Client ID and Secret match what you pasted.'
    Write-Host '   3. Authentication Method is "Client ID and Secret (Services)".'
    Write-Host '   4. Login Type is "Agent" AND an agent is actually selected underneath.'
    Write-Host '   5. The Permissions tab has edit:tickets ticked.'
    Write-Host ''
    return
}

$headers = @{ Authorization = "Bearer $token" }

# --- helpers --------------------------------------------------------------
function Get-HaloList {
    # Halo returns either a bare array or an object wrapping one, depending on
    # the endpoint. Normalise so callers do not have to care.
    param([string]$Path, [string]$Property)
    try {
        $r = Invoke-RestMethod -Uri "$ApiBase/$Path" -Headers $headers -Method Get
        if ($null -eq $r) { return @() }
        if ($Property -and $r.PSObject.Properties.Name -contains $Property) { return @($r.$Property) }
        if ($r -is [array]) { return $r }
        foreach ($p in @('tickets','types','agents','priorities','record','records')) {
            if ($r.PSObject.Properties.Name -contains $p) { return @($r.$p) }
        }
        return @($r)
    }
    catch {
        Write-Host "    could not read /$Path - $($_.Exception.Message)" -ForegroundColor DarkYellow
        return @()
    }
}

function Format-D {
    # Halo returns dates as JSON strings, but Invoke-RestMethod deserialises
    # them into [datetime], whose default string form follows the machine
    # locale. Normalise to yyyy-MM-dd before comparing anything.
    param($Value)
    if ($null -eq $Value) { return 'not set' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }

    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return 'not set' }
    if ($s -match '^(\d{4}-\d{2}-\d{2})') { return $Matches[1] }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($s, [System.Globalization.CultureInfo]::InvariantCulture,
                             [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.ToString('yyyy-MM-dd')
    }
    return $s
}

function Get-TaskSummary {
    param($Rec)
    if ($PrefixSummaryWithId) { return "$($Rec.id) - $($Rec.summary)" }
    return $Rec.summary
}

function Build-Details {
    param($Rec)

    $days = 0
    if ($Rec.PSObject.Properties.Name -contains 'days' -and $Rec.days) { $days = [int]$Rec.days }
    if ($days -le 0) {
        $s = ConvertTo-Date $Rec.start
        $t = ConvertTo-Date $Rec.target
        if ($s -and $t) { $days = Get-WorkingDays -From $s -To $t }
    }

    $lines = @(
        "Workstream:  $($Rec.workstream)"
        "Owner:       $($Rec.owner)"
        "Priority:    $($Rec.priority)"
        "Planned:     $($Rec.start) to $($Rec.target)  ($days working days)"
        "Depends on:  $($Rec.deps)"
        ''
        $Rec.description
        ''
        "Done when: $($Rec.donewhen)"
        ''
        $SourceNote
    )
    return ($lines -join "`r`n")
}

function New-HaloTicket {
    # POST /api/Tickets takes an array. No id in the payload means create.
    param([hashtable]$Fields)
    $payload = ConvertTo-Json -Depth 5 -InputObject @($Fields)
    $r = Invoke-RestMethod -Uri "$ApiBase/Tickets" -Headers $headers -Method Post `
                           -Body $payload -ContentType 'application/json'
    if ($r -is [array]) { return $r[0] }
    return $r
}

function Get-ParentId {
    # The field comes back under different spellings depending on version.
    param($Ticket)
    foreach ($n in @('parent_id','parentid','parent_ticket_id')) {
        if ($Ticket.PSObject.Properties.Name -contains $n -and $Ticket.$n) { return [int]$Ticket.$n }
    }
    return 0
}

# --- discovery ------------------------------------------------------------
Write-Host ''
Write-Host '  Reading tenant configuration...' -ForegroundColor DarkGray

$types = Get-HaloList -Path 'TicketType' -Property 'ticket_types'
if ($types.Count -gt 0) {
    Write-Host ''
    Write-Host '  Ticket types (use these for $ProjectTypeId / $TaskTypeId):' -ForegroundColor Cyan
    foreach ($t in ($types | Sort-Object -Property id)) {
        $flag = ''
        if ("$($t.name)" -match 'project') { $flag = '   <-- looks like a project type' }
        Write-Host ("    {0,-6} {1}{2}" -f $t.id, $t.name, $flag)
    }
}

if (-not $Apply) {
    $prios = Get-HaloList -Path 'Priority' -Property 'priorities'
    if ($prios.Count -gt 0) {
        Write-Host ''
        Write-Host '  Priorities (use these for $PriorityMap):' -ForegroundColor Cyan
        foreach ($p in $prios) { Write-Host ("    {0,-6} {1}" -f $p.id, $p.name) }
    }

    $agents = Get-HaloList -Path 'Agent' -Property 'agents'
    if ($agents.Count -gt 0) {
        # Match on whatever keys are in $AgentMap, so renaming the people in
        # $Records is the only edit needed here.
        $pattern = (($AgentMap.Keys | ForEach-Object { [regex]::Escape($_) }) -join '|')
        $hits = if ($pattern) { $agents | Where-Object { "$($_.name)" -match $pattern } } else { @() }
        if ($hits) {
            Write-Host ''
            Write-Host '  Agents matching the $AgentMap keys (use these for $AgentMap):' -ForegroundColor Cyan
            foreach ($a in $hits) { Write-Host ("    {0,-6} {1}" -f $a.id, $a.name) }
        }
        else {
            Write-Host ''
            Write-Host "  No agent name matched the `$AgentMap keys. Here are the first 25 of" -ForegroundColor DarkYellow
            Write-Host "  $($agents.Count) agent(s) - rename the keys to match:" -ForegroundColor DarkYellow
            foreach ($a in ($agents | Select-Object -First 25)) { Write-Host ("    {0,-6} {1}" -f $a.id, $a.name) }
        }
    }
}

if ($ProjectTypeId -le 0 -or $TaskTypeId -le 0) {
    Write-Host ''
    Write-Host '  $ProjectTypeId and $TaskTypeId are not set.' -ForegroundColor Yellow
    Write-Host '  Set them from the list above, then run again.' -ForegroundColor Yellow
    if ($Apply) {
        Write-Host '  Nothing has been created.' -ForegroundColor Yellow
        Write-Host ''
        return
    }
}

# --- what we would create -------------------------------------------------
$mode = if ($Apply) { 'APPLY - CREATING RECORDS' } else { 'DRY RUN - nothing will be created' }
$col  = if ($Apply) { 'Yellow' } else { 'Cyan' }
$legend = Get-WorkstreamLegend -Selected $work

Write-Host ''
Write-Host "  Halo project build - $mode" -ForegroundColor $col
Write-Host ''
Write-Host "  Project: $ProjectSummary"
Write-Host "           $ProjectStart -> $ProjectTarget"
Write-Host "  Tasks:   $($work.Count)"
if ($legend.Count -gt 0) {
    Write-Host "  Streams: $($legend.Count)"
    foreach ($w in $legend) { Write-Host "             $w" -ForegroundColor DarkGray }
}
Write-Host ''
Write-Host ('  {0,-6}{1,-5}{2,-26}{3,-10}{4}' -f 'ID','WS','Dates','Owner','Summary')
Write-Host ('  ' + ('-' * 118)) -ForegroundColor DarkGray

foreach ($rec in $work) {
    $name = "$($rec.summary)"
    if ($name.Length -gt 52) { $name = $name.Substring(0, 52) }
    Write-Host ('  {0,-6}{1,-5}{2,-26}{3,-10}{4}' -f `
        $rec.id, (Get-WorkstreamTag $rec.workstream), "$($rec.start) -> $($rec.target)", $rec.agentkey, $name)
}
Write-Host ('  ' + ('-' * 118)) -ForegroundColor DarkGray

if (-not $Apply) {
    Write-Host ''
    Write-Host "  Would create 1 project and $($work.Count) task(s). Nothing has been written." -ForegroundColor Yellow
    Write-Host '  If the type ids and the list above look right, set $Apply = $true' -ForegroundColor Yellow
    Write-Host '  near the top and run again.' -ForegroundColor Yellow
    Write-Host ''
    return
}

# --- the project ----------------------------------------------------------
$projectId = $ExistingProjectId

if ($projectId -gt 0) {
    Write-Host ''
    Write-Host "  Using existing project id $projectId." -ForegroundColor DarkGray
}
else {
    $detailLines = @($ProjectBlurb, '', "Tasks: $($work.Count).")
    if ($legend.Count -gt 0) {
        $detailLines += ''
        $detailLines += 'Workstreams:'
        foreach ($w in $legend) { $detailLines += "  $w" }
    }
    $detailLines += ''
    $detailLines += $SourceNote

    $projFields = @{
        tickettype_id = $ProjectTypeId
        summary       = $ProjectSummary
        details       = ($detailLines -join "`r`n")
        startdate     = "$ProjectStart$TimeOfDay"
        targetdate    = "$ProjectTarget$TimeOfDay"
    }
    if ($HaloClientId -gt 0) { $projFields.client_id = $HaloClientId }
    if ($HaloSiteId   -gt 0) { $projFields.site_id   = $HaloSiteId }
    if ($ProjectOwnerKey -and $AgentMap[$ProjectOwnerKey] -gt 0) { $projFields.agent_id = $AgentMap[$ProjectOwnerKey] }

    Write-Host ''
    Write-Host '  Creating the project...' -ForegroundColor DarkGray
    try {
        $projRecord = New-HaloTicket -Fields $projFields
        $projectId = [int]$projRecord.id
    }
    catch {
        $b = Get-WebErrorBody $_
        Write-Host "  FAILED to create the project: $($_.Exception.Message)" -ForegroundColor Red
        if ($b) { Write-Host "  Server said: $b" -ForegroundColor Red }
        Write-Host ''
        return
    }

    if ($projectId -le 0) {
        Write-Host '  Halo accepted the request but returned no id - stopping before the tasks.' -ForegroundColor Red
        Write-Host ''
        return
    }

    Write-Host "  Created project id $projectId ($($projRecord.ref))" -ForegroundColor Green
    Write-Host "  Paste that id into `$ExistingProjectId to re-run tasks against it." -ForegroundColor Green
}

# --- skip anything already there ------------------------------------------
Write-Host ''
Write-Host '  Checking for tasks that already exist...' -ForegroundColor DarkGray

$existing = @()
$children = Get-HaloList -Path "Tickets?parent_id=$projectId&count=500" -Property 'tickets'
foreach ($c in $children) {
    # The filter may be ignored by older versions, so confirm the parent
    # client-side rather than trusting the server to have honoured it.
    if ((Get-ParentId $c) -eq $projectId) { $existing += "$($c.summary)".Trim() }
}
Write-Host "  $($existing.Count) task(s) already on this project." -ForegroundColor DarkGray

# --- the tasks ------------------------------------------------------------
Write-Host ''
$created = @(); $skipped = @(); $failed = @()
$linkChecked = $false

foreach ($rec in $work) {
    $summary = Get-TaskSummary -Rec $rec

    if ($existing -contains $summary) {
        Write-Host ("  {0,-6} skipped - already on the project" -f $rec.id) -ForegroundColor DarkGray
        $skipped += $rec
        continue
    }

    $fields = @{
        tickettype_id = $TaskTypeId
        summary       = $summary
        details       = Build-Details -Rec $rec
        startdate     = "$($rec.start)$TimeOfDay"
        targetdate    = "$($rec.target)$TimeOfDay"
    }
    $fields[$LinkField] = $projectId

    if ($HaloClientId -gt 0) { $fields.client_id = $HaloClientId }
    if ($HaloSiteId   -gt 0) { $fields.site_id   = $HaloSiteId }

    $prioId = $PriorityMap[$rec.priority]
    if ($prioId -and $prioId -gt 0) { $fields.priority_id = $prioId }

    $agtId = $AgentMap[$rec.agentkey]
    if ($agtId -and $agtId -gt 0) { $fields.agent_id = $agtId }

    try {
        $new = New-HaloTicket -Fields $fields
        $newId = [int]$new.id
        Write-Host ("  {0,-6} created  id {1,-8} {2}" -f $rec.id, $newId, $summary)
        $created += [pscustomobject]@{ Rec = $rec; Id = $newId; Summary = $summary }
    }
    catch {
        $b = Get-WebErrorBody $_
        $detail = if ($b) { " - $b" } else { '' }
        Write-Host ("  {0,-6} FAILED: {1}{2}" -f $rec.id, $_.Exception.Message, $detail) -ForegroundColor Red
        $failed += [pscustomobject]@{ Id = $rec.id; Error = "$($_.Exception.Message)$detail" }
        continue
    }

    # Canary: check the first task actually attached to the project before
    # creating the rest. A dry run cannot test this, so it is tested here.
    if (-not $linkChecked) {
        $linkChecked = $true
        Start-Sleep -Milliseconds 400
        try {
            $check = Invoke-RestMethod -Uri "$ApiBase/Tickets/$($created[0].Id)" -Headers $headers -Method Get
            if ((Get-ParentId $check) -ne $projectId) {
                Write-Host ''
                Write-Host "  STOPPING. The first task was created (id $($created[0].Id)) but it is not" -ForegroundColor Red
                Write-Host "  attached to project $projectId - '$LinkField' is not the linking field in" -ForegroundColor Red
                Write-Host '  this tenant. Open that task in Halo, see how a project task is linked,' -ForegroundColor Red
                Write-Host "  set `$LinkField accordingly, delete the one task, put the project id in" -ForegroundColor Red
                Write-Host '  $ExistingProjectId and run again. Only one record has been created.' -ForegroundColor Red
                Write-Host ''
                return
            }
            Write-Host "         link confirmed - tasks are attaching to project $projectId" -ForegroundColor DarkGray
        }
        catch {
            Write-Host '         (could not verify the project link - continuing)' -ForegroundColor DarkYellow
        }
    }

    Start-Sleep -Milliseconds 150
}

# --- summary --------------------------------------------------------------
Write-Host ''
Write-Host ('  ' + ('-' * 118)) -ForegroundColor DarkGray
Write-Host "  Project id: $projectId"
Write-Host "  Created: $($created.Count)    Skipped as existing: $($skipped.Count)    Failed: $($failed.Count)"

if ($failed.Count -gt 0) {
    Write-Host ''
    Write-Host '  Failures:' -ForegroundColor Red
    $failed | ForEach-Object { Write-Host "    $($_.Id): $($_.Error)" }
}

if ($created.Count -eq 0) { Write-Host ''; return }

Write-Host ''
Write-Host '  Verifying every created task...' -ForegroundColor DarkGray
$bad = @()
foreach ($c in $created) {
    try {
        $t = Invoke-RestMethod -Uri "$ApiBase/Tickets/$($c.Id)" -Headers $headers -Method Get
        $problems = @()
        if ((Format-D $t.startdate)  -ne $c.Rec.start)  { $problems += "start $(Format-D $t.startdate)" }
        if ((Format-D $t.targetdate) -ne $c.Rec.target) { $problems += "target $(Format-D $t.targetdate)" }
        if ((Get-ParentId $t) -ne $projectId)           { $problems += 'not linked to the project' }
        if ($problems.Count -gt 0) {
            $bad += [pscustomobject]@{ Id = $c.Rec.id; HaloId = $c.Id; Problem = ($problems -join ', ') }
        }
    }
    catch {
        $bad += [pscustomobject]@{ Id = $c.Rec.id; HaloId = $c.Id; Problem = "could not re-read: $($_.Exception.Message)" }
    }
}

Write-Host ''
if ($bad.Count -gt 0) {
    Write-Host "  $($bad.Count) task(s) did not come back as sent:" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host "    $($_.Id) (halo id $($_.HaloId)): $($_.Problem)" }
    Write-Host ''
    Write-Host '  Dates that shift on their own usually mean Halo is applying a working' -ForegroundColor Red
    Write-Host '  calendar or an SLA to the ticket type. Check the ticket type''s SLA and' -ForegroundColor Red
    Write-Host '  working-hours settings, then correct the dates on the listed tickets.' -ForegroundColor Red
}
else {
    Write-Host "  All $($created.Count) verified - dates and project link correct." -ForegroundColor Green
}
Write-Host ''
