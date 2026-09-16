<#
    New-HaloSolarProject-ISE.ps1
    ----------------------------
    TEMPLATE / EXAMPLE. Creates a "Solar PV Estate" project in Halo ITSM and
    its 30 project tasks, from a facilities meeting held 15 September 2026.

    Every name, site, supplier, account and reference below is a placeholder.
    Treat the 30 records as a worked example of the shape - workstream, owner,
    dates, dependencies, description, done-when - and replace the content with
    your own before running it anywhere real.

    Same shape as Apply-HaloDates-ISE.ps1: open it in PowerShell ISE and
    press F5. No parameters, no external files.

      1. Leave $Apply as $false the first time. That is a DRY RUN - it gets a
         token, reads your ticket types, priorities and agents, prints every
         record it would create, and writes nothing.
      2. Fill in $ProjectTypeId, $TaskTypeId and the two maps from what the
         dry run prints.
      3. Run the dry run again and read it properly.
      4. Set $Apply = $true and run to create.

    This script CREATES records. Unlike the date script there is no undo, so
    the dry run defaults to on and stays on until you change it.

    Re-running is safe: before creating anything it reads the tasks already
    hanging off the project and skips any whose summary it already sees. If
    the project itself already exists, put its id in $ExistingProjectId.

    API key: Halo > Configuration > Integrations > Halo API > View
    Applications > New. Authentication Method "Client ID and Secret
    (Services)", Login Type "Agent", permission edit:tickets.
#>

# ===========================================================================
#  THE ONLY SWITCH YOU NEED
# ===========================================================================

$Apply = $false        # $false = dry run (safe).  $true = create records.

# Paste your Halo API credentials between the quotes.
$ClientId     = ''
$ClientSecret = ''

# ===========================================================================
#  WHAT GETS CREATED
# ===========================================================================

$ProjectSummary = 'Solar PV Estate - Monitoring, Supply and Remediation'
$ProjectStart   = '2026-09-16'
$ProjectTarget  = '2027-01-29'

$ExistingProjectId = 0   # 0 = create the project. After a successful run the
                         # script prints the new id - paste it here to add or
                         # re-run tasks against the same project.

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
# and Halo applies its own default.
$PriorityMap = @{
    'Critical' = 0
    'High'     = 0
    'Medium'   = 0
    'Low'      = 0
}

# The keys are the agentkey values used in $Records below - rename both
# together if you replace the cast.
$AgentMap = @{
    'Alex'   = 0
    'Jordan' = 0
}

$PrefixSummaryWithId = $true   # "A1 - Create Facilities SharePoint site..."
                               # Keeps them in order in Halo and is what the
                               # duplicate check matches on. Leave it on.

$OnlyIds = @()                 # Optional: limit to certain task ids,
                               # e.g. @('B1','B3'). Empty = all 30.

# ===========================================================================
# NOTE: if you leave the credentials empty, the script falls back to
# $env:HALO_CLIENT_ID / $env:HALO_CLIENT_SECRET, and then to prompting you.
# ===========================================================================

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Tenant    = 'contoso'
$AuthUrl   = "https://$Tenant.haloitsm.com/auth/token"
$ApiBase   = "https://$Tenant.haloitsm.com/api"
$TimeOfDay = 'T12:00:00'
$Scope     = 'edit:tickets'

# --- the 30 tasks ---------------------------------------------------------
# Placeholder cast, so the examples read properly:
#   Alex Morgan    - facilities/IT lead, owns most of this
#   Jordan Blake   - infrastructure, owns the firewall work and ticket chasing
#   Sam Rivers     - facilities, holds the site history
#   Chris Doyle    - budget holder, signs off spend
#   Pat Ellis, Dana Webb - management, quote approval
#   Robin Hale     - finance
#   Casey Nolan    - service desk, dealt with the installer by phone
#   Northwind Renewables - PV installer and monitoring provider
#                    (Riley Stone and Morgan Pike; Drew Kelly has left)
#   Fabrikam Energy    - electricity supplier and FIT administrator
#   Tailspin Networks  - distribution network operator (DNO)
#   Woodgrove Safety   - outsourced health and safety adviser
#   Litware Purchasing - energy purchasing broker
#   Site A / B / C / D - the sites in question
$Records = @(
    [pscustomobject]@{ id='A1'; summary='Create Facilities SharePoint site and solar document library'; workstream='A - Documentation and governance'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-16'; target='2026-09-18'; days=3; deps='-'; description='Create a Facilities Team site with a Solar PV library (01 Installation records, 02 Assessments and inspection reports, 03 Quotes and orders, 04 Service agreements, 05 Monitoring, 06 Electricity supply and FIT, 07 Correspondence) plus Health and Safety and Maintenance and PPM at top level. Alex owner; Sam, Jordan and Chris members; Pat and Dana read-only if needed. Enable versioning, surface in the Facilities Teams channel, apply standard retention/sensitivity label.'; donewhen='Site live, structure created, Sam and Jordan confirm access.' }
    [pscustomobject]@{ id='A2'; summary='Circulate the 15 September meeting notes'; workstream='A - Documentation and governance'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-17'; target='2026-09-17'; days=1; deps='-'; description='Send the notes to Jordan, Sam and Chris with the action table intact and a line confirming the actions are being loaded into Halo. File a copy in Solar PV > 07 Correspondence.'; donewhen='Notes sent and filed.' }
    [pscustomobject]@{ id='A3'; summary='Collate all existing solar documentation into the new library'; workstream='A - Documentation and governance'; owner='All (Alex to chase)'; agentkey='Alex'; priority='High'; start='2026-09-21'; target='2026-10-02'; days=10; deps='A1'; description='Everyone dumps emails, reports, quotes and installation records into the library - no tidying first. Search mailboxes for the installer name, solar, PV, feed-in, MPAN and the supplier name. Drag threads in from Outlook or forward to a dedicated mailbox and file from there. Alex to include the supplier and DNO correspondence. Then sort into the folder structure and flag gaps, particularly the Site B quote/approval and the Site C inspection report.'; donewhen='Each attendee confirms nothing relevant remains outside the library.' }
    [pscustomobject]@{ id='A4'; summary='Build a solar PV asset register'; workstream='A - Documentation and governance'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-10-05'; target='2026-10-16'; days=10; deps='A3, B1'; description='Per site capture: panel count, array/string layout, inverter make/model/serial, commissioning date, installer, import and export MPANs, FIT registration number, monitoring hardware and install date, known faults, warranty expiry. Build from the installer''s installation data (Morgan Pike supplied the original data), their reports and the supplier and DNO correspondence. Hold as a spreadsheet in 01 Installation records; consider loading assets into Halo CMDB so tickets can be raised against them.'; donewhen='Register complete per site with gaps listed as open questions.' }
    [pscustomobject]@{ id='B1'; summary='Send the installer a status recap and request all documentation'; workstream='B - Installer engagement and outstanding works'; owner='Alex'; agentkey='Alex'; priority='Critical'; start='2026-09-16'; target='2026-09-17'; days=2; deps='-'; description='To Riley Stone and Morgan Pike at Northwind Renewables (note Drew Kelly has left, so nothing sits in their name). Structure: our understanding per site, then a numbered list of what we believe is outstanding, then the request. State: Site A monitoring installed and panels repaired; Site B unconfirmed, quote believed approved by Dana, 13 August install date referenced but unverified; Site C kit installed after the 20 August visit (reported 26 August) but monitoring not functional, no update since. Request all installation records, assessments and inspection reports, quotes and order confirmations, commissioning certificates, warranty documentation and their view of what remains outstanding per site. Ask them to confirm the service agreement position and the monitoring login process, name a single point of contact, and confirm Tuesday''s visit. Keep it factual - a statement of position and a request.'; donewhen='Email sent and the installer replies with a per-site outstanding list.' }
    [pscustomobject]@{ id='B2'; summary='Ask Casey what they handled with the installer by phone'; workstream='B - Installer engagement and outstanding works'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-16'; target='2026-09-17'; days=2; deps='-'; description='Sam raised a ticket forwarding the Site B and Site A install and Casey appears to have dealt with the installer by phone; nothing is written down. Get from Casey: which sites, what was agreed, any dates committed, and the original ticket reference. Record against the Halo ticket and in SharePoint, and feed anything material into B1 before it goes out if timing allows.'; donewhen='Casey''s account recorded against the ticket.' }
    [pscustomobject]@{ id='B3'; summary='Support the installer site visit - Tue 22 September 10:30 (firewall access)'; workstream='B - Installer engagement and outstanding works'; owner='Jordan'; agentkey='Jordan'; priority='Critical'; start='2026-09-17'; target='2026-09-22'; days=4; deps='-'; description='Ask the installer in advance what the monitoring kit needs - outbound destinations, ports, static internal IP or DHCP reservation. Raise the firewall change through the normal change process rather than ad hoc. Prefer a dedicated VLAN or IoT segment over the corporate LAN. Confirm the site and arrange an escort. After the visit record what was changed and whether access must remain permanent.'; donewhen='Access in place, visit completed, changes documented.' }
    [pscustomobject]@{ id='B4'; summary='Raise and track a Halo ticket per outstanding installer item'; workstream='B - Installer engagement and outstanding works'; owner='Jordan'; agentkey='Jordan'; priority='High'; start='2026-09-28'; target='2026-10-30'; days=25; deps='B1'; description='Once the installer replies, each outstanding item becomes its own ticket under this project with Site as a custom field and Northwind Renewables as supplier. Jordan owns chasing to completion and reports status at the facilities catch-up. Close items on evidence (report, photo, monitoring data), never on a verbal assurance.'; donewhen='Every item on the installer''s list has a ticket and the tickets are being worked.' }
    [pscustomobject]@{ id='B5'; summary='Confirm the definitive list of PV sites'; workstream='B - Installer engagement and outstanding works'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-21'; target='2026-10-02'; days=10; deps='B1'; description='The notes contradict themselves: the installer is recorded as assessing Site C, Site D and Site A, but the site status table covers Site A, Site B and Site C. Resolve with the installer and Sam whether Site D was assessed and is missing from tracking, or was named in error. Also confirm whether any other site has PV that has never been assessed. Monitoring logins, service agreement scope and MPAN checks all depend on this.'; donewhen='Confirmed site list with Site D''s status explicitly stated.' }
    [pscustomobject]@{ id='B6'; summary='Confirm Site B status - quote approval and 13 August install'; workstream='B - Installer engagement and outstanding works'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-21'; target='2026-10-02'; days=10; deps='B1, B2'; description='Three unverified facts: a quote went to Pat Ellis and Dana, Dana is believed to have approved it, and a 13 August install date was referenced. Ask Dana directly and get approval in writing. Ask finance/purchasing whether a PO was raised and whether the installer has invoiced - an invoice is the strongest evidence the work happened. Ask the installer to confirm the install took place and supply the commissioning record. If it did not happen it goes onto the outstanding list.'; donewhen='Site B status is a fact not a belief, and recorded in the asset register.' }
    [pscustomobject]@{ id='B7'; summary='Resolve the Site C monitoring fault'; workstream='B - Installer engagement and outstanding works'; owner='Alex (Jordan to ticket)'; agentkey='Alex'; priority='High'; start='2026-09-21'; target='2026-10-02'; days=10; deps='B1'; description='Kit installed after the 20 August visit (reported 26 August) but monitoring not functional. The installer was to confirm whether it can be fixed by phone or needs a return visit - no update received. Chase for a decision and book a date rather than leaving it open. Check first whether the fault is network-side on our part - the same firewall and segmentation questions as B3 apply at Site C.'; donewhen='Site C monitoring reporting data, or a return visit booked with a date.' }
    [pscustomobject]@{ id='B8'; summary='Put a service agreement in place with the installer'; workstream='B - Installer engagement and outstanding works'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-10-05'; target='2026-10-30'; days=20; deps='B1, B5'; description='Agreed this will otherwise be forgotten. Must cover: scope (panels, inverters, monitoring hardware and platform) at every confirmed site; servicing frequency and what an annual service includes (visual inspection, string testing, inverter check, thermal imaging if offered); fault response times and the route for raising them; what is included versus chargeable (parts, call-outs, panel replacement); monitoring platform access, number of logins and data ownership; term, price, notice period and annual uplift. Confirm no overlap with the twice-yearly panel clean already arranged with the window cleaners. Installer to quote, Alex to review, Chris to sign off spend, then file in 04 Service agreements.'; donewhen='Signed agreement filed and renewal date diarised.' }
    [pscustomobject]@{ id='C1'; summary='Request monitoring logins for Alex, Jordan, Sam and Chris'; workstream='C - Monitoring'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-21'; target='2026-10-02'; days=10; deps='B1'; description='Pat and Dana do not need access. Use company email addresses, not shared credentials. Ask the installer whether the platform supports SSO or MFA and what access levels exist (read-only versus configuration). Record platform name, URL and account owner in the asset register and add the accounts to the leavers process so they are not orphaned.'; donewhen='All four can log in and see live data.' }
    [pscustomobject]@{ id='C2'; summary='Confirm what the monitoring covers - generation only or import/export'; workstream='C - Monitoring'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-10-05'; target='2026-10-09'; days=5; deps='C1'; description='Assumed to be solar generation only, not import/export - needs confirming. If generation only, establish what would be needed to see import and export (separate meter, or supplier data), because without it the BESS case cannot be built and the FIT position cannot be validated.'; donewhen='Scope confirmed in writing and any gap to import/export visibility costed.' }
    [pscustomobject]@{ id='C3'; summary='Book a monitoring walkthrough from the installer'; workstream='C - Monitoring'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-10-05'; target='2026-10-16'; days=10; deps='C1'; description='We expect to be shown how it works once live. Book a short session (screen share is fine) covering navigation, string-level views, alerting and thresholds, data export, and who to call when something looks wrong. Record it and write it up into the Halo Knowledgebase so it is not knowledge held by one person.'; donewhen='Session held and knowledgebase article published.' }
    [pscustomobject]@{ id='C4'; summary='Baseline string-level performance review'; workstream='C - Monitoring'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-10-19'; target='2026-11-06'; days=15; deps='C1, C3, B7'; description='There were known faults with some strings and panels at the time of the site visits. Once monitoring is live at all sites take a baseline: which strings are underperforming, by how much, against what expected output. Cross-reference the installer''s inspection reports to confirm known faults were actually repaired and raise tickets for anything outstanding. Set a recurring review, monthly to begin with, so degradation is spotted rather than discovered.'; donewhen='Baseline documented, discrepancies ticketed, recurring review scheduled.' }
    [pscustomobject]@{ id='C5'; summary='Record the twice-yearly panel clean as a recurring task'; workstream='C - Monitoring'; owner='Alex'; agentkey='Alex'; priority='Low'; start='2026-09-21'; target='2026-10-02'; days=10; deps='-'; description='Already arranged with the window cleaners but currently held in one person''s head. Set up as a recurring Halo task or PPM entry with contractor, sites covered, frequency and expected months. Note that output should be compared before and after the clean once monitoring is live, as evidence of value.'; donewhen='Recurring task created with the next two dates set.' }
    [pscustomobject]@{ id='D1'; summary='Confirm the electricity supplier for import and export across all sites'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex'; agentkey='Alex'; priority='Critical'; start='2026-09-16'; target='2026-09-25'; days=8; deps='-'; description='Foundation for the whole workstream. Go via Robin in finance so we get all sites rather than just Site A. Ask per site for: supplier name, account number, import and export MPANs, contract end date, current invoices, and which entity the account sits under. Also ask whether any site is on a broker arrangement. Build a single supply schedule and file it in 06 Electricity supply and FIT.'; donewhen='Supply schedule complete for every site.' }
    [pscustomobject]@{ id='D2'; summary='Answer the supplier''s export MPAN query'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-28'; target='2026-11-06'; days=30; deps='D1, D7'; description='Fabrikam Energy ask whether we still have an export MPAN on the account; if not we need to deappoint the meter operator and data collector. Confirm from the supply schedule and the meter itself if necessary. If there is no export MPAN establish why - almost certainly the meter replacement, the same root cause suspected under D7. Do not deappoint until D7 is understood or we may remove what is needed to reinstate export. If deappointment is right, ask the supplier to confirm in writing what it means for the FIT and whether it is reversible.'; donewhen='Query answered in writing and deappointment taken or consciously deferred.' }
    [pscustomobject]@{ id='D3'; summary='Explain the disconnected import MPAN on FIT 0000 00000000'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-28'; target='2026-10-16'; days=15; deps='D1'; description='The supplier reports the import MPAN linked to FIT 0000 00000000 as disconnected and wants an explanation; suspected to be our old MPAN following a meter replacement. Identify the current MPAN for that site from the supply schedule and compare against the FIT record. Get the meter replacement date and the old/new MPAN pair from the supplier or meter operator. Reply factually with what the meter change was, when, and the correct current MPAN, and ask them to update the FIT registration.'; donewhen='The supplier confirms the FIT record is updated to the live MPAN.' }
    [pscustomobject]@{ id='D4'; summary='Recover held feed-in tariff payments and the signed FIT plan'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-11-09'; target='2026-11-27'; days=15; deps='D2, D3'; description='The supplier has updated the account and changed generator details, with a signed feed-in tariff plan to follow, but is holding our FIT payments pending the queries above. Chase the signed plan, confirm what period the held payments cover and how much, get a release date, and make sure finance know it is coming. Agree the ongoing reading submission process so payments do not stall again.'; donewhen='Plan signed and filed, held payments released, submission process documented.' }
    [pscustomobject]@{ id='D5'; summary='Reconcile the unused supplier purchase account'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex with finance'; agentkey='Alex'; priority='Low'; start='2026-09-28'; target='2026-10-16'; days=15; deps='D1'; description='There is a supplier purchase account (reference E00) registered to an out-of-area address against which we have bought nothing. Establish with finance whether it relates to our supply at all or is a legacy/duplicate. If it is not ours, have it closed so it stops muddying the picture.'; donewhen='Account identified and either linked to a real supply or closed.' }
    [pscustomobject]@{ id='D6'; summary='Account for missing electricity invoices since June 2024'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex with finance'; agentkey='Alex'; priority='High'; start='2026-09-28'; target='2026-10-16'; days=15; deps='D1'; description='Last invoice on record is June 2024. Either we have been billed elsewhere since (an unrecorded supplier change) or we have not been billed at all, which would mean a catch-up bill is possible. Ask Robin to check the ledger for all electricity spend since June 2024 by supplier and reconcile against the supply schedule. Flag any unbilled exposure to Chris early.'; donewhen='Every site''s electricity billing accounted for from June 2024 to date.' }
    [pscustomobject]@{ id='D7'; summary='Resolve the zero feed-in tariff and export permission with the DNO'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex'; agentkey='Alex'; priority='Critical'; start='2026-09-28'; target='2026-10-30'; days=25; deps='D1'; description='Tailspin Networks have come back saying our feed-in tariff is zero, suggesting we are not permitted to export at all; suspicion is the export registration was not carried over at the meter replacement. Contact the DNO quoting site MPANs and the FIT reference. Ask specifically whether there is a valid G98/G99 connection agreement per generating site and whether export is permitted or limited to zero. If export was never re-registered, establish what is needed to reinstate it and whether it is a form or a full application. Get the answer in writing - this determines whether we have been generating without being able to sell anything back. Quantify the loss once known, for Chris.'; donewhen='Export position confirmed in writing and a reinstatement route agreed if it is wrong.' }
    [pscustomobject]@{ id='D8'; summary='Check the purchasing broker for supply and FIT history'; workstream='D - Electricity supply and feed-in tariff'; owner='Alex'; agentkey='Alex'; priority='Low'; start='2026-09-16'; target='2026-09-25'; days=8; deps='-'; description='Litware Purchasing may hold useful knowledge on the supply arrangements and the history behind the FIT. Worth a call before reconstructing it from scratch - run in parallel with D1.'; donewhen='Contacted and anything useful fed into the supply schedule.' }
    [pscustomobject]@{ id='E1'; summary='Confirm the insurance position on overnight battery charging'; workstream='E - BESS (separate workstream)'; owner='Alex'; agentkey='Alex'; priority='High'; start='2026-09-28'; target='2026-10-23'; days=20; deps='-'; description='Chris has indicated there may be an insurance reason we cannot charge batteries overnight, suspected to relate to EV charging rather than static storage. Ask the broker specifically: does the policy restrict overnight charging, and does that restriction apply to static battery storage as well as EVs? Ask what conditions would make a static installation acceptable - separate container, distance from buildings, suppression, monitoring. Get it in writing. This gates the rest of the workstream, so do it first.'; donewhen='Broker''s written position held and its scope (EV versus static) clear.' }
    [pscustomobject]@{ id='E2'; summary='Scope BESS options with the installer and one other supplier'; workstream='E - BESS (separate workstream)'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-10-26'; target='2026-11-20'; days=20; deps='E1, C2'; description='Northwind Renewables are the agreed starting point. Ask for sizing recommendations per site based on our actual consumption profile, chemistry options (LFP rather than NMC given the fire discussion), containerised versus indoor, indicative cost, and what comparable sites they have done. Reference point discussed: a nearby retail park uses contactors to disable overnight charging and required a full sprinkler system - ask whether that is typical or specific to an indoor installation. Get a second quote so we are not single-sourced on a capital item.'; donewhen='At least two indicative proposals held.' }
    [pscustomobject]@{ id='E3'; summary='Health and safety and fire assessment with the H and S adviser'; workstream='E - BESS (separate workstream)'; owner='Alex with Woodgrove Safety'; agentkey='Alex'; priority='Medium'; start='2026-11-23'; target='2026-12-11'; days=15; deps='E2'; description='Woodgrove Safety are our virtual health and safety officer - bring them in before any commitment. Points to put to them: lithium cells release oxygen during thermal runaway so a fire is self-sustaining and cannot be smothered; the working assumption is a separate metal container sited away from buildings; what separation distance, suppression, ventilation and emergency procedure would be required; and what the fire service would want to know. Have them confirm in writing what a compliant installation looks like at our sites.'; donewhen='Written requirements held and reflected in the proposals.' }
    [pscustomobject]@{ id='E4'; summary='Build the BESS commercial case and propose a single-site trial'; workstream='E - BESS (separate workstream)'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-12-14'; target='2027-01-29'; days=32; deps='E1, E2, E3, D1, D7'; description='The case is: charge from the grid overnight at the cheaper rate and from solar by day, then run the premises off the battery at peak rate - reported to be common at comparable sites in our sector. Get half-hourly consumption data per site from the supplier, without which sizing and saving are guesswork. Model the saving from the day/night differential against capital cost and expected cycle life. Factor in the export position - if we cannot export, the case rests on self-consumption and peak avoidance, which changes the sizing. Include the health and safety and siting costs, which are not trivial if a container and suppression are needed. Recommend one site for a trial with the reasoning and take it to Chris with a payback figure.'; donewhen='Paper with a payback figure and recommended trial site submitted to Chris.' }
    [pscustomobject]@{ id='F1'; summary='Brief Morgan once the information is collated'; workstream='F - Closeout'; owner='Alex'; agentkey='Alex'; priority='Medium'; start='2026-11-02'; target='2026-11-06'; days=5; deps='A4, B1, D1, D7'; description='Agreed that once all the information is collated and understood, Morgan is to be briefed. CHECK BEFORE CREATING: the notes name Morgan Pike at the installer, but the sense of the action reads as briefing someone internally. Confirm which Morgan is meant and adjust dependencies - if it is Morgan Pike this is an installer status review and belongs in workstream B.'; donewhen='Briefing delivered and any actions arising captured.' }
)

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

$clientIdValue     = Resolve-Credential -Inline $ClientId     -EnvName 'HALO_CLIENT_ID'     -Prompt 'Halo Client ID'
$clientSecretValue = Resolve-Credential -Inline $ClientSecret -EnvName 'HALO_CLIENT_SECRET' -Prompt 'Halo Client Secret'

if ([string]::IsNullOrWhiteSpace($clientIdValue) -or [string]::IsNullOrWhiteSpace($clientSecretValue)) {
    throw 'No credentials supplied - stopping.'
}

$work = if ($OnlyIds.Count -gt 0) { $Records | Where-Object { $OnlyIds -contains $_.id } } else { $Records }
if (-not $work) { throw 'Nothing to do - check $OnlyIds.' }

# --- token ----------------------------------------------------------------
function Get-WebErrorBody {
    # A 400 from Halo carries a JSON body naming the real problem
    # (invalid_client, invalid_scope, unsupported_grant_type...). PowerShell
    # hides it behind a generic message, so dig it out. Works on 5.1 and 7.
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return $ErrorRecord.ErrorDetails.Message
    }
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($null -eq $resp) { return '' }
        $stream = $resp.GetResponseStream()
        $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream)
        return $reader.ReadToEnd()
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
    $lines = @(
        "Workstream:  $($Rec.workstream)"
        "Owner:       $($Rec.owner)"
        "Priority:    $($Rec.priority)"
        "Planned:     $($Rec.start) to $($Rec.target)  ($($Rec.days) working days)"
        "Depends on:  $($Rec.deps)"
        ''
        $Rec.description
        ''
        "Done when: $($Rec.donewhen)"
        ''
        'Source: Facilities / Solar PV meeting notes, 15 September 2026.'
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
        # Match on whatever keys are in $AgentMap, so replacing the cast in
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
            Write-Host '  No agent name matched the $AgentMap keys. Rename the keys to match your' -ForegroundColor DarkYellow
            Write-Host '  agents, or list them all with the /api/Agent endpoint.' -ForegroundColor DarkYellow
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

Write-Host ''
Write-Host "  Halo project build - $mode" -ForegroundColor $col
Write-Host ''
Write-Host "  Project: $ProjectSummary"
Write-Host "           $ProjectStart -> $ProjectTarget"
Write-Host "  Tasks:   $($work.Count)"
Write-Host ''
Write-Host ('  {0,-5}{1,-4}{2,-26}{3,-10}{4}' -f 'ID','WS','Dates','Owner','Summary')
Write-Host ('  ' + ('-' * 118)) -ForegroundColor DarkGray

foreach ($rec in $work) {
    $name = $rec.summary
    if ($name.Length -gt 52) { $name = $name.Substring(0, 52) }
    Write-Host ('  {0,-5}{1,-4}{2,-26}{3,-10}{4}' -f `
        $rec.id, $rec.workstream.Substring(0,1), "$($rec.start) -> $($rec.target)", $rec.agentkey, $name)
}
Write-Host ('  ' + ('-' * 118)) -ForegroundColor DarkGray

if (-not $Apply) {
    Write-Host ''
    Write-Host "  Would create 1 project and $($work.Count) tasks. Nothing has been written." -ForegroundColor Yellow
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
    $projFields = @{
        tickettype_id = $ProjectTypeId
        summary       = $ProjectSummary
        details       = @(
            'Solar PV estate: monitoring, electricity supply, feed-in tariff and BESS.'
            'Raised from the Facilities / Solar PV meeting of 15 September 2026.'
            'Attendees: Alex Morgan, Jordan Blake, Sam Rivers.'
            ''
            'Six workstreams:'
            '  A  Documentation and governance'
            '  B  Installer engagement and outstanding works'
            '  C  Monitoring'
            '  D  Electricity supply and feed-in tariff'
            '  E  BESS (separate workstream)'
            '  F  Closeout'
        ) -join "`r`n"
        startdate     = "$ProjectStart$TimeOfDay"
        targetdate    = "$ProjectTarget$TimeOfDay"
    }
    if ($HaloClientId -gt 0) { $projFields.client_id = $HaloClientId }
    if ($HaloSiteId   -gt 0) { $projFields.site_id   = $HaloSiteId }
    if ($AgentMap['Alex'] -gt 0) { $projFields.agent_id = $AgentMap['Alex'] }

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
        Write-Host ("  {0,-5} skipped - already on the project" -f $rec.id) -ForegroundColor DarkGray
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
        Write-Host ("  {0,-5} created  id {1,-8} {2}" -f $rec.id, $newId, $summary)
        $created += [pscustomobject]@{ Rec = $rec; Id = $newId; Summary = $summary }
    }
    catch {
        $b = Get-WebErrorBody $_
        $detail = if ($b) { " - $b" } else { '' }
        Write-Host ("  {0,-5} FAILED: {1}{2}" -f $rec.id, $_.Exception.Message, $detail) -ForegroundColor Red
        $failed += [pscustomobject]@{ Id = $rec.id; Error = "$($_.Exception.Message)$detail" }
        continue
    }

    # Canary: check the first task actually attached to the project before
    # creating the other 29. A dry run cannot test this, so it is tested here.
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
            Write-Host "        link confirmed - tasks are attaching to project $projectId" -ForegroundColor DarkGray
        }
        catch {
            Write-Host '        (could not verify the project link - continuing)' -ForegroundColor DarkYellow
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
    Write-Host '  calendar or an SLA to the ticket type. Apply-HaloDates-ISE.ps1 will put' -ForegroundColor Red
    Write-Host '  them back if you add the new ids to its $Records list.' -ForegroundColor Red
}
else {
    Write-Host "  All $($created.Count) verified - dates and project link correct." -ForegroundColor Green
}
Write-Host ''
