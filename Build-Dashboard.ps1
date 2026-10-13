<#
.SYNOPSIS
    Builds dashboard/data.json for the Prior Auth Dashboard from an Excel or CSV export.

.DESCRIPTION
    Reads the Orders & Appointments export, normalizes the key fields, and emits
    PRE-AGGREGATED counts only. Row-level data is never written to the output, so
    member identifiers cannot reach the published site.

    Bot Run Fail Reason values are scrubbed of embedded identifiers (member IDs,
    auth numbers) before aggregation. Anything that still looks like it carries an
    identifier is bucketed as "Other / unclassified" rather than published verbatim.

.PARAMETER Source
    Path to the .xlsx or .csv export. Defaults to the newest
    Orders_and_Appointments_Dashboard.* in the script directory.

.PARAMETER OutFile
    Path to write. Defaults to docs/data.json beside this script (docs/ is the
    folder GitHub Pages publishes).

.EXAMPLE
    .\Build-Dashboard.ps1
    .\Build-Dashboard.ps1 -Source .\Orders_and_Appointments_Dashboard.xlsx
#>
[CmdletBinding()]
param(
    [string] $Source,
    [string] $OutFile,
    [switch] $SkipSanitizedCsv
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }

# ---------------------------------------------------------------- resolve source
if (-not $Source) {
    $cand = Get-ChildItem -Path $root -File |
            Where-Object { $_.Name -match '^Orders_and_Appointments_Dashboard\.(xlsx|xls|csv)$' } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
    if (-not $cand) {
        throw "No Orders_and_Appointments_Dashboard.(xlsx|csv) found in $root. Pass -Source explicitly."
    }
    $Source = $cand.FullName
}
$Source = (Resolve-Path -LiteralPath $Source).Path
if (-not $OutFile) { $OutFile = Join-Path $root 'docs\data.json' }

Write-Host "Source : $Source"
Write-Host "Output : $OutFile"

# ---------------------------------------------------------------- load rows
# .xlsx is converted to a temp CSV via Excel COM; .csv is read directly.
function Import-SourceRows {
    param([string] $Path)

    if ($Path -match '\.csv$') {
        return Import-Csv -LiteralPath $Path
    }

    Write-Host "Converting workbook via Excel COM..."
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("oad_{0}.csv" -f ([guid]::NewGuid().ToString('N')))
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Path, 0, $true)   # read-only
        $wb.Worksheets.Item(1).Activate() | Out-Null
        $wb.SaveAs($tmp, 6)                            # 6 = xlCSV
        $wb.Close($false)
        return Import-Csv -LiteralPath $tmp
    }
    finally {
        if ($wb)    { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if (Test-Path $tmp) { Remove-Item $tmp -Force }
    }
}

$rows = @(Import-SourceRows -Path $Source)
Write-Host "Rows read: $($rows.Count)"
if ($rows.Count -eq 0) { throw "Source contained no data rows." }

# ---------------------------------------------------------------- validate schema
$COL_DATE   = 'Date Created (Roll-Up)'
$COL_CDATE  = 'Date Auth Completed Roll Up'
$COL_TYPE   = 'Type'
$COL_STATUS = 'Status (Referral Coordinator)'
$COL_RC     = 'RC Assigned'
$COL_FDATE  = 'Bot Run Fail Date'
$COL_FREASON= 'Bot Run Fail Reason'
# Arrived in V1.07. OPTIONAL on purpose: an export that predates it still has to
# build, so this column is read when present and skipped when not.
$COL_PORTAL = 'Insurance Portal'
# Also arrived in V1.07, also optional. This column holds provider NAMES - the
# same class of PPI as the coordinator names in RC Assigned - so it is read for
# exactly one purpose: deciding whether an Appointment row is a hard or an easy
# schedule. The name is converted to that one word during the sanitize stage and
# the original is overwritten in place, so nothing downstream of ingest can see
# it, let alone publish it. The roster that drives the decision lives in a
# gitignored file (see $HARD_ROSTER) rather than in this script, which is public.
$COL_PROVIDER = 'Encounter Provider'

# Every date on the dashboard buckets on THIS column. It is the completion date,
# so the dashboard reports throughput (what got finished in a period), not intake
# (what arrived). Changing this one line changes the meaning of every chart, and
# the two cohorts genuinely differ - a September completion report and a September
# creation report overlap by barely half. Don't switch it casually.
$COL_BASIS  = $COL_CDATE

$present = $rows[0].PSObject.Properties.Name
$missing = @($COL_DATE, $COL_CDATE, $COL_TYPE, $COL_STATUS, $COL_RC, $COL_FDATE, $COL_FREASON |
             Where-Object { $_ -notin $present })
if ($missing.Count) {
    throw "Source is missing required column(s): $($missing -join ', ')`nFound: $($present -join ', ')"
}

# The dashboard hides its PORTAL filter when this comes back false, so an older
# export degrades to the V1.05 dashboard instead of showing an empty control.
$hasPortal = $COL_PORTAL -in $present
Write-Host ("Insurance Portal: {0}" -f $(if ($hasPortal) { 'present' } else { "absent - PORTAL filter will be hidden" }))

# ---------------------------------------------------------------- hard roster
# The list of providers whose appointments count as "hard" to schedule. Kept in
# a gitignored sidecar file, NOT in this script: the patterns are provider names
# and this script is published. Missing or empty file degrades to the V1.07
# behaviour - one undivided 'Appointment' type - rather than failing the build,
# so a fresh clone of the public repo still produces a working dashboard.
$HARD_ROSTER = Join-Path $root 'providers.hard.txt'

# Hyphens become spaces and runs of whitespace collapse on BOTH sides of the
# comparison, so one roster entry covers both the hyphenated and the spaced
# spelling of a double-barrelled surname without needing two rules.
#
# (Writing the real pair out here as an example is what the repo gate below
# exists to stop, and it duly caught the first draft of this comment.)
function Normalize-ProviderText {
    param([string] $Value)
    return ((("$Value") -replace '-', ' ') -replace '\s+', ' ').Trim()
}

# Type arrives with inconsistent casing (Appointment / DIAGNOSTIC / REFERRAL / Procedure).
# Defined here rather than with the other normalizers because the sanitize stage
# below needs it to recognize Appointment rows.
function Normalize-Type {
    param([string] $Value)
    $v = ("$Value").Trim()
    if (-not $v) { return '(Blank)' }
    switch -Regex ($v) {
        '^appointment$' { return 'Appointment' }
        '^diagnostic$'  { return 'Diagnostic' }
        '^referral$'    { return 'Referral' }
        '^procedure$'   { return 'Procedure' }
        default         { return [cultureinfo]::InvariantCulture.TextInfo.ToTitleCase($v.ToLower()) }
    }
}

$hardPatterns = @()
if (Test-Path -LiteralPath $HARD_ROSTER) {
    $hardPatterns = @(Get-Content -LiteralPath $HARD_ROSTER |
        ForEach-Object { ($_ -split '#')[0] } |
        ForEach-Object { Normalize-ProviderText -Value $_ } |
        Where-Object { $_ } |
        Select-Object -Unique)
}
$hasProvider = $COL_PROVIDER -in $present
$canSplit    = $hasProvider -and $hardPatterns.Count -gt 0

# Deliberately prints the COUNT and never the patterns. This output gets pasted
# into README.md and into chat, and a roster of provider names is precisely what
# this design keeps out of both.
if ($canSplit) {
    Write-Host ("Encounter Provider: present - splitting Appointment on {0} roster pattern(s)" -f $hardPatterns.Count)
} elseif (-not $hasProvider) {
    Write-Host "Encounter Provider: absent - Appointment stays undivided" -ForegroundColor Yellow
} elseif (-not (Test-Path -LiteralPath $HARD_ROSTER)) {
    Write-Host ("Encounter Provider: present, but {0} is missing - Appointment stays undivided" -f (Split-Path -Leaf $HARD_ROSTER)) -ForegroundColor Yellow
} else {
    Write-Host ("Encounter Provider: present, but {0} lists no patterns - Appointment stays undivided" -f (Split-Path -Leaf $HARD_ROSTER)) -ForegroundColor Yellow
}

# ================================================================ SANITIZE
# Everything below this block only ever sees scrubbed values. PPI is destroyed
# here, at ingest, rather than filtered out later.
#
#   RC Assigned          -> 'svc-quickbase' | '(human)' | '(unassigned)'
#                           Coordinator names and user IDs are discarded; only
#                           the bot-or-not distinction is needed downstream.
#   Encounter Provider   -> 'Easy' | 'Hard' | '(n/a)'
#                           Provider names are discarded the same way; only the
#                           hard-roster verdict survives, and only on Appointment
#                           rows. '(n/a)' is every other Type, where the split
#                           has no meaning.
#   Bot Run Fail Reason  -> names and identifiers replaced with [name] / [id] / [date]

# Coordinator names are harvested from the file itself, so the redaction list
# stays current as staff change - there is no roster to maintain by hand.
$personNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($r in $rows) {
    $v = "$($r.$COL_RC)"
    if ($v -imatch 'svc-quickbase') { continue }
    $display = (($v -split '<')[0]).Trim()
    if (-not $display) { continue }
    [void]$personNames.Add($display)
    # Individual given/family names too, in case only one part appears in free text.
    foreach ($part in ($display -split '[\s,]+')) {
        if ($part.Length -ge 4) { [void]$personNames.Add($part) }
    }
}

function Protect-Text {
    param([string] $Value)

    $v = ("$Value") -replace '\s+', ' '
    $v = $v.Trim()
    if (-not $v) { return '' }

    # Dates first, so they do not get shredded into [id] fragments.
    $v = [regex]::Replace($v, '\b\d{1,2}/\d{1,2}/\d{2,4}\b', '[date]')

    # Explicit name cues.
    $v = [regex]::Replace($v, '(?i)\b(?:Dr|Mr|Mrs|Ms)\.?\s+[A-Z][\w\-'']+(?:\s+[A-Z][\w\-'']+)?', '[name]')
    $v = [regex]::Replace($v, '(?i)\b(?:patient|member|provider)\s+name\s*[:=]\s*\S+(?:\s+\S+)?', 'name: [name]')

    # Known coordinator names, whole-word only.
    foreach ($n in $personNames) {
        $v = [regex]::Replace($v, '\b' + [regex]::Escape($n) + '\b', '[name]', 'IgnoreCase')
    }

    # Identifiers. Order matters: strip digits glued to letters BEFORE the
    # whole-token rule, so 'word' + digit-run becomes 'word[id]' and not '[id]'.
    # Losing the word would break failure-category matching below.
    #
    # Examples are deliberately written as shapes, not real values - this file is
    # published to a public repo, so a pasted sample ID would defeat the point.
    $v = [regex]::Replace($v, '(?<=[A-Za-z])\d{3,}', '[id]')     # trailing digits glued to a word
    $v = [regex]::Replace($v, '\b\d{3,}', '[id]')                # standalone digit runs
    # Leftover mixed alphanumeric codes (both a letter and a digit, 6+ chars).
    $v = [regex]::Replace($v, '\b(?=[A-Za-z0-9]*\d)(?=[A-Za-z0-9]*[A-Za-z])[A-Za-z0-9]{6,}\b', '[id]')

    # Collapse runs of redactions produced by shredded codes.
    $v = [regex]::Replace($v, '(?:\[id\][\s\-:]*){2,}', '[id] ')
    return $v.Trim()
}

$sanBot = 0; $sanHuman = 0; $sanBlank = 0; $sanReasons = 0

# Every identifier-shaped token seen in the RAW reason text, captured here
# because sanitization overwrites the column in place - after this loop the real
# values are gone. The repo gate below needs them to prove no publishable file
# quotes one.
$exportIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

# Same idea for provider names: captured here, before the column is overwritten,
# so the gates below can prove no published file quotes one. Full display values
# only - NOT split into word parts the way coordinator names are. Provider values
# include practice-style names whose individual words ('Health', 'Family') occur
# in ordinary README prose, and a gate that cries wolf gets switched off.
$providerValues = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

# Per-pattern row counts, so the operator can confirm each roster entry hits the
# provider it was meant to hit. Counts only - the patterns are never printed.
$patternHits = @{}
foreach ($p in $hardPatterns) { $patternHits[$p] = 0 }
$sanEasy = 0; $sanHard = 0; $sanNa = 0

foreach ($r in $rows) {
    $rc = "$($r.$COL_RC)"
    if ($rc -imatch 'svc-quickbase')           { $r.$COL_RC = 'svc-quickbase'; $sanBot++ }
    elseif ([string]::IsNullOrWhiteSpace($rc)) { $r.$COL_RC = '(unassigned)';  $sanBlank++ }
    else                                       { $r.$COL_RC = '(human)';       $sanHuman++ }

    # --- Encounter Provider -> Easy/Hard verdict, name destroyed in place.
    # The verdict is only computed for Appointment rows; the split is defined
    # as a subdivision of that Type and means nothing for Diagnostic/Referral.
    if ($hasProvider) {
        $prov = Normalize-ProviderText -Value "$($r.$COL_PROVIDER)"
        if ($prov) { [void]$providerValues.Add($prov) }
        if ($canSplit -and (Normalize-Type -Value "$($r.$COL_TYPE)") -eq 'Appointment') {
            $verdict = 'Easy'
            foreach ($p in $hardPatterns) {
                if ($prov -imatch [regex]::Escape($p)) { $verdict = 'Hard'; $patternHits[$p]++ }
            }
            if ($verdict -eq 'Hard') { $sanHard++ } else { $sanEasy++ }
            $r.$COL_PROVIDER = $verdict
        } else {
            $sanNa++
            $r.$COL_PROVIDER = '(n/a)'
        }
    }

    $raw = "$($r.$COL_FREASON)"
    if (-not [string]::IsNullOrWhiteSpace($raw)) {
        foreach ($m in [regex]::Matches($raw, '\b(?=[A-Za-z0-9]*\d)[A-Za-z0-9]{5,}\b')) {
            [void]$exportIds.Add($m.Value)
        }
        $clean = Protect-Text -Value $raw
        if ($clean -ne (($raw -replace '\s+', ' ').Trim())) { $sanReasons++ }
        $r.$COL_FREASON = $clean
    }
}

Write-Host ("Sanitized: RC Assigned -> {0} bot / {1} human / {2} unassigned ({3} coordinator names discarded)" -f `
    $sanBot, $sanHuman, $sanBlank, $personNames.Count) -ForegroundColor DarkGray
Write-Host ("Sanitized: {0} fail reasons had names/identifiers redacted" -f $sanReasons) -ForegroundColor DarkGray
if ($hasProvider) {
    Write-Host ("Sanitized: Encounter Provider -> {0} hard / {1} easy / {2} n-a ({3} provider names discarded)" -f `
        $sanHard, $sanEasy, $sanNa, $providerValues.Count) -ForegroundColor DarkGray
    if ($canSplit) {
        # Pattern index, not the pattern. Line up with providers.hard.txt to read it.
        $i = 0
        foreach ($p in $hardPatterns) {
            $i++
            if ($patternHits[$p] -eq 0) {
                Write-Host ("  roster pattern #{0}: 0 rows - stale entry, or the name is spelled differently in this export" -f $i) -ForegroundColor Yellow
            } else {
                Write-Host ("  roster pattern #{0}: {1:N0} rows" -f $i, $patternHits[$p]) -ForegroundColor DarkGray
            }
        }
    }
}

# A row-level copy with PPI removed. The original export is never modified -
# this is written alongside it so you can verify the scrub and, if you prefer,
# keep only the sanitized version. Still excluded from git by .gitignore.
if (-not $SkipSanitizedCsv) {
    $sanDir = Join-Path $root 'sanitized'
    if (-not (Test-Path $sanDir)) { New-Item -ItemType Directory -Force -Path $sanDir | Out-Null }
    $sanPath = Join-Path $sanDir ([IO.Path]::GetFileNameWithoutExtension($Source) + '.sanitized.csv')
    # Insurance Portal joins the list as-is. Encounter Provider joins it only
    # under a renamed header: the column no longer holds a provider by this
    # point, it holds the Easy/Hard verdict, and a column labelled 'Encounter
    # Provider' containing the word 'Hard' invites someone to re-join it to a
    # name later.
    $sanCols = @($COL_DATE, $COL_CDATE, $COL_TYPE, $COL_STATUS, $COL_RC, $COL_FDATE, $COL_FREASON)
    if ($hasPortal) { $sanCols += $COL_PORTAL }
    if ($hasProvider) { $sanCols += @{ Name = 'Schedule Class'; Expression = { $_.$COL_PROVIDER } } }
    $rows |
        Select-Object $sanCols |
        Export-Csv -LiteralPath $sanPath -NoTypeInformation -Encoding utf8
    Write-Host ("Sanitized copy: {0}" -f $sanPath) -ForegroundColor DarkGray
}

# ---------------------------------------------------------------- normalizers
# Normalize-Type lives further up, above the sanitize stage: the Easy/Hard split
# has to ask "is this an Appointment?" while the provider name is still readable,
# and PowerShell needs the function defined before that point. One definition,
# used by both stages, so the casing rules cannot drift apart.

# The Type dimension as the dashboard publishes it. 'Appointment' subdivides into
# 'Easy Schedules' / 'Hard Schedules' using the verdict the sanitize stage left
# behind; every other Type passes through untouched. Both output labels are fixed
# literals, so no provider name can reach the published dimension even if the
# roster or the export changes shape. With no roster the verdict is '(n/a)' and
# this collapses back to plain 'Appointment'.
function Resolve-Type {
    param([string] $Type, [string] $Class)
    $t = Normalize-Type -Value $Type
    if ($t -ne 'Appointment') { return $t }
    switch ($Class) {
        'Hard'  { return 'Hard Schedules' }
        'Easy'  { return 'Easy Schedules' }
        default { return 'Appointment' }
    }
}

# Portal names are payer/vendor brands (IEHP, EzNet, Medpoint), so they are NOT
# title-cased the way Type is - that would turn IEHP into Iehp. Whitespace is
# collapsed and casing is left exactly as the export wrote it.
#
# A blank portal is a real answer, not missing data: ~32% of rows are blank and
# 24k of those are 'Auth Not Required', i.e. no portal was ever involved. Only 2
# of ~19.5k RPA-touched rows are blank, so this bucket is almost entirely manual
# work that needed no portal.
function Normalize-Portal {
    param([string] $Value)
    $v = (("$Value") -replace '\s+', ' ').Trim()
    if (-not $v) { return '(Blank)' }
    return $v
}

function Normalize-Status {
    param([string] $Value)
    $v = ("$Value").Trim()
    if (-not $v) { return '(Blank)' }
    switch -Regex ($v) {
        '^auth\s+approved$'     { return 'Auth Approved' }
        '^auth\s+not\s+required$' { return 'Auth Not Required' }
        '^auth\s+blocked$'      { return 'Auth Blocked' }
        '^auth\s+denied$'       { return 'Auth Denied' }
        default                 { return $v }
    }
}

# Collapse the raw fail strings into stable categories AND strip embedded
# member IDs / auth numbers. Order matters: FIRST MATCH WINS, so a broad rule
# placed early will starve the specific rules below it.
#
# When a new export pushes 'Other / unclassified' up the chart, add rules here -
# that bucket is the designated signal, not a problem to ignore. The V1.02
# export introduced ~10 new message shapes (marked NEW below).
$REASON_RULES = @(
    @{ Pattern = 'referral is already submitted';                 Label = 'Referral already submitted' }
    @{ Pattern = 'previous auth(orization)? on file';             Label = 'Previous authorization on file' }
    @{ Pattern = 'patient not found in referral status search';   Label = 'Patient not found in referral status search' }
    @{ Pattern = 'could not find the (user-interface \(ui\)|ui) element'; Label = 'UI element not found / invalid' }
    # NEW - an ambiguous selector is a different bot defect from a missing one,
    # so it gets its own category rather than folding into 'not found'.
    @{ Pattern = 'could not uniquely identify|multiple similar matches found'; Label = 'UI element ambiguous' }
    @{ Pattern = 'ui element is invalid|target element did not appear'; Label = 'UI element not found / invalid' }  # NEW
    @{ Pattern = 'cannot select item';                            Label = 'Cannot select item in list' }
    @{ Pattern = 'cannot bring the target application';           Label = 'Cannot focus target application' }
    @{ Pattern = 'windows session (is locked|was disconnected)';  Label = 'Windows session locked / disconnected' }
    @{ Pattern = 'type activity verification failed';             Label = 'Type activity verification failed' }
    @{ Pattern = 'input row count is not match';                  Label = 'Input row count mismatch' }
    @{ Pattern = 'due date expired';                              Label = 'Due date expired' }
    @{ Pattern = 'member id (is )?not (found|match)|data not available for this member id'; Label = 'Member ID not found / mismatch' }
    @{ Pattern = 'member is not eligible';                        Label = 'Member not eligible' }
    @{ Pattern = 'provider (not|is not|address|adress)|unable to find (the )?provider'; Label = 'Provider not found / not matched' }
    @{ Pattern = 'servicing provider';                            Label = 'Servicing provider issue' }
    @{ Pattern = 'supervising physician is empty';                Label = 'Supervising physician empty' }
    # NEW - sits AFTER the servicing-provider rule on purpose, so the ~250
    # 'Unable to find Servicing Provider ...' rows keep their own category
    # instead of being swallowed by this broader provider phrasing.
    @{ Pattern = 'unable to (select|find) (the )?(referral|referring) provider|unable to select provider|provider details not found|out-of-network provider|(referring|supervision) provider is empty'; Label = 'Provider not found / not matched' }
    @{ Pattern = 'facility\s*(location)?\s*not\s*match';          Label = 'Facility not matched' }
    @{ Pattern = 'document not signed';                           Label = 'Document not signed' }
    @{ Pattern = 'special(i)?ty\s+not\s+(present|match)|unable to find the special(i)?ty'; Label = 'Specialty not found / not matched' }
    @{ Pattern = 'no icd found|icd is empty|unable to find icd';  Label = 'ICD missing' }
    @{ Pattern = 'invalid cpt|no services were found that match the cpt'; Label = 'CPT not found / invalid' }
    @{ Pattern = 'activity timeout exceeded|timeout reached';     Label = 'Timeout' }
    @{ Pattern = 'utilization management workflow';               Label = 'Routed to UM workflow' }
    @{ Pattern = 'canceled iehp';                                 Label = 'Canceled in IEHP' }
    @{ Pattern = 'request got (denied|cancelled)|(is|got) cancelled'; Label = 'Request cancelled by payer' }  # NEW
    @{ Pattern = 'value does not fall within the expected range'; Label = 'Value out of expected range' }
    @{ Pattern = 'target element is disabled|button is not enabled|is not enabled'; Label = 'Control disabled / not enabled' }
    @{ Pattern = 'cannot communicate with the browser|uipath extension'; Label = 'Browser automation extension issue' }
    @{ Pattern = 'outside of screen bounds';                      Label = 'Element outside screen bounds' }
    @{ Pattern = 'element was not found|unable to find the searched element'; Label = 'UI element not found / invalid' }
    @{ Pattern = 'not recognized as a valid datetime';            Label = 'Invalid date value' }
    @{ Pattern = 'patient not found in capella';                  Label = 'Patient not found in Capella' }
    @{ Pattern = 'no authorization letter data';                  Label = 'No authorization letter data' }
    @{ Pattern = 'no reference number|referral reference not found'; Label = 'Reference number missing' }
    @{ Pattern = 'issue with optum portal';                       Label = 'Optum portal issue' }
    @{ Pattern = 'port(a|e)l exception';                          Label = 'Payer portal exception' }      # NEW
    @{ Pattern = 'unable to find servicing lab location';         Label = 'Servicing lab location not found' }
    @{ Pattern = 'exception from medpoint';                       Label = 'Medpoint exception' }
    # NEW - bot-side data problems, not payer problems. Worth separating because
    # the fix lives in the automation, not the portal.
    @{ Pattern = 'does not belong to table|contains no datarows|sequence contains no elements'; Label = 'Input data / schema error' }
    @{ Pattern = 'cannot access the file';                        Label = 'File locked by another process' }  # NEW
)

# The published label set is CLOSED: every value returned here is either one of
# the fixed labels in $REASON_RULES or the literal 'Other / unclassified'. Input
# text is never echoed through, so no name or identifier can reach the output
# even if a future export invents a failure message nobody has seen.
$REASON_OTHER = 'Other / unclassified'

function Normalize-Reason {
    param([string] $Value)

    $v = ("$Value") -replace '\s+', ' '
    $v = $v.Trim()
    if (-not $v) { return $null }

    foreach ($rule in $REASON_RULES) {
        if ($v -imatch $rule.Pattern) { return $rule.Label }
    }
    return $REASON_OTHER
}

# ---------------------------------------------------------------- aggregate
# Bucket semantics (per the agreed overlap rule: bot failed then succeeded = success)
#   0 bot_ok       : RC Assigned = svc-quickbase, no fail date   -> touched, success
#   1 bot_retry_ok : RC Assigned = svc-quickbase, HAS fail date  -> touched, success (retried)
#   2 bot_fail     : fail date, NOT svc-quickbase                -> touched, failure
#   3 human        : neither                                     -> not touched
$BUCKETS = @('bot_ok', 'bot_retry_ok', 'bot_fail', 'human')

$typeIdx   = [ordered]@{}
$statusIdx = [ordered]@{}
$reasonIdx = [ordered]@{}
$portalIdx = [ordered]@{}
$dateIdx   = [ordered]@{}
$cells     = @{}

$badDates = 0
$futureDates = 0
$minDate = [datetime]::MaxValue
$maxDate = [datetime]::MinValue

# A handful of rows carry a completion date past the day the export was pulled -
# 23 of them in V1.05, the furthest five months out. They are data-entry noise,
# but they are NOT harmless: the dashboard anchors its date presets to the latest
# date in the file, so one stray 2027 row silently drags the "14 Days" window into
# a stretch of calendar with nothing in it and the dashboard opens empty. Cut
# anything after the export's own pull date and report the count.
$dateCeiling = (Get-Item -LiteralPath $Source).LastWriteTime.Date

function Get-Index {
    param($Map, [string] $Key)
    if (-not $Map.Contains($Key)) { $Map[$Key] = $Map.Count }
    return $Map[$Key]
}

foreach ($r in $rows) {
    # --- date
    $raw = ("$($r.$COL_BASIS)").Trim()
    $dt = [datetime]::MinValue
    if (-not [datetime]::TryParse($raw, [ref]$dt)) { $badDates++; continue }
    if ($dt.Date -gt $dateCeiling) { $futureDates++; continue }
    if ($dt -lt $minDate) { $minDate = $dt }
    if ($dt -gt $maxDate) { $maxDate = $dt }
    $dKey = $dt.ToString('yyyy-MM-dd')

    # --- buckets
    $isBot  = ("$($r.$COL_RC)")    -imatch 'svc-quickbase'
    $hasFail= -not [string]::IsNullOrWhiteSpace("$($r.$COL_FDATE)")
    $b = if ($isBot -and $hasFail) { 1 } elseif ($isBot) { 0 } elseif ($hasFail) { 2 } else { 3 }

    # --- reason only carries meaning when the bot actually failed
    $reason = $null
    if ($hasFail) {
        $reason = Normalize-Reason -Value "$($r.$COL_FREASON)"
        if (-not $reason) { $reason = 'Other / unclassified' }
    }

    $di = Get-Index $dateIdx   $dKey
    $ti = Get-Index $typeIdx   (Resolve-Type -Type "$($r.$COL_TYPE)" -Class $(if ($hasProvider) { "$($r.$COL_PROVIDER)" } else { '' }))
    $si = Get-Index $statusIdx (Normalize-Status -Value "$($r.$COL_STATUS)")
    $ri = if ($reason) { Get-Index $reasonIdx $reason } else { -1 }
    $pi = if ($hasPortal) { Get-Index $portalIdx (Normalize-Portal -Value "$($r.$COL_PORTAL)") } else { -1 }

    $key = "$di|$ti|$si|$b|$ri|$pi"
    if ($cells.ContainsKey($key)) { $cells[$key]++ } else { $cells[$key] = 1 }
}

$kept = $rows.Count - $badDates - $futureDates
Write-Host "Date basis     : $COL_BASIS"
Write-Host "Rows aggregated: $kept  (dropped for unparsable date: $badDates, dated after the $($dateCeiling.ToString('yyyy-MM-dd')) pull: $futureDates)"
Write-Host "Distinct cells : $($cells.Count)"
if ($kept -le 0) { throw "No rows had a parsable '$COL_BASIS'." }

# --- remap dimension indices to a stable, meaningful published order.
# Dates MUST end up chronological: the client slices a date range by index.
$dateOrder = @($dateIdx.Keys | Sort-Object)
$dateRemap = @{}
for ($i = 0; $i -lt $dateOrder.Count; $i++) { $dateRemap[$dateIdx[$dateOrder[$i]]] = $i }

# Types/statuses get a fixed business order so chart categories don't shuffle
# between builds; '(Blank)' and anything unrecognized sort last.
function Get-OrderedKeys {
    param($Map, [string[]] $Preferred)
    $keys = @($Map.Keys)
    return @($keys | Sort-Object @{ Expression = {
        $i = [array]::IndexOf($Preferred, $_)
        if ($i -ge 0) { $i } else { $Preferred.Count }
    }}, @{ Expression = { $_ } })
}
# 'Appointment' stays in the preferred order even though the split replaces it:
# an export without the provider column still produces it, and it has to sort
# where it always did rather than falling to the bottom with the strays.
$typeOrder   = Get-OrderedKeys $typeIdx   @('Easy Schedules', 'Hard Schedules', 'Appointment', 'Diagnostic', 'Procedure', 'Referral')
$statusOrder = Get-OrderedKeys $statusIdx @('Auth Approved', 'Auth Not Required', 'Auth Denied', 'Auth Blocked')

# Portals get no hand-written preferred list: the roster is the payer's to
# change, and a new portal appearing in an export must sort somewhere sensible
# on its own. Alphabetical, with '(Blank)' forced to the end so the 'no portal
# involved' bucket never leads the checkbox list.
$portalOrder = @($portalIdx.Keys | Sort-Object `
    @{ Expression = { if ($_ -eq '(Blank)') { 1 } else { 0 } } }, `
    @{ Expression = { $_ } })

$typeRemap = @{};   for ($i = 0; $i -lt $typeOrder.Count;   $i++) { $typeRemap[$typeIdx[$typeOrder[$i]]]     = $i }
$statusRemap = @{}; for ($i = 0; $i -lt $statusOrder.Count; $i++) { $statusRemap[$statusIdx[$statusOrder[$i]]] = $i }
$portalRemap = @{}; for ($i = 0; $i -lt $portalOrder.Count; $i++) { $portalRemap[$portalIdx[$portalOrder[$i]]] = $i }

# Fact tuple: [date, type, status, bucket, reason, COUNT, portal]
#
# The portal index is APPENDED AFTER the count rather than slotted in with the
# other dimensions, which is the wrong shape semantically and deliberate anyway.
# GitHub Pages serves a stale data.json for up to ~10 minutes after a push, so
# the new index.html and the old data.json are guaranteed to meet. Keeping the
# count at index 5 means that pairing renders the previous dashboard exactly and
# just hides the PORTAL filter; moving the count to index 6 would have turned
# every number on the page into NaN for those minutes.
$facts = [System.Collections.Generic.List[object]]::new()
foreach ($k in $cells.Keys) {
    $p = $k -split '\|'
    $row = @(
        $dateRemap[[int]$p[0]],
        $typeRemap[[int]$p[1]],
        $statusRemap[[int]$p[2]],
        [int]$p[3],
        [int]$p[4],
        $cells[$k]
    )
    if ($hasPortal) { $row += $portalRemap[[int]$p[5]] }
    $facts.Add($row)
}
# Sort by the remapped (chronological) date index.
$facts = [System.Collections.Generic.List[object]](@($facts | Sort-Object { $_[0] }))

$payload = [ordered]@{
    generatedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    sourceFile   = [IO.Path]::GetFileName($Source)
    rowsRead     = $rows.Count
    rowsUsed     = $kept
    rowsDropped  = $badDates
    rowsFuture   = $futureDates
    dateBasis    = $COL_BASIS
    dateMin      = $minDate.ToString('yyyy-MM-dd')
    dateMax      = $maxDate.ToString('yyyy-MM-dd')
    buckets      = $BUCKETS
    dates        = $dateOrder
    types        = $typeOrder
    statuses     = $statusOrder
    # Always emitted. Empty means "this export has no Insurance Portal column",
    # which is the single condition the client tests to hide the PORTAL filter.
    portals      = @(if ($hasPortal) { $portalOrder } else { @() })
    reasons      = @($reasonIdx.Keys)
    facts        = $facts
}

$outDir = Split-Path -Parent $OutFile
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# Depth 4 is enough for the nested fact arrays and keeps the file compact.
$json = $payload | ConvertTo-Json -Depth 4 -Compress
[IO.File]::WriteAllText($OutFile, $json, (New-Object Text.UTF8Encoding $false))

# ---------------------------------------------------------------- publish gate
# Last line of defence before this file is pushed to a public site. The reason
# scrubber is pattern-based, so a new failure message in a future export could
# introduce an identifier the rules do not cover. Fail loudly rather than publish.
$allowedLabels = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($rule in $REASON_RULES) { [void]$allowedLabels.Add($rule.Label) }
[void]$allowedLabels.Add($REASON_OTHER)

$leaks = @()
# Primary check: the label set is closed, so anything unrecognized is a defect.
foreach ($label in $reasonIdx.Keys) {
    if (-not $allowedLabels.Contains($label)) { $leaks += "label outside the closed set: '$label'" }
}
# Belt and braces: assert no identifier-shaped text anywhere in the payload.
foreach ($label in $reasonIdx.Keys) {
    if ($label -match '\d{3,}') { $leaks += "digit run in label: '$label'" }
    elseif ($label -match '(?i)member\s*id\s*[:\-]\s*\S') { $leaks += "member id phrasing: '$label'" }
    elseif ($label -match '(?i)authorization number is') { $leaks += "auth number phrasing: '$label'" }
}
# Portal labels pass through from the export verbatim - unlike reasons, this
# label set is OPEN, because the payer roster is not ours to close. So the same
# identifier checks that guard the reason labels have to cover these too.
foreach ($label in $portalOrder) {
    if ($label -match '\d{3,}') { $leaks += "digit run in portal: '$label'" }
    elseif ($label.Length -gt 40) { $leaks += "portal label too long to be a brand name: '$label'" }
}

# Coordinator names must not survive anywhere in the published dimensions.
foreach ($n in $personNames) {
    if ($n.Length -lt 4) { continue }
    foreach ($label in (@($reasonIdx.Keys) + @($typeOrder) + @($statusOrder) + @($portalOrder))) {
        if ($label -imatch ('\b' + [regex]::Escape($n) + '\b')) { $leaks += "coordinator name '$n' in '$label'" }
    }
}
# Provider names likewise. Resolve-Type is built to emit fixed literals, so this
# can only fire if someone rewires it to pass a provider value through - which is
# exactly the mistake worth catching, since the Type dimension is published and
# the whole point of the Easy/Hard split is that the roster stays private.
foreach ($pv in $providerValues) {
    if ($pv.Length -lt 4) { continue }
    foreach ($label in (@($typeOrder) + @($statusOrder) + @($portalOrder) + @($reasonIdx.Keys))) {
        if ((Normalize-ProviderText -Value $label) -imatch [regex]::Escape($pv)) {
            $leaks += "provider name in published label '$label'"
        }
    }
}
if ($leaks.Count) {
    Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "PUBLISH GATE FAILED - output deleted, nothing is safe to push." -ForegroundColor Red
    $leaks | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    throw "Potential identifiers survived normalization. Add a rule to `$REASON_RULES for the pattern(s) above, then re-run."
}
Write-Host ("Publish gate: no identifier-like text in {0} reason categories / {1} portals." -f $reasonIdx.Count, $portalOrder.Count) -ForegroundColor DarkGray

# ------------------------------------------------- repo gate (documentation too)
# data.json is not the only thing that gets pushed. Documenting a new failure
# category is a natural moment to paste a real error message into README.md or
# into a comment here - and a real error message carries a real member ID. This
# scan compares every publishable file against the identifiers that actually
# exist in THIS export, so it cannot fire on ordinary prose or on the counts and
# pixel sizes that legitimately contain digit runs.
# $exportIds was captured during the sanitize stage, before the raw values were
# overwritten - harvesting it here would scan already-redacted text and produce
# an empty set, making this whole gate silently vacuous.
if ($exportIds.Count -eq 0) {
    throw "Repo gate has nothing to check - `$exportIds is empty. The capture in the sanitize stage is broken; fix it rather than shipping a gate that always passes."
}
# Single-word name parts only; multi-word forms are covered by their parts.
$nameParts = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($n in $personNames) {
    if ($n.Length -ge 4 -and $n -notmatch '[\s,]') { [void]$nameParts.Add($n) }
}

$repoFiles = @('README.md', 'Build-Dashboard.ps1', 'Serve-Dashboard.ps1', '.gitignore') |
    ForEach-Object { Join-Path $root $_ }
$repoFiles += $OutFile
$repoFiles += (Join-Path $root 'docs\index.html')

$repoLeaks = @()
foreach ($f in $repoFiles) {
    if (-not (Test-Path -LiteralPath $f)) { continue }
    $txt = [IO.File]::ReadAllText($f)
    $name = Split-Path -Leaf $f
    foreach ($m in [regex]::Matches($txt, '\b[A-Za-z0-9]{4,}\b')) {
        $tok = $m.Value
        if ($exportIds.Contains($tok)) { $repoLeaks += "${name}: export identifier '$tok'" }
        elseif ($nameParts.Contains($tok)) { $repoLeaks += "${name}: coordinator name '$tok'" }
    }
    # Provider names and roster patterns, matched as whole values against the
    # normalized text rather than token-by-token. Documenting the Easy/Hard split
    # is the obvious moment to paste in "the providers are X, Y, Z" as an example,
    # and that one sentence would publish the roster this design keeps private.
    $flat = Normalize-ProviderText -Value $txt
    foreach ($pv in $providerValues) {
        if ($pv.Length -lt 6) { continue }
        if ($flat -imatch ('\b' + [regex]::Escape($pv) + '\b')) { $repoLeaks += "${name}: provider name from the export" }
    }
    foreach ($p in $hardPatterns) {
        if ($p.Length -lt 4) { continue }
        if ($flat -imatch ('\b' + [regex]::Escape($p) + '\b')) { $repoLeaks += "${name}: hard-schedule roster pattern" }
    }
}
if ($repoLeaks.Count) {
    Write-Host ""
    Write-Host "REPO GATE FAILED - a publishable file contains real export data." -ForegroundColor Red
    $repoLeaks | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    throw "Replace the value(s) above with a described shape (e.g. '<11-digit number>'), then re-run. Never paste raw error text into a published file."
}
Write-Host ("Repo gate: {0} files clean against {1} export identifiers / {2} name forms / {3} provider names." -f `
    $repoFiles.Count, $exportIds.Count, $nameParts.Count, $providerValues.Count) -ForegroundColor DarkGray

# ---------------------------------------------------------------- console summary
# "Schedule auths" spans every label the Appointment type can appear under, so
# this total is comparable across builds whether or not the roster split it.
function Measure-Types {
    param([string[]] $Labels)
    $idx = @($Labels | ForEach-Object { [array]::IndexOf($typeOrder, $_) } | Where-Object { $_ -ge 0 })
    $m = [ordered]@{ total = 0; ok = 0; retry = 0; fail = 0 }
    if (-not $idx.Count) { return $m }
    foreach ($f in $facts) {
        if ($idx -notcontains $f[1]) { continue }
        $m.total += $f[5]
        switch ($f[3]) { 0 { $m.ok += $f[5] } 1 { $m.retry += $f[5] } 2 { $m.fail += $f[5] } }
    }
    return $m
}

$SCHED_TYPES = @('Easy Schedules', 'Hard Schedules', 'Appointment')
$m = Measure-Types -Labels $SCHED_TYPES
$touched = $m.ok + $m.retry + $m.fail
$success = $m.ok + $m.retry

Write-Host ""
Write-Host ("=== Schedule Auth Metrics (all dates, {0}) ===" -f $(if ($canSplit) { 'Easy + Hard Schedules' } else { 'Type = Appointment' })) -ForegroundColor Cyan
Write-Host ("  Total Schedule Auths            : {0:N0}" -f $m.total)
Write-Host ("  Touched by RPA                  : {0:N0}" -f $touched)
Write-Host ("  Percent Touched by RPA          : {0:N2}%" -f (100 * $touched / [math]::Max($m.total, 1)))
Write-Host ("  RPA Success Rate                : {0:N2}%" -f (100 * $success / [math]::Max($touched, 1)))
Write-Host ("    bot success            : {0:N0}" -f $m.ok)
Write-Host ("    bot success after retry: {0:N0}" -f $m.retry)
Write-Host ("    bot failed             : {0:N0}" -f $m.fail)

# The split is only worth having if the two halves behave differently, so print
# them side by side - this is the line that tells you whether "hard" is real.
if ($canSplit) {
    Write-Host ""
    Write-Host "  --- split ---" -ForegroundColor Cyan
    foreach ($lbl in @('Easy Schedules', 'Hard Schedules')) {
        $s = Measure-Types -Labels @($lbl)
        $st = $s.ok + $s.retry + $s.fail
        Write-Host ("  {0,-15}: {1,7:N0} auths | {2,6:N0} touched ({3,5:N2}%) | success {4,5:N2}%" -f `
            $lbl, $s.total, $st, (100 * $st / [math]::Max($s.total, 1)), (100 * ($s.ok + $s.retry) / [math]::Max($st, 1)))
    }
}
Write-Host ""
Write-Host ("Wrote {0} ({1:N1} KB)" -f $OutFile, ((Get-Item $OutFile).Length / 1KB)) -ForegroundColor Green
