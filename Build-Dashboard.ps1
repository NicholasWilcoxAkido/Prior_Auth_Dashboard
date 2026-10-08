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
$COL_TYPE   = 'Type'
$COL_STATUS = 'Status (Referral Coordinator)'
$COL_RC     = 'RC Assigned'
$COL_FDATE  = 'Bot Run Fail Date'
$COL_FREASON= 'Bot Run Fail Reason'

$present = $rows[0].PSObject.Properties.Name
$missing = @($COL_DATE, $COL_TYPE, $COL_STATUS, $COL_RC, $COL_FDATE, $COL_FREASON |
             Where-Object { $_ -notin $present })
if ($missing.Count) {
    throw "Source is missing required column(s): $($missing -join ', ')`nFound: $($present -join ', ')"
}

# ================================================================ SANITIZE
# Everything below this block only ever sees scrubbed values. PPI is destroyed
# here, at ingest, rather than filtered out later.
#
#   RC Assigned          -> 'svc-quickbase' | '(human)' | '(unassigned)'
#                           Coordinator names and user IDs are discarded; only
#                           the bot-or-not distinction is needed downstream.
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

foreach ($r in $rows) {
    $rc = "$($r.$COL_RC)"
    if ($rc -imatch 'svc-quickbase')           { $r.$COL_RC = 'svc-quickbase'; $sanBot++ }
    elseif ([string]::IsNullOrWhiteSpace($rc)) { $r.$COL_RC = '(unassigned)';  $sanBlank++ }
    else                                       { $r.$COL_RC = '(human)';       $sanHuman++ }

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

# A row-level copy with PPI removed. The original export is never modified -
# this is written alongside it so you can verify the scrub and, if you prefer,
# keep only the sanitized version. Still excluded from git by .gitignore.
if (-not $SkipSanitizedCsv) {
    $sanDir = Join-Path $root 'sanitized'
    if (-not (Test-Path $sanDir)) { New-Item -ItemType Directory -Force -Path $sanDir | Out-Null }
    $sanPath = Join-Path $sanDir ([IO.Path]::GetFileNameWithoutExtension($Source) + '.sanitized.csv')
    $rows |
        Select-Object $COL_DATE, $COL_TYPE, $COL_STATUS, $COL_RC, $COL_FDATE, $COL_FREASON |
        Export-Csv -LiteralPath $sanPath -NoTypeInformation -Encoding utf8
    Write-Host ("Sanitized copy: {0}" -f $sanPath) -ForegroundColor DarkGray
}

# ---------------------------------------------------------------- normalizers

# Type arrives with inconsistent casing (Appointment / DIAGNOSTIC / REFERRAL / Procedure).
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
$dateIdx   = [ordered]@{}
$cells     = @{}

$badDates = 0
$minDate = [datetime]::MaxValue
$maxDate = [datetime]::MinValue

function Get-Index {
    param($Map, [string] $Key)
    if (-not $Map.Contains($Key)) { $Map[$Key] = $Map.Count }
    return $Map[$Key]
}

foreach ($r in $rows) {
    # --- date
    $raw = ("$($r.$COL_DATE)").Trim()
    $dt = [datetime]::MinValue
    if (-not [datetime]::TryParse($raw, [ref]$dt)) { $badDates++; continue }
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
    $ti = Get-Index $typeIdx   (Normalize-Type   -Value "$($r.$COL_TYPE)")
    $si = Get-Index $statusIdx (Normalize-Status -Value "$($r.$COL_STATUS)")
    $ri = if ($reason) { Get-Index $reasonIdx $reason } else { -1 }

    $key = "$di|$ti|$si|$b|$ri"
    if ($cells.ContainsKey($key)) { $cells[$key]++ } else { $cells[$key] = 1 }
}

$kept = $rows.Count - $badDates
Write-Host "Rows aggregated: $kept  (dropped for unparsable date: $badDates)"
Write-Host "Distinct cells : $($cells.Count)"
if ($kept -le 0) { throw "No rows had a parsable '$COL_DATE'." }

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
$typeOrder   = Get-OrderedKeys $typeIdx   @('Appointment', 'Diagnostic', 'Procedure', 'Referral')
$statusOrder = Get-OrderedKeys $statusIdx @('Auth Approved', 'Auth Not Required', 'Auth Denied', 'Auth Blocked')

$typeRemap = @{};   for ($i = 0; $i -lt $typeOrder.Count;   $i++) { $typeRemap[$typeIdx[$typeOrder[$i]]]     = $i }
$statusRemap = @{}; for ($i = 0; $i -lt $statusOrder.Count; $i++) { $statusRemap[$statusIdx[$statusOrder[$i]]] = $i }

$facts = [System.Collections.Generic.List[object]]::new()
foreach ($k in $cells.Keys) {
    $p = $k -split '\|'
    $facts.Add(@(
        $dateRemap[[int]$p[0]],
        $typeRemap[[int]$p[1]],
        $statusRemap[[int]$p[2]],
        [int]$p[3],
        [int]$p[4],
        $cells[$k]
    ))
}
# Sort by the remapped (chronological) date index.
$facts = [System.Collections.Generic.List[object]](@($facts | Sort-Object { $_[0] }))

$payload = [ordered]@{
    generatedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    sourceFile   = [IO.Path]::GetFileName($Source)
    rowsRead     = $rows.Count
    rowsUsed     = $kept
    rowsDropped  = $badDates
    dateMin      = $minDate.ToString('yyyy-MM-dd')
    dateMax      = $maxDate.ToString('yyyy-MM-dd')
    buckets      = $BUCKETS
    dates        = $dateOrder
    types        = $typeOrder
    statuses     = $statusOrder
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
# Coordinator names must not survive anywhere in the published dimensions.
foreach ($n in $personNames) {
    if ($n.Length -lt 4) { continue }
    foreach ($label in (@($reasonIdx.Keys) + @($typeOrder) + @($statusOrder))) {
        if ($label -imatch ('\b' + [regex]::Escape($n) + '\b')) { $leaks += "coordinator name '$n' in '$label'" }
    }
}
if ($leaks.Count) {
    Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "PUBLISH GATE FAILED - output deleted, nothing is safe to push." -ForegroundColor Red
    $leaks | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    throw "Potential identifiers survived normalization. Add a rule to `$REASON_RULES for the pattern(s) above, then re-run."
}
Write-Host "Publish gate: no identifier-like text in $($reasonIdx.Count) reason categories." -ForegroundColor DarkGray

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
}
if ($repoLeaks.Count) {
    Write-Host ""
    Write-Host "REPO GATE FAILED - a publishable file contains real export data." -ForegroundColor Red
    $repoLeaks | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    throw "Replace the value(s) above with a described shape (e.g. '<11-digit number>'), then re-run. Never paste raw error text into a published file."
}
Write-Host ("Repo gate: {0} files clean against {1} export identifiers / {2} name forms." -f $repoFiles.Count, $exportIds.Count, $nameParts.Count) -ForegroundColor DarkGray

# ---------------------------------------------------------------- console summary
$apptIdx = [array]::IndexOf($typeOrder, 'Appointment')
$tot = 0; $ok = 0; $retry = 0; $fail = 0
foreach ($f in $facts) {
    if ($f[1] -ne $apptIdx) { continue }
    $tot += $f[5]
    switch ($f[3]) { 0 { $ok += $f[5] } 1 { $retry += $f[5] } 2 { $fail += $f[5] } }
}
$touched = $ok + $retry + $fail
$success = $ok + $retry

Write-Host ""
Write-Host "=== Schedule Auth Metrics (all dates, Type = Appointment) ===" -ForegroundColor Cyan
Write-Host ("  Total Schedule Auths            : {0:N0}" -f $tot)
Write-Host ("  Touched by RPA                  : {0:N0}" -f $touched)
Write-Host ("  Percent Touched by RPA          : {0:N2}%" -f (100 * $touched / [math]::Max($tot, 1)))
Write-Host ("  RPA Success Rate                : {0:N2}%" -f (100 * $success / [math]::Max($touched, 1)))
Write-Host ("    bot success            : {0:N0}" -f $ok)
Write-Host ("    bot success after retry: {0:N0}" -f $retry)
Write-Host ("    bot failed             : {0:N0}" -f $fail)
Write-Host ""
Write-Host ("Wrote {0} ({1:N1} KB)" -f $OutFile, ((Get-Item $OutFile).Length / 1KB)) -ForegroundColor Green
