<#
.SYNOPSIS
    Builds a single self-contained HTML dashboard with the data inlined.

.DESCRIPTION
    Emits ONE file that needs no web server, no data.json beside it and no
    passphrase. Intended for sharing as an attachment or from a network drive -
    including as a Notion file attachment, which downloads and opens locally
    (Notion does not render uploaded HTML inline; only an /embed of a hosted URL
    displays in the page).

    docs/index.html is the single template for BOTH delivery modes. This script
    only ever READS it, so running this does not affect the hosted GitHub Pages
    build in any way. The template picks its mode at runtime: the injected
    #pa-inline-data block below is what switches its loadPayload() from fetching
    data.json to reading the payload inline.

    All of Build-Dashboard.ps1 runs first, so ingest sanitization, the publish
    gate and the repo gate all still apply - this does not bypass a safeguard.

    THE OUTPUT IS PLAINTEXT. Anyone holding the file can read every figure in
    it. That is the trade this mode makes: it swaps the passphrase gate for
    whatever access control the place you put the file provides. Do not post it
    anywhere public. standalone/ is gitignored for exactly that reason.

.PARAMETER Source
    Path to the .xlsx or .csv export. Forwarded to Build-Dashboard.ps1, which
    applies its own default when this is omitted.

.PARAMETER OutFile
    Path to write. Defaults to standalone/Prior_Auth_Dashboard_<dateMax>.html.
    Refuses any path inside docs/.

.EXAMPLE
    .\Build-Standalone.ps1 -Source "Orders_and_AppointmentsV1.08.csv"
#>
[CmdletBinding()]
param(
    [string] $Source,
    [string] $OutFile
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }

$template = Join-Path $root 'docs\index.html'
$builder  = Join-Path $root 'Build-Dashboard.ps1'
foreach ($p in @($template, $builder)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "Required file not found: $p" }
}

# GitHub Pages publishes docs/ wholesale, so a plaintext standalone file written
# in there would go up with the next upload, at a guessable URL, and reduce the
# access gate on data.json to decoration. Refuse rather than warn.
#
# Checked HERE, before the build, for the same reason Build-Dashboard.ps1 resolves
# its passphrase up front: a bad path should cost a second, not a full aggregation
# run. Re-checked after the build, once the default filename is known.
$docsFull = [IO.Path]::GetFullPath((Join-Path $root 'docs'))
function Assert-OutsideDocs {
    param([string] $Candidate)
    if ($Candidate.StartsWith($docsFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to write a plaintext standalone file inside docs/ - GitHub Pages publishes that folder, which would expose every figure at a public URL. Pick a path outside docs/."
    }
}
if ($OutFile) { Assert-OutsideDocs ([IO.Path]::GetFullPath($OutFile)) }

# Set-StrictMode makes a missing property a terminating error, and these fields
# are all absent from a data.json built before the version that added them.
function Get-Field {
    param($Obj, [string] $Name, $Default = $null)
    if ($Obj.PSObject.Properties.Name -contains $Name) { return $Obj.$Name }
    return $Default
}

# ---------------------------------------------------------------- build payload
# The plaintext payload goes to a TEMP path, never to docs/data.json. Letting
# -NoEncrypt fall through to its default output would overwrite the published
# ciphertext with plaintext, and the next upload would then serve every figure
# to anyone with the Pages URL. Passing the path explicitly is what prevents it.
$tmpJson = Join-Path ([IO.Path]::GetTempPath()) ("pa_standalone_{0}.json" -f [guid]::NewGuid().ToString('N'))

try {
    $buildArgs = @{
        OutFile          = $tmpJson
        NoEncrypt        = $true
        SkipSanitizedCsv = $true
    }
    if ($Source) { $buildArgs['Source'] = $Source }

    Write-Host "--- Build-Dashboard.ps1 (sanitization + publish gate + repo gate) ---" -ForegroundColor Cyan
    & $builder @buildArgs
    Write-Host "--- end of build ---" -ForegroundColor Cyan
    Write-Host ""

    if (-not (Test-Path -LiteralPath $tmpJson)) {
        throw "Build-Dashboard.ps1 did not produce $tmpJson. Nothing to inline."
    }

    $dataText = [IO.File]::ReadAllText($tmpJson)
    $payload  = $dataText | ConvertFrom-Json

    # -NoEncrypt is supposed to leave plaintext. If an envelope comes back, the
    # switch has stopped working and inlining the result would ship a file that
    # prompts for the very passphrase this mode exists to remove.
    if ($payload.PSObject.Properties.Name -contains 'enc') {
        throw "Payload is encrypted despite -NoEncrypt. Inlining it would produce a standalone file that still demands a passphrase."
    }

    # ------------------------------------------------------------ check template
    $html = [IO.File]::ReadAllText($template)

    # Without the inline-aware loader the injection still "succeeds", and the
    # result then ignores the block at runtime, fetches a data.json that is not
    # there, and shows the file:// error - a broken file that looks correctly
    # built. Caught here instead, loudly.
    if ($html -notmatch 'pa-inline-data') {
        throw @"
$template has no #pa-inline-data branch, so it cannot read an inlined payload.
It predates standalone mode, or was overwritten with an older copy. Restore the
inline-aware loadPayload() in docs/index.html before building standalone.
"@
    }
    if ($html -match '<script\s+id="pa-inline-data"') {
        throw "$template already holds an injected data block - it should be the clean template, not a previously built standalone file."
    }

    $tagCount = ([regex]::Matches($html, '<script')).Count
    if ($tagCount -ne 1) {
        throw "Expected exactly 1 <script> tag in the template but found $tagCount, so the injection point is ambiguous. Update this script to match the new structure."
    }

    # ------------------------------------------------------------------- inject
    # A '<' cannot occur outside a string value in compact JSON, so replacing it
    # wholesale keeps the text valid JSON while guaranteeing that no '</script>'
    # in the data can close the block early and spill the rest of the payload
    # into the page as markup. JSON.parse decodes the escape transparently, so
    # the browser still sees the original character.
    #
    # Built from a char code because a literal backslash-u sequence in this
    # source is too easy to mangle into a no-op replace, and an escape that
    # silently stops escaping is exactly the bug this line exists to prevent.
    $ltEscape = ([string][char]0x5C) + 'u003c'
    $escaped  = $dataText.Replace('<', $ltEscape)
    if ($dataText.Contains('<') -and -not $escaped.Contains($ltEscape)) {
        throw "The '<' escape did not apply - refusing to inline a payload that could break out of the script block."
    }

    $block = @(
        '<!-- ====================================================================',
        '     INLINED DASHBOARD DATA - PLAINTEXT AGGREGATES, NO PASSPHRASE.',
        '     Injected by Build-Standalone.ps1. The presence of this block is what',
        '     makes loadPayload() below read the payload from here instead of',
        '     fetching data.json, which is why this file needs no server.',
        '     Anyone holding this file can read every figure in it.',
        '     =================================================================== -->',
        ('<script id="pa-inline-data" type="application/json">' + $escaped + '</script>'),
        ''
    ) -join "`n"

    # IndexOf/Substring rather than .Replace: there is exactly one <script> tag
    # (asserted above) and this inserts before it without touching the tags the
    # injected block itself contains.
    $idx  = $html.IndexOf('<script>')
    $html = $html.Substring(0, $idx) + $block + $html.Substring($idx)

    # -------------------------------------------------------------- resolve out
    $stamp = Get-Field $payload 'dateMax' ((Get-Date).ToString('yyyy-MM-dd'))
    if (-not $OutFile) {
        $outDir = Join-Path $root 'standalone'
        if (-not (Test-Path -LiteralPath $outDir)) {
            New-Item -ItemType Directory -Path $outDir | Out-Null
        }
        $OutFile = Join-Path $outDir ("Prior_Auth_Dashboard_{0}.html" -f $stamp)
    } else {
        $parent = Split-Path -Parent $OutFile
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent | Out-Null
        }
    }

    # Re-checked here because the default filename is only known now, after the
    # build supplied dateMax. The early check above covers an explicit -OutFile.
    $outFull = [IO.Path]::GetFullPath($OutFile)
    Assert-OutsideDocs $outFull

    [IO.File]::WriteAllText($outFull, $html, (New-Object Text.UTF8Encoding $false))

    # --------------------------------------------------------------- report out
    Write-Host ("Template      : {0} (read only - unchanged)" -f $template)
    Write-Host ("Inlined       : {0:N1} KB of plaintext aggregates" -f ($dataText.Length / 1KB))
    Write-Host ("Data basis    : completions through {0}, source {1}" -f $stamp, (Get-Field $payload 'sourceFile' 'unknown'))
    Write-Host ("Wrote         : {0} ({1:N1} KB)" -f $outFull, ((Get-Item -LiteralPath $outFull).Length / 1KB)) -ForegroundColor Green
    Write-Host ""
    Write-Host "This file is PLAINTEXT and carries no passphrase prompt." -ForegroundColor Yellow
    Write-Host "  Anyone who opens it reads every figure. Share it only where access is already controlled." -ForegroundColor Yellow
    Write-Host "  Open it by double-clicking - it needs no server, unlike the hosted build." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "docs/data.json and docs/index.html were not modified; the GitHub build is unaffected." -ForegroundColor DarkGray
}
finally {
    # Aggregate-only and row-free, but there is no reason to leave a plaintext
    # copy lying around in %TEMP%.
    if (Test-Path -LiteralPath $tmpJson) { Remove-Item -LiteralPath $tmpJson -Force }
}
