<#
.SYNOPSIS
    Serves the dashboard folder over HTTP so you can preview it locally.

.DESCRIPTION
    index.html loads data.json with fetch(), which browsers block on file:// URLs.
    This starts a minimal static file server on localhost. Needs no admin rights
    because it binds to localhost only.

    Press Ctrl+C to stop.

.EXAMPLE
    .\Serve-Dashboard.ps1
    .\Serve-Dashboard.ps1 -Port 9000
#>
[CmdletBinding()]
param(
    [int]    $Port = 8080,
    [string] $Path,
    [switch] $NoLaunch
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
if (-not $Path) { $Path = Join-Path $root 'docs' }
$Path = (Resolve-Path -LiteralPath $Path).Path

if (-not (Test-Path (Join-Path $Path 'index.html'))) {
    throw "No index.html in $Path"
}
if (-not (Test-Path (Join-Path $Path 'data.json'))) {
    Write-Warning "data.json is missing in $Path - run .\Build-Dashboard.ps1 first."
}

$mime = @{
    '.html' = 'text/html; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.js'   = 'text/javascript; charset=utf-8'
    '.svg'  = 'image/svg+xml'
    '.png'  = 'image/png'
    '.ico'  = 'image/x-icon'
    '.md'   = 'text/markdown; charset=utf-8'
}

$listener = [Net.HttpListener]::new()
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)

try { $listener.Start() }
catch { throw "Could not bind $prefix - port may be in use. Try -Port 9000. ($($_.Exception.Message))" }

Write-Host "Serving $Path" -ForegroundColor Green
Write-Host "  -> $prefix" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop." -ForegroundColor DarkGray
if (-not $NoLaunch) { Start-Process $prefix }

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $res = $ctx.Response
        try {
            $rel = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath).TrimStart('/')
            if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'index.html' }

            # Resolve inside $Path only - refuse anything that escapes the root.
            $full = [IO.Path]::GetFullPath((Join-Path $Path $rel))
            if (-not $full.StartsWith($Path, [StringComparison]::OrdinalIgnoreCase)) {
                $res.StatusCode = 403
                $res.Close()
                continue
            }

            if (Test-Path -LiteralPath $full -PathType Leaf) {
                $bytes = [IO.File]::ReadAllBytes($full)
                $ext = [IO.Path]::GetExtension($full).ToLower()
                $res.ContentType = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { 'application/octet-stream' }
                # Always revalidate so a rebuilt data.json shows up on refresh.
                $res.Headers.Add('Cache-Control', 'no-store, must-revalidate')
                $res.ContentLength64 = $bytes.Length
                $res.OutputStream.Write($bytes, 0, $bytes.Length)
                Write-Host ("  200  " + $rel) -ForegroundColor DarkGray
            }
            else {
                $res.StatusCode = 404
                $msg = [Text.Encoding]::UTF8.GetBytes("404 - $rel not found")
                $res.ContentType = 'text/plain; charset=utf-8'
                $res.OutputStream.Write($msg, 0, $msg.Length)
                Write-Host ("  404  " + $rel) -ForegroundColor Yellow
            }
        }
        catch {
            Write-Warning $_.Exception.Message
            try { $res.StatusCode = 500 } catch {}
        }
        finally { try { $res.Close() } catch {} }
    }
}
finally {
    try { $listener.Stop(); $listener.Close() } catch {}
    Write-Host "`nServer stopped." -ForegroundColor DarkGray
}
