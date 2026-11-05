# build-pdf.ps1 — regenerate docs/flowchart.html and export docs/flowchart.pdf.
#
# Why two steps: printing the *dynamic* page lets Chromium paginate while Mermaid
# is still re-laying out, which produces duplicated/blank pages. So we first let
# Edge render the page (JS runs, diagrams become inline SVG), dump the finished
# DOM to a static HTML, and print THAT — stable pagination, no async surprises.
#
# Edge is a GUI-subsystem process: plain `&` invocation does not reliably wait for
# it, so every call here uses Start-Process -Wait with redirected output.
#
# Requires Microsoft Edge (present on Windows 10/11). Print rules live in the
# generated HTML: A4 landscape, diagrams kept on one page where they fit.
#
# Usage: powershell -ExecutionPolicy Bypass -File docs/build-pdf.ps1

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

function Find-Edge {
    @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Invoke-Edge {
    param(
        [string]$EdgePath,
        [string[]]$Arguments,
        [string]$StdOut,
        [string]$StdErr,
        [string]$WorkDir
    )
    $p = Start-Process -FilePath $EdgePath -ArgumentList $Arguments -Wait -PassThru `
        -NoNewWindow -RedirectStandardOutput $StdOut -RedirectStandardError $StdErr `
        -WorkingDirectory $WorkDir
    return $p.ExitCode
}

Write-Host "[1/4] Regenerating flowchart.html ..."
py -3 (Join-Path $here "build-flowchart-html.py")
if ($LASTEXITCODE -ne 0) { throw "build-flowchart-html.py failed" }

$edge = Find-Edge
if (-not $edge) { throw "Microsoft Edge not found; install it or print the HTML manually." }

$dynamic = Join-Path $here "flowchart.html"
$static = Join-Path $here "flowchart.print.html"
$pdf = Join-Path $here "flowchart.pdf"
$dynamicUri = ([uri]$dynamic).AbsoluteUri

$job = Join-Path $env:TEMP ("em_pdf_job_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $job -Force | Out-Null
$domFile = Join-Path $job "dom.html"
$errFile = Join-Path $job "edge.log"

try {
    Write-Host "[2/4] Rendering the page and freezing it to static HTML ..."
    $profile1 = Join-Path $job "profile1"
    $args1 = @(
        "--headless=new", "--disable-gpu", "--no-first-run",
        "--user-data-dir=$profile1",
        "--virtual-time-budget=30000",
        "--dump-dom", $dynamicUri
    )
    $code = Invoke-Edge -EdgePath $edge -Arguments $args1 -StdOut $domFile -StdErr $errFile -WorkDir $job
    if (-not (Test-Path $domFile)) { throw "Edge produced no DOM (exit $code). See $errFile" }
    $html = [IO.File]::ReadAllText($domFile)
    if (-not $html -or $html.Length -lt 5000) {
        throw "Frozen DOM looks empty ($($html.Length) chars). See $errFile"
    }
    # Drop the scripts: everything is already rendered inline. Leaving them in
    # would re-run marked/mermaid inside the print job.
    $html = [regex]::Replace($html, '(?s)<script\b.*?</script>', '')
    [IO.File]::WriteAllText($static, $html, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ("       frozen DOM: {0:N0} KB" -f ($html.Length / 1KB))

    Write-Host "[3/4] Printing the static page to PDF ..."
    $profile2 = Join-Path $job "profile2"
    $staticUri = ([uri]$static).AbsoluteUri
    $args2 = @(
        "--headless=new", "--disable-gpu", "--no-first-run",
        "--user-data-dir=$profile2",
        "--virtual-time-budget=5000",
        "--no-pdf-header-footer",
        "--print-to-pdf=`"$pdf`"",
        $staticUri
    )
    $code2 = Invoke-Edge -EdgePath $edge -Arguments $args2 -StdOut (Join-Path $job "print.out") -StdErr $errFile -WorkDir $job
    if (-not (Test-Path $pdf)) { throw "PDF was not created (exit $code2). See $errFile" }
    Write-Host ("[4/4] Done: {0} ({1:N0} KB)" -f $pdf, ((Get-Item $pdf).Length / 1KB))
}
finally {
    if (Test-Path $static) { Remove-Item $static -Force -ErrorAction SilentlyContinue }
    if (Test-Path $job) { Remove-Item $job -Recurse -Force -ErrorAction SilentlyContinue }
}
