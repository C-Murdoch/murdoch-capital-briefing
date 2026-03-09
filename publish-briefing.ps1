param(
    [Parameter(Mandatory = $true)]
    [string]$BriefingFile,

    [string]$DateLabel = "",

    [switch]$SkipGit
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$indexPath = Join-Path $repoRoot "index.html"

if (-not (Test-Path $indexPath)) {
    throw "index.html not found at: $indexPath"
}

if (-not (Test-Path $BriefingFile)) {
    throw "Briefing file not found: $BriefingFile"
}

$briefingHtml = Get-Content -Raw -Path $BriefingFile
$indexHtml = Get-Content -Raw -Path $indexPath

$startMarker = "<!-- BRIEFING_CONTENT_START -->"
$endMarker = "<!-- BRIEFING_CONTENT_END -->"

if ($indexHtml -notmatch [regex]::Escape($startMarker) -or $indexHtml -notmatch [regex]::Escape($endMarker)) {
    throw "Could not find briefing markers in index.html"
}

$newBlock = @"
$startMarker
<div id="briefing-content">
$briefingHtml
</div>
$endMarker
"@

$pattern = "(?s)" + [regex]::Escape($startMarker) + ".*?" + [regex]::Escape($endMarker)
$updatedHtml = [regex]::Replace($indexHtml, $pattern, $newBlock)

Set-Content -Path $indexPath -Value $updatedHtml -NoNewline
Write-Host "Updated briefing content in index.html"

if ($SkipGit) {
    Write-Host "SkipGit enabled. No commit/push performed."
    exit 0
}

Push-Location $repoRoot
try {
    git add index.html

    if ([string]::IsNullOrWhiteSpace($DateLabel)) {
        $DateLabel = Get-Date -Format "yyyy-MM-dd"
    }

    git commit -m "Update daily briefing $DateLabel"
    git push origin main

    Write-Host "Pushed to GitHub. Railway will auto-deploy."
}
finally {
    Pop-Location
}
