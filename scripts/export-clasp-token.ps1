Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Exports your clasp token JSON for GitHub Actions.
# Usage: .\scripts\export-clasp-token.ps1

$possible = @(
  Join-Path $HOME ".clasprc.json",
  Join-Path $HOME ".config\@google\clasp\config.json"
)

$tokenPath = $possible | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $tokenPath) {
  throw "Could not find clasp token file. Run 'clasp login' first. Checked: $($possible -join ', ')"
}

$raw = Get-Content $tokenPath -Raw
# Validate JSON
$null = $raw | ConvertFrom-Json

Write-Host "Found token file: $tokenPath"
Write-Host "Copy the JSON below into GitHub repo secret named CLASP_TOKEN."
Write-Host "---"
Write-Output $raw
