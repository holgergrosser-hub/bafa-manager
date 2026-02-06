Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Open Apps Script project in browser
$id = (Get-Content "${PSScriptRoot}\..\.clasp.json" -Raw | ConvertFrom-Json).scriptId
$url = "https://script.google.com/d/$id/edit"
Write-Host $url
Start-Process $url
