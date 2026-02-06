Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Push-Location "${PSScriptRoot}\.."
try {
  clasp pull
} finally {
  Pop-Location
}
