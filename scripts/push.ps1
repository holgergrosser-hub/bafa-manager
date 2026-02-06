Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Push-Location "${PSScriptRoot}\.."
try {
  clasp push
} finally {
  Pop-Location
}
