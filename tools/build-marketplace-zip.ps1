$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root "dist app B24 zip"
$runtimeFiles = @(
  "install.html",
  "install.js",
  "install.css",
  "index.html",
  "app.js",
  "style.css"
)

if (-not (Test-Path $dist)) {
  New-Item -ItemType Directory -Path $dist | Out-Null
}

$version = Get-Date -Format "yyyyMMdd-HHmmss"
$archive = Join-Path $dist "excel-tab-b24-marketplace-$version.zip"

if (Test-Path $archive) {
  throw "Archive already exists: $archive"
}

$missing = $runtimeFiles | Where-Object { -not (Test-Path (Join-Path $root $_)) }
if ($missing.Count -gt 0) {
  throw "Missing runtime files: $($missing -join ', ')"
}

$paths = $runtimeFiles | ForEach-Object { Join-Path $root $_ }
Compress-Archive -Path $paths -DestinationPath $archive -CompressionLevel Optimal
Write-Host "Created $archive"
