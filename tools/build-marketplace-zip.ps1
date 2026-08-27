$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root "dist app B24 zip"
$archiveBaseName = "Excel Tab B24"
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

$nextVersion = 14

$archive = Join-Path $dist "$archiveBaseName v.$nextVersion.zip"

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
