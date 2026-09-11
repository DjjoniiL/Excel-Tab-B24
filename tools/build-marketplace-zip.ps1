$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root "dist app B24 zip"
$archiveBaseName = [Text.Encoding]::UTF8.GetString(
  [Convert]::FromBase64String("RXhjZWwg0YLQsNCx0LvQuNGG0LAg0LIg0YHQtNC10LvQutC1INC4INGN0LrRgdC/0L7RgNGCIE1WUCBGaW5hbA==")
)
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

$nextVersion = 34

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
