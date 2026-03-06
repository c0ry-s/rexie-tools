# tools/bump-version.ps1
# Usage: pwsh ./tools/bump-version.ps1 -Version 1.0.2 [-NoTag]

[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [ValidatePattern('^\d+\.\d+\.\d+$')]
  [string]$Version,

  [switch]$NoTag
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = (git rev-parse --show-toplevel) 2>$null
if (-not $repoRoot) { throw "Not inside a git repository." }
Set-Location $repoRoot

$ps1Path = Join-Path $repoRoot 'RexieTools.ps1'
$verPath = Join-Path $repoRoot 'version.txt'

if (-not (Test-Path $ps1Path)) { throw "Missing RexieTools.ps1 at repo root." }
if (-not (Test-Path $verPath)) { throw "Missing version.txt at repo root." }

$ps1 = Get-Content $ps1Path -Raw

# 1) Update .VERSION block value (the line after ".VERSION")
$ps1 = [regex]::Replace(
  $ps1,
  '(\.VERSION\s*\r?\n\s*)(\d+\.\d+\.\d+)',
  "`$1$Version",
  'IgnoreCase'
)

# 2) Update $currentVersion = [version]"x.y.z"
$ps1 = [regex]::Replace(
  $ps1,
  '(\$currentVersion\s*=\s*\[version\]")(\d+\.\d+\.\d+)(")',
  "`$1$Version`$3",
  'IgnoreCase'
)

Set-Content -Path $ps1Path -Value $ps1 -Encoding UTF8
Set-Content -Path $verPath -Value ($Version + "`n") -Encoding UTF8

# Sanity: ensure both show the same version
$check = Select-String -Path $ps1Path -Pattern '\.VERSION|\$currentVersion' -Context 0,2
Write-Host $check

git add RexieTools.ps1 version.txt

$commitMsg = "Bump version to $Version"
git commit -m $commitMsg | Out-Null

if (-not $NoTag) {
  $tag = "v$Version"
  git tag -a $tag -m $tag
}

git push origin main
if (-not $NoTag) { git push origin ("v$Version") }

Write-Host "Done. Version bumped to $Version" -ForegroundColor Green