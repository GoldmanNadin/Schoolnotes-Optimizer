# Baut die Notizblock-App und installiert sie (mit Symbol auf dem Desktop).
# Rechtsklick → „Mit PowerShell ausführen“.  Braucht Node.js und Internet (beim ersten Mal).
# Gebaut wird außerhalb von OneDrive, damit nicht hunderte MB synchronisiert werden.
param([switch]$NichtInstallieren)
$ErrorActionPreference = "Stop"
$quelle = $PSScriptRoot
$bau = Join-Path $env:LOCALAPPDATA "Notizblock-Build"

Write-Host "1/3 Dateien vorbereiten ..."
New-Item -ItemType Directory -Force -Path $bau | Out-Null
foreach ($ordner in "desktop", "app", "build") {
  $ziel = Join-Path $bau $ordner
  if (Test-Path $ziel) { Remove-Item -Recurse -Force $ziel }
  Copy-Item -Recurse (Join-Path $quelle $ordner) $ziel
}
Copy-Item (Join-Path $quelle "package.json") $bau -Force

Push-Location $bau
try {
  Write-Host "2/3 Programmteile laden (beim ersten Mal ein paar Minuten) ..."
  npm install --no-audit --no-fund --silent
  if ($LASTEXITCODE -ne 0) { throw "npm install ist fehlgeschlagen" }
  Write-Host "3/3 App bauen ..."
  npx electron-builder --win nsis --x64 --publish never
  if ($LASTEXITCODE -ne 0) { throw "Bauen ist fehlgeschlagen" }
} finally { Pop-Location }

$setup = Get-ChildItem (Join-Path $bau "dist") -Filter "Notizblock-Setup-*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host ""
Write-Host "Fertig gebaut: $($setup.FullName)" -ForegroundColor Green
if (-not $NichtInstallieren) {
  Write-Host "Installiere ... (danach ist 'Notizblock' auf dem Desktop)"
  Start-Process $setup.FullName -ArgumentList "/S" -Wait
  Write-Host "Installiert!" -ForegroundColor Green
}
