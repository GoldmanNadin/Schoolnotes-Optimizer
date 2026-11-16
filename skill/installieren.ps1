# Installiert (oder aktualisiert) den Skill „mitschrift“ für Claude.
# Rechtsklick auf diese Datei → „Mit PowerShell ausführen“.
$ErrorActionPreference = "Stop"
$quelle = Join-Path $PSScriptRoot "mitschrift"
$ziel   = Join-Path $env:USERPROFILE ".claude\skills\mitschrift"

Write-Host "Installiere den Skill 'mitschrift' nach $ziel ..."
New-Item -ItemType Directory -Force -Path $ziel | Out-Null
# Dateien kopieren (vorhandene node_modules bleiben erhalten)
Get-ChildItem $quelle -Recurse -File | Where-Object { $_.FullName -notmatch "\\node_modules\\" } | ForEach-Object {
  $rel = $_.FullName.Substring($quelle.Length + 1)
  $dest = Join-Path $ziel $rel
  New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
  Copy-Item $_.FullName $dest -Force
}
Write-Host "Lade die Word-Bibliothek (einmalig, braucht Internet) ..."
Push-Location (Join-Path $ziel "scripts")
npm install --no-audit --no-fund --silent
Pop-Location
Write-Host ""
Write-Host "Fertig! In der Claude-App kannst du jetzt /mitschrift verwenden." -ForegroundColor Green
Start-Sleep -Seconds 3
