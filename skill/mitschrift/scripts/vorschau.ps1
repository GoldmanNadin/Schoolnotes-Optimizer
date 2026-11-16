# Wandelt ein Word-Dokument mit dem installierten Microsoft Word in ein PDF um (zur Kontrolle).
#   powershell -ExecutionPolicy Bypass -File vorschau.ps1 <datei.docx> <ziel.pdf>
param([Parameter(Mandatory=$true)][string]$Docx, [Parameter(Mandatory=$true)][string]$Pdf)
$ErrorActionPreference = "Stop"
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open((Resolve-Path $Docx).Path, $false, $true)   # schreibgeschützt öffnen
  $doc.ExportAsFixedFormat([System.IO.Path]::GetFullPath($Pdf), 17)       # 17 = PDF
  $doc.Close($false)
  Write-Output "OK: $Pdf"
} catch {
  Write-Output "FEHLER: $($_.Exception.Message)"
  exit 1
} finally {
  if ($word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
