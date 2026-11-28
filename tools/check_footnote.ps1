param([string]$Path = "samples/ms/footnote.doc")
# Read back a .doc: paragraphs, footnote count and each footnote's text, so a
# sample can be trusted before its bytes are measured.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$exe = '"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" /Automation'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  # Word remembers every path it failed to open in `Resiliency\DisabledItems` and
  # then hangs on that path forever, even after the file is fixed. Clear the
  # list before each run so a repaired sample can be checked again.
  Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  $wd.DisplayAlerts = 0
  Write-Output ("app path: " + $wd.Path)
  if ([System.IO.Path]::IsPathRooted($Path)) { $full = $Path } else { $full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")) }
  $doc = $wd.Documents.Open($full, $false, $true)
  Write-Output ("paras=" + $doc.Paragraphs.Count + " footnotes=" + $doc.Footnotes.Count + " endnotes=" + $doc.Endnotes.Count)
  foreach ($p in $doc.Paragraphs) {
    $codes = @()
    foreach ($c in $p.Range.Characters) { $codes += [int][char]$c.Text }
    Write-Output ("  para: [" + $p.Range.Text.Trim() + "] codes=" + ($codes -join ","))
  }
  foreach ($f in $doc.Footnotes) {
    Write-Output ("  footnote " + $f.Index + ": [" + $f.Range.Text.Trim() + "] ref=" + $f.Reference.Text)
  }
  $doc.Close($false)
  $wd.Quit()
} catch {
  Write-Output ("ERR " + $_.Exception.Message)
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) {
    New-Item -Path $clsid -Force | Out-Null
    Set-ItemProperty -Path $clsid -Name '(default)' -Value $old
  }
}