param([string]$Path = "samples/ms/endnote.doc")
# Read back a .doc: paragraphs, footnote and endnote counts and each note's
# text, so a sample can be trusted before its bytes are measured.
#
# Notes are read by index, not by `foreach ($n in $doc.Endnotes)`: enumerating
# the collection hands back COM objects whose `Range` throws on some files our
# writer produces, while `Item(i)` on the same file reads fine.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
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
  Write-Output "step: opening"
  $doc = $wd.Documents.Open($full, $false, $true)
  Write-Output "step: opened"
  Write-Output ("paras=" + $doc.Paragraphs.Count + " footnotes=" + $doc.Footnotes.Count + " endnotes=" + $doc.Endnotes.Count)
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    Write-Output ("  para: [" + $doc.Paragraphs.Item($i).Range.Text.Trim() + "]")
  }
  for ($i = 1; $i -le $doc.Footnotes.Count; $i++) {
    $n = $doc.Footnotes.Item($i)
    $ref = ''
    if ($n.Reference) { $ref = [string]$n.Reference.Text }
    Write-Output ("  footnote " + $i + ": [" + $n.Range.Text + "] ref=[" + $ref + "]")
  }
  for ($i = 1; $i -le $doc.Endnotes.Count; $i++) {
    $n = $doc.Endnotes.Item($i)
    $ref = ''
    if ($n.Reference) { $ref = [string]$n.Reference.Text }
    Write-Output ("  endnote " + $i + ": [" + $n.Range.Text + "] ref=[" + $ref + "]")
  }
  $doc.Close($false)
  $wd.Quit()
} catch {
  Write-Output ("ERR at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message)
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) {
    New-Item -Path $clsid -Force | Out-Null
    Set-ItemProperty -Path $clsid -Name '(default)' -Value $old
  }
}