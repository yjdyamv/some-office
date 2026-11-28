param([string]$Path = "samples/written-metadata.doc")
# Read back a .doc's File > Properties metadata through Word itself: the
# BuiltinDocumentProperties collection is what the Summary tab shows.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  $wd.DisplayAlerts = 0
  Write-Output ("app path: " + $wd.Path)
  if ([System.IO.Path]::IsPathRooted($Path)) { $full = $Path } else { $full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")) }
  $doc = $wd.Documents.Open($full, $false, $true)
  Write-Output ("paras=" + $doc.Paragraphs.Count + " text=[" + $doc.Content.Text.Trim() + "]")
  # Property names are localised (a Chinese Word calls the title 「标题」), so
  # walk the collection by index and print each entry's own Name.
  $props = $doc.BuiltInDocumentProperties
  Write-Output ("builtin properties=" + $props.Count)
  for ($i = 1; $i -le $props.Count; $i++) {
    $p = $props.Item($i)
    $value = ""
    try { $value = [string]$p.Value } catch { $value = "<unreadable>" }
    Write-Output ("  [" + $p.Name + "] = [" + $value + "]")
  }
  $custom = $doc.CustomDocumentProperties
  Write-Output ("custom properties=" + $custom.Count)
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
