param([string]$Path = "samples/tmp-bk.doc")
# Read the bookmarks back with real Word.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  $wd.DisplayAlerts = 0
  Write-Output ("app path: " + $wd.Path)
  if ([System.IO.Path]::IsPathRooted($Path)) { $full = $Path } else { $full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")) }
  $doc = $wd.Documents.Open($full, $false, $true)
  Write-Output ("paras=" + $doc.Paragraphs.Count + " bookmarks=" + $doc.Bookmarks.Count)
  for ($i = 1; $i -le $doc.Bookmarks.Count; $i++) {
    $b = $doc.Bookmarks.Item($i)
    $t = $b.Range.Text
    Write-Output ("  " + $b.Name + " = [" + $t + "] start=" + $b.Start + " end=" + $b.End)
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