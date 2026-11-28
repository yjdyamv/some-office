param([string]$Path = "samples/written-lists.doc")
# Read back a .doc with lists:每段的 ListFormat 编号/项目符号、ListString 与层级。
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$exe = '"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" /Automation'
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
  Write-Output ("paras=" + $doc.Paragraphs.Count + " listtemplates=" + $doc.ListTemplates.Count)
  $i = 0
  foreach ($p in $doc.Paragraphs) {
    $i++
    $lf = $p.Range.ListFormat
    $ls = ""
    try { $ls = $lf.ListString } catch {}
    Write-Output ("  para " + $i + ": type=" + $lf.ListType + " level=" + $lf.ListLevelNumber + " listString=[" + $ls + "] text=[" + $p.Range.Text.Trim() + "]")
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
