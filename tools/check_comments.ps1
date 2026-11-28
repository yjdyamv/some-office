param([string]$Path = "samples/ms/comments.doc")
# Read back a .doc: paragraphs, comment count, each comment's author initials and
# text, so a sample can be trusted before its bytes are measured.
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
  Write-Output ("paras=" + $doc.Paragraphs.Count + " comments=" + $doc.Comments.Count)
  foreach ($p in $doc.Paragraphs) {
    Write-Output ("  para: [" + $p.Range.Text.Trim() + "]")
  }
  $n = [int]$doc.Comments.Count
  for ($i = 1; $i -le $n; $i++) {
    $c = $doc.Comments.Item($i)
    $txt = ""
    try { $txt = [string]$c.Range.Text } catch {}
    $scope = ""
    try { $scope = [string]$c.Scope.Text } catch {}
    Write-Output ("  comment " + $i + ": author=" + $c.Author + " initials=" + $c.Initial + " scope=[" + $scope.Trim() + "] text=[" + $txt.Trim() + "]")
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
