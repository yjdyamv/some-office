param(
  [string]$Path = "samples/written-metadata.doc",
  [string]$Out = "samples/tmp-resaved.doc"
)
# Open a .doc in real Word and save it again. Anything Word did not understand
# in our property sets is dropped, so the re-saved file is the proof that the
# metadata was read correctly.
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
  if ([System.IO.Path]::IsPathRooted($Out)) { $target = $Out } else { $target = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Out")) }
  if (Test-Path $target) { Remove-Item $target -Force }
  $doc = $wd.Documents.Open($full, $false, $true)
  Write-Output ("opened paras=" + $doc.Paragraphs.Count)
  # 0 = wdFormatDocument (Word 97-2003), which keeps the property sets.
  $doc.SaveAs2($target, 0)
  $doc.Close($false)
  $wd.Quit()
  Write-Output ("resaved to " + $target)
} catch {
  Write-Output ("ERR " + $_.Exception.Message)
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) {
    New-Item -Path $clsid -Force | Out-Null
    Set-ItemProperty -Path $clsid -Name '(default)' -Value $old
  }
}